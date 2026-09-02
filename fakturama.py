"""Driving the Fakturama desktop app through Microsoft UI Automation.

The only module that talks to the running application. It speaks in terms of
the UI ("open a New Order", "set the Date field") and takes plain values, so
the extraction side (`sales_order`) never has to know a window exists.

Fakturama is an Eclipse RCP/SWT application: its controls carry no automation
ids, so everything here is located by control type plus visible name. Those
names are the English UI strings -- a differently-localised install needs its
own map.

We drive controls through their UIA patterns (Invoke, Value, ...) rather than
by clicking screen coordinates: a coordinate click depends on the window being
foreground, unobscured and unmoved, and silently hits the neighbouring toolbar
button when it is not.
"""

import time
from datetime import date, datetime
from typing import Any

import tracing
from pywinauto import Application
from pywinauto.application import WindowSpecification
from pywinauto.keyboard import send_keys


#: Main window title; the rest of the caption is the data directory path.
APP_TITLE_RE = r"^Fakturama.*"

#: Toolbar button that opens a blank order editor (step 1.3).
NEW_ORDER_BUTTON = "Create: New Order"

#: Editors appear as tabs in the main window; a fresh order's tab is called
#: this until the order is saved and takes its number. The leading star is
#: Eclipse's unsaved marker -- it appears the moment we type into a field, so
#: any lookup by tab name has to tolerate it.
NEW_ORDER_TAB_RE = r"^\*?New Order$"

#: Label of the external-reference box. The trailing dot is part of the UI
#: string ("Cust.Ref."), and the box carries it as its accessible name.
CUSTOMER_REFERENCE_LABEL = "Cust.Ref."

#: The price-mode combo in the editor's header row carries no accessible name
#: (the VAT and Shipping combos below it do), so we identify it by being the
#: unnamed one -- and confirm by its current value being one of these.
PRICE_MODES = ("Net", "Gross")

#: Price mode the imported order is entered in (step 1.7).
NET_PRICE_MODE = "Net"

#: VAT mode step 1.7 wants left in place. Fakturama's default for a new order,
#: so normally nothing to do -- but we assert it rather than assume it.
WITH_VAT = "With VAT"

#: How Fakturama renders dates in this (English) install: "Jul 14, 2026".
DATE_DISPLAY_FORMAT = "%b %d, %Y"

#: Order of the date widget's editable fields, matching DATE_DISPLAY_FORMAT.
#: Typing digits fills the focused field and advances to the next, so this is
#: the order we feed them in. `set_date` reads the result back, so a locale
#: with a different field order fails loudly instead of writing a wrong date.
DATE_FIELD_ORDER = ("month", "day", "year")

#: Vertical slack, in pixels, for deciding that a label and an input sit on
#: the same line of the form.
LINE_TOLERANCE = 15

#: Seconds to wait for the app to build an editor. Generous: the first editor
#: of a session also warms up the form toolkit.
EDITOR_TIMEOUT = 30


def connect() -> WindowSpecification:
    """Attach to the already-running Fakturama main window.

    Returns:
        WindowSpecification: The main window, focused.

    Raises:
        RuntimeError: If no Fakturama window is open.
    """
    try:
        app = Application(backend="uia").connect(title_re=APP_TITLE_RE, timeout=10)
    except Exception as exc:  # pywinauto raises several unrelated types here
        raise RuntimeError(
            "No running Fakturama window found. Start Fakturama first; this "
            "module attaches to a running instance rather than launching one."
        ) from exc
    win = app.window(title_re=APP_TITLE_RE)
    win.wait("visible ready", timeout=EDITOR_TIMEOUT)
    win.set_focus()
    return win


def open_new_order(win: WindowSpecification) -> Any:
    """Click Order in the top toolbar and wait for the New Order editor (step 1.3).

    Args:
        win (WindowSpecification): The main window from `connect`.

    Returns:
        Any: The editor pane -- the tab body holding the order's fields, to
            scope later field lookups to this editor rather than to any other
            one the user has open.
    """
    with tracing.step(f"click {NEW_ORDER_BUTTON!r} in the toolbar"):
        button = win.child_window(title=NEW_ORDER_BUTTON, control_type="Button").wrapper_object()
        tracing.point_at(button)
        button.iface_invoke.Invoke()
    with tracing.step("wait for the New Order editor"):
        editor = wait_for_order_editor(win)
    tracing.point_at(editor, colour=tracing.CONFIRM)
    return editor


def wait_for_order_editor(win: WindowSpecification, timeout: int = EDITOR_TIMEOUT) -> Any:
    """Wait until the New Order editor is the selected tab, and return its body.

    The editor is a tab inside the main window, not a separate top-level
    window, so there is no window handle to wait on. We wait for the tab body
    instead: Fakturama names it after whichever tab is selected, and only the
    selected editor's controls are in the tree, so finding it by name both
    confirms the editor opened and gives us a container scoped to it.

    Args:
        win (WindowSpecification): The main window.
        timeout (int): Seconds to wait before giving up.

    Returns:
        Any: The editor pane.

    Raises:
        TimeoutError: If the editor does not appear in time.
    """
    deadline = time.monotonic() + timeout
    while True:
        editor = win.child_window(title_re=NEW_ORDER_TAB_RE, control_type="Tab")
        if editor.exists():
            return editor.wrapper_object()
        if time.monotonic() >= deadline:
            raise TimeoutError(
                f"No New Order editor appeared within {timeout}s after "
                f"invoking {NEW_ORDER_BUTTON!r}."
            )
        time.sleep(0.25)


def _field(editor: Any, label: str) -> Any:
    """Find the input belonging to `label`, by name or by position (step 1.5 on).

    Some inputs carry their label as their UIA name and can be found directly.
    The date widget and the No. box carry no name at all, so for those we take
    the labelled static text and pick the nearest input to its right on the
    same line -- how the form reads visually.

    Args:
        editor (Any): The editor pane from `open_new_order`.
        label (str): The visible label text, e.g. "Date".

    Returns:
        Any: The input control.

    Raises:
        LookupError: If the label, or an input beside it, is not there.
    """
    named = [c for c in editor.descendants(control_type="Edit") if c.element_info.name == label]
    if named:
        return named[0]

    statics = [c for c in editor.descendants(control_type="Text") if c.element_info.name == label]
    if not statics:
        raise LookupError(f"No {label!r} label in the order editor.")
    anchor = statics[0].rectangle()
    same_line = [
        c
        for c in editor.descendants(control_type="Edit")
        if c.rectangle().left >= anchor.right
        and abs(c.rectangle().mid_point().y - anchor.mid_point().y) <= LINE_TOLERANCE
    ]
    if not same_line:
        raise LookupError(f"Found the {label!r} label but no input beside it.")
    return min(same_line, key=lambda c: c.rectangle().left)


def set_date(editor: Any, order_date: str | date) -> None:
    """Set the order's Date to the extracted order date (step 1.5).

    Date is an SWT date widget, not a plain text box: writing its value
    through UIA looks like it worked but is discarded as soon as focus moves
    (verified -- the field snaps back to today). It does accept typed digits,
    filling one date field at a time and advancing, so that is what we do.

    Args:
        editor (Any): The editor pane from `open_new_order`.
        order_date (str | date): The date, as a `date` or the ISO
            "YYYY-MM-DD" string that `SalesOrder.order_date` holds.

    Raises:
        ValueError: If the field ends up showing a different date than asked
            for -- e.g. an install whose date fields run day-first.
    """
    wanted = date.fromisoformat(order_date) if isinstance(order_date, str) else order_date
    digits = {"day": f"{wanted.day:02d}", "month": f"{wanted.month:02d}", "year": f"{wanted.year:04d}"}

    with tracing.step(f"set Date to {wanted.isoformat()}"):
        field = _field(editor, "Date")
        tracing.point_at(field)
        field.set_focus()
        send_keys("{HOME}" + "".join(digits[name] for name in DATE_FIELD_ORDER))

    shown = _field(editor, "Date").get_value()
    try:
        got = datetime.strptime(shown, DATE_DISPLAY_FORMAT).date()
    except ValueError as exc:
        raise ValueError(f"Could not read the Date field back; it shows {shown!r}.") from exc
    if got != wanted:
        raise ValueError(
            f"Date field shows {shown!r} after entering {wanted.isoformat()}. "
            f"This install's date fields are probably not ordered "
            f"{'/'.join(DATE_FIELD_ORDER)}."
        )
    tracing.point_at(_field(editor, "Date"), colour=tracing.CONFIRM)


def set_customer_reference(editor: Any, reference: str) -> None:
    """Enter the extracted external reference in Cust.Ref. (step 1.6).

    Unlike Date, this is a plain SWT text box and honours the UIA Value
    pattern -- the value survives losing focus and marks the editor dirty --
    so we set it directly rather than typing, which keeps the reference exact
    (no keyboard layout or dead-key surprises with characters like '-').

    Args:
        editor (Any): The editor pane from `open_new_order`.
        reference (str): The order's external reference, e.g.
            "WEB-2026-0714-A17" (`SalesOrder.external_reference`).

    Raises:
        ValueError: If the field does not hold the reference afterwards.
    """
    with tracing.step(f"enter {reference!r} in Cust.Ref."):
        field = _field(editor, CUSTOMER_REFERENCE_LABEL)
        tracing.point_at(field)
        field.iface_value.SetValue(reference)

    shown = _field(editor, CUSTOMER_REFERENCE_LABEL).get_value()
    if shown != reference:
        raise ValueError(f"Cust.Ref. holds {shown!r} after entering {reference!r}.")
    tracing.point_at(_field(editor, CUSTOMER_REFERENCE_LABEL), colour=tracing.CONFIRM)


def _price_mode_combo(editor: Any) -> Any:
    """Find the header row's price-mode combo (Net/Gross).

    Args:
        editor (Any): The editor pane from `open_new_order`.

    Returns:
        Any: The combo box.

    Raises:
        LookupError: If no unnamed combo showing a known price mode is there.
    """
    for combo in editor.descendants(control_type="ComboBox"):
        if not combo.element_info.name and combo.selected_text() in PRICE_MODES:
            return combo
    raise LookupError(
        f"No price-mode combo (showing one of {'/'.join(PRICE_MODES)}) in the order editor."
    )


def set_price_mode(editor: Any, mode: str = NET_PRICE_MODE) -> None:
    """Set the document's price mode, leaving VAT alone (step 1.7).

    Args:
        editor (Any): The editor pane from `open_new_order`.
        mode (str): "Net" or "Gross"; the import uses Net, since the extracted
            line items carry net unit prices with VAT stated separately.

    Raises:
        ValueError: If the combo does not end up on `mode`.
    """
    combo = _price_mode_combo(editor)
    if combo.selected_text() != mode:
        with tracing.step(f"set the price mode to {mode}"):
            tracing.point_at(combo)
            combo.select(mode)

    shown = _price_mode_combo(editor).selected_text()
    if shown != mode:
        raise ValueError(f"Price mode is {shown!r} after selecting {mode!r}.")
    tracing.point_at(_price_mode_combo(editor), colour=tracing.CONFIRM)


def set_vat_mode(editor: Any, mode: str = WITH_VAT) -> None:
    """Keep VAT as With VAT (step 1.7).

    A new order already opens on this setting, so this is normally a no-op --
    it exists to make the requirement explicit and to catch an install (or a
    later step) that leaves the editor on a different VAT mode.

    Args:
        editor (Any): The editor pane from `open_new_order`.
        mode (str): The VAT mode to end up on.

    Raises:
        ValueError: If the combo does not end up on `mode`.
    """
    combo = _field_combo(editor, "VAT")
    if combo.selected_text() != mode:
        with tracing.step(f"set VAT to {mode}"):
            tracing.point_at(combo)
            combo.select(mode)

    shown = _field_combo(editor, "VAT").selected_text()
    if shown != mode:
        raise ValueError(f"VAT is {shown!r} after selecting {mode!r}.")
    tracing.point_at(_field_combo(editor, "VAT"), colour=tracing.CONFIRM)


def _field_combo(editor: Any, label: str) -> Any:
    """Find the combo box whose accessible name is `label`.

    Args:
        editor (Any): The editor pane from `open_new_order`.
        label (str): The combo's name, e.g. "VAT".

    Returns:
        Any: The combo box.

    Raises:
        LookupError: If there is no such combo.
    """
    for combo in editor.descendants(control_type="ComboBox"):
        if combo.element_info.name == label:
            return combo
    raise LookupError(f"No {label!r} combo in the order editor.")


if __name__ == "__main__":
    tracing.configure(visual=True)
    try:
        window = connect()
        print("connected:", window.element_info.name)
        order_editor = open_new_order(window)
        print("editor:", order_editor.element_info.name)
        set_date(order_editor, "2026-07-14")
        print("date:", _field(order_editor, "Date").get_value())
        set_customer_reference(order_editor, "WEB-2026-0714-A17")
        print("cust.ref:", _field(order_editor, CUSTOMER_REFERENCE_LABEL).get_value())
        set_price_mode(order_editor)
        set_vat_mode(order_editor)
        print("price mode:", _price_mode_combo(order_editor).selected_text())
        print("vat:", _field_combo(order_editor, "VAT").selected_text())
    finally:
        tracing.stop()
