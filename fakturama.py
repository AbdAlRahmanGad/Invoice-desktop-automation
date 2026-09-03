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

import ctypes
import re
import time
from datetime import date, datetime
from typing import Any

import table
import tracing
import win32api
import win32con
import win32gui
import win32process
import win32ui
from PIL import Image
from pywinauto import Application
from pywinauto.application import WindowSpecification
from pywinauto.keyboard import send_keys


#: Everything that is not a letter or a digit, for comparing an address the
#: app composed against the one the document printed.
_NON_ALNUM_RE = re.compile(r"[^0-9A-Za-z]+")

#: Main window title; the rest of the caption is the data directory path.
APP_TITLE_RE = r"^Fakturama.*"

#: Toolbar button that opens a blank order editor (step 1.3).
NEW_ORDER_BUTTON = "Create: New Order"

#: Editors appear as tabs in the main window; a fresh order's tab is called
#: this until the order is saved and takes its number. The leading star is
#: Eclipse's unsaved marker -- it appears the moment we type into a field, so
#: any lookup by tab name has to tolerate it.
NEW_ORDER_TAB_RE = r"^\*?New Order$"

#: The Addresses row's two icons: the upper one picks an existing contact,
#: the lower green + starts a new one. We only ever want the upper (step 2.1),
#: so the lookup insists on finding exactly this many and takes the topmost.
ADDRESS_ICON_COUNT = 2

#: Modal that the upper icon opens.
ADDRESS_DIALOG_TITLE = "Select the address"

#: Seconds to wait for a modal dialog to appear.
DIALOG_TIMEOUT = 10

#: How many times to click a control that should open a dialog, and how long
#: to give it each time.
DIALOG_ATTEMPTS = 4
DIALOG_RETRY_SECONDS = 3

#: The columns step 2.3 compares, and which extracted value belongs in each.
#: The address list shows a No. column too, which is deliberately not compared:
#: it is Fakturama's own numbering, not something the document knows, and it
#: is the one column the recogniser reads unreliably (O for zero).
DEBTOR_MATCH_COLUMNS = {
    "company": "Company",
    "first_name": "First Name",
    "last_name": "Name",
    "postcode": "ZIP",
    "city": "City",
}
COMPANY_COLUMN = "Company"

#: How a cell that is too narrow for its value ends, as the recogniser writes
#: it. Matched on the visible part in that case.
TRUNCATION_MARKS = ("...", "..", ".", "\u2026")

#: Where along a row to click when selecting it, as a fraction of the list's
#: width. Anywhere in the row will do; this keeps clear of the narrow marker
#: column on the left and of the horizontal scrollbar's reach on the right.
ROW_CLICK_X = 0.15

#: Seconds to wait for a list to show that a row was selected.
SELECTION_TIMEOUT = 5

#: Label beside the dialog's search box.
SEARCH_LABEL = "Search:"

#: How the result list is watched for "stopped changing" (step 2.2): grab the
#: list this often, and call it settled after this many identical grabs in a
#: row. Two is enough because the filter redraws in one go; the interval is
#: what actually decides how long we wait after the last redraw.
RESULTS_POLL_SECONDS = 0.2
RESULTS_STABLE_FRAMES = 2

#: Seconds to wait for the list to settle before giving up on it.
RESULTS_TIMEOUT = 15

#: How many times to write the search term before calling it a failure.
SEARCH_ATTEMPTS = 5

#: PrintWindow flag: render the whole window, including parts covered by
#: other windows, so a capture works with Fakturama in the background.
PW_RENDERFULLCONTENT = 2

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

#: Toolbar button that saves the editor in front (steps 2.10.6, 2.11).
SAVE_BUTTON = "Save the current contents"

#: Eclipse's unsaved marker, prefixed to an editor tab's name.
DIRTY_MARKER = "*"

#: Creating a payment method (steps 2.10.3-2.10.5): the view's green +, the
#: editor it opens, and the fields that must end up zero or blank.
NEW_TERM_BUTTON = "Create a new term of payment"
NEW_TERM_TAB_RE = r"^\*?New Term of Payment$"
CASH_DISCOUNT_LABEL = "Cash discount"
DAY_LABELS = ("Discount Days", "Net Days")
BLANK_TEXT_LABELS = ("Text 'unpaid'", "Text 'deposit'", "Text 'paid'")
NO_CASH_DISCOUNT = "0%"
ZERO_DAYS = "0"

#: The payment-code combo. Its label is an untranslated message key in this
#: build rather than a caption, which is what the editor shows on screen too.
PAYMENT_CODE_LABEL = "!editorPaymentPaymentcode!"

#: Step 2.10.4's mapping from the document's wording to Fakturama's code.
#: Anything not listed here is not something this import knows how to file,
#: and `create_payment_method` refuses rather than guessing a code.
PAYMENT_CODES = {
    "Bank Transfer": "Credit transfer",
    "Credit Card": "Credit card",
    "SEPA Direct Debit": "SEPA direct debit",
}

#: The Data list holding payment methods (step 2.10.1).
TERMS_OF_PAYMENT_VIEW = "terms of payment"

#: The debtor's payment method (step 2.10). This install keeps it on the
#: Miscellaneous tab rather than a Payment tab of its own.
PAYMENT_LABEL = "Payment"

#: Seconds to wait for a dropped-down combo to fill in its options, and how
#: many times to ask it to select one before giving up.
COMBO_TIMEOUT = 5
COMBO_ATTEMPTS = 4

#: The debtor's Miscellaneous tab and the three fields step 2.9 sets there.
#: Discount is a percentage box that already reads 0% on a new debtor.
MISCELLANEOUS_TAB = "Miscellaneous"
ALIAS_LABEL = "Alias name"
DISCOUNT_LABEL = "Discount"
NET_GROSS_LABEL = "Net or Gross"
NO_DISCOUNT = "0%"

#: The "address type" row holds the address's roles, behind a button that
#: opens a two-checkbox popup (step 2.8).
ADDRESS_TYPE_LABEL = "address type"
ROLE_INVOICE = "Invoice address"
ROLE_DELIVERY = "Delivery address"

#: Window class of Fakturama's dialogs and popups.
DIALOG_CLASS = "#32770"

#: Seconds to wait for a tab to come to the front.
TAB_TIMEOUT = 5

#: Tabs of the debtor editor that hold the billing address (step 2.7).
ADDRESSES_TAB = "Addresses"
MAIN_ADDRESS_TAB = "Main address"

#: Postcode and city share one label, over two unnamed boxes.
POSTCODE_CITY_LABEL = "ZIP - City"

#: The debtor editor labels First Name and Last Name once, over two unnamed
#: boxes, so both are found by position under this one label.
NAME_LABEL = "First Name Last Name"

#: What the Salutation combo shows when no salutation is chosen (step 2.6).
NO_SALUTATION = "---"

#: Left-hand New panel entry that opens a debtor editor (step 2.5), and the
#: tab that entry produces. Fakturama calls the command "New Contact" and the
#: editor "New Debtor".
NEW_CONTACT_LINK = "New Contact"
NEW_DEBTOR_TAB_RE = r"^\*?New Debtor$"

#: How Fakturama renders dates in this (English) install: "Jul 14, 2026".
DATE_DISPLAY_FORMAT = "%b %d, %Y"

#: Order of the date widget's editable fields, matching DATE_DISPLAY_FORMAT.
#: Typing digits fills the focused field and advances to the next, so this is
#: the order we feed them in. `set_date` reads the result back, so a locale
#: with a different field order fails loudly instead of writing a wrong date.
DATE_FIELD_ORDER = ("month", "day", "year")

#: Seconds to wait for typed characters to land in a text box.
TYPING_TIMEOUT = 5

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
    # "visible", not "ready": a modal dialog disables the main window, and a
    # run that has to reopen the address selector legitimately attaches while
    # one is open.
    win.wait("visible", timeout=EDITOR_TIMEOUT)
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


def wait_for_editor(win: WindowSpecification, title_re: str, what: str, timeout: int = EDITOR_TIMEOUT) -> Any:
    """Wait until an editor is the selected tab, and return its body.

    Editors are tabs inside the main window, not separate top-level windows,
    so there is no window handle to wait on. We wait for the tab body instead:
    Fakturama names it after whichever tab is selected, and only the selected
    editor's controls are in the tree, so finding it by name both confirms the
    editor opened and gives us a container scoped to it.

    Args:
        win (WindowSpecification): The main window.
        title_re (str): Pattern the tab body's name must match.
        what (str): The editor's name, for the error message.
        timeout (int): Seconds to wait before giving up.

    Returns:
        Any: The editor pane.

    Raises:
        TimeoutError: If the editor does not appear in time.
    """
    deadline = time.monotonic() + timeout
    while True:
        editor = win.child_window(title_re=title_re, control_type="Tab")
        if editor.exists():
            return editor.wrapper_object()
        if time.monotonic() >= deadline:
            raise TimeoutError(f"No {what} editor appeared within {timeout}s.")
        time.sleep(0.25)


def wait_for_order_editor(win: WindowSpecification, timeout: int = EDITOR_TIMEOUT) -> Any:
    """Wait for the New Order editor and return its body.

    Args:
        win (WindowSpecification): The main window.
        timeout (int): Seconds to wait before giving up.

    Returns:
        Any: The editor pane.
    """
    return wait_for_editor(win, NEW_ORDER_TAB_RE, "New Order", timeout)


def _fields(editor: Any, label: str, control_type: str = "Edit") -> list[Any]:
    """Find the inputs belonging to `label`, left to right.

    Some inputs carry their label as their UIA name and can be found directly.
    Others carry no name at all -- the date widget, the No. box, and the pairs
    that share one label ("First Name Last Name", "ZIP - City") -- so for
    those we take the labelled static text and collect the inputs to its right
    on the same line, in the order they read.

    Args:
        editor (Any): An editor pane.
        label (str): The visible label text, e.g. "Date".
        control_type (str): Which kind of input to look for.

    Returns:
        list[Any]: The inputs, leftmost first.

    Raises:
        LookupError: If the label, or an input beside it, is not there.
    """
    named = [c for c in editor.descendants(control_type=control_type) if c.element_info.name == label]
    if named:
        return named

    statics = [c for c in editor.descendants(control_type="Text") if c.element_info.name == label]
    if not statics:
        raise LookupError(f"No {label!r} label in the editor.")
    anchor = statics[0].rectangle()

    def on_line(rect: Any) -> bool:
        return abs(rect.mid_point().y - anchor.mid_point().y) <= LINE_TOLERANCE

    # These forms are two columns wide, so "to the right of the label" runs on
    # into the next column's fields. The next label on the line is where this
    # label's inputs stop.
    next_label = min(
        (
            c.rectangle().left
            for c in editor.descendants(control_type="Text")
            if c.element_info.name and c.rectangle().left > anchor.right and on_line(c.rectangle())
        ),
        default=None,
    )
    same_line = [
        c
        for c in editor.descendants(control_type=control_type)
        if c.rectangle().left >= anchor.right
        and on_line(c.rectangle())
        and (next_label is None or c.rectangle().left < next_label)
    ]
    if not same_line:
        raise LookupError(f"Found the {label!r} label but no {control_type} beside it.")
    return sorted(same_line, key=lambda c: c.rectangle().left)


def _field(editor: Any, label: str, control_type: str = "Edit") -> Any:
    """Find the single input belonging to `label`.

    Args:
        editor (Any): An editor pane.
        label (str): The visible label text.
        control_type (str): Which kind of input to look for.

    Returns:
        Any: The leftmost matching input.
    """
    return _fields(editor, label, control_type)[0]


def _type_into(field: Any, value: str) -> None:
    """Replace a text box's contents by typing into it.

    Not by the Value pattern: some of these boxes accept a written value,
    report it back happily, and never tell the application about it. The
    multi-line ones do exactly that -- Company took a value, read it back,
    stayed clean in the tab title, and saved empty. Characters posted to the
    control go through the app's own key handling, so the model sees them.

    Args:
        field (Any): The text box.
        value (str): What it should contain.

    Raises:
        LookupError: If the box has no window handle to type into.
    """
    handle = getattr(field.element_info, "handle", None)
    if not handle:
        raise LookupError("Text box has no window handle; cannot type into it.")
    win32gui.SendMessage(handle, win32con.EM_SETSEL, 0, -1)
    win32gui.SendMessage(handle, win32con.WM_CLEAR, 0, 0)
    for character in value:
        win32gui.PostMessage(handle, win32con.WM_CHAR, ord(character), 0)


def _set_text(editor: Any, label: str, value: str, index: int = 0) -> None:
    """Put `value` into a text box and check it stuck.

    Args:
        editor (Any): An editor pane.
        label (str): The box's label.
        value (str): What to write.
        index (int): Which box, when one label covers several.

    Raises:
        ValueError: If the box does not hold `value` afterwards.
    """
    field = _fields(editor, label)[index]
    tracing.point_at(field)
    _type_into(field, value)
    # Typed characters arrive asynchronously, so give the box a moment to
    # catch up before deciding it refused the value.
    deadline = time.monotonic() + TYPING_TIMEOUT
    while True:
        shown = _fields(editor, label)[index].get_value()
        if shown == value:
            return
        if time.monotonic() >= deadline:
            raise ValueError(f"{label!r} holds {shown!r} after entering {value!r}.")
        time.sleep(0.1)


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

    Typed in, like every other text box here: writing a value straight into
    the control is what a UIA Value pattern offers, but some of these boxes
    take such a write without telling the application, and then save empty.

    Args:
        editor (Any): The editor pane from `open_new_order`.
        reference (str): The order's external reference, e.g.
            "WEB-2026-0714-A17" (`SalesOrder.external_reference`).

    Raises:
        ValueError: If the field does not hold the reference afterwards.
    """
    with tracing.step(f"enter {reference!r} in Cust.Ref."):
        _set_text(editor, CUSTOMER_REFERENCE_LABEL, reference)
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
            _select_option(combo, mode)

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
            _select_option(combo, mode)

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


def _tab_count(win: WindowSpecification, title_re: str) -> int:
    """Count open editor tabs whose label matches `title_re`.

    Args:
        win (WindowSpecification): The main window.
        title_re (str): Pattern to match tab labels against.

    Returns:
        int: How many are open.
    """
    pattern = re.compile(title_re)
    return sum(1 for t in win.descendants(control_type="TabItem") if pattern.match(t.element_info.name or ""))


def split_contact_name(contact_name: str) -> tuple[str, str]:
    """Split a contact into the first and last name Fakturama asks for.

    The document gives one name ("Marta Klein"); the editor has two boxes. We
    take the first word as the first name and everything after it as the last
    name, so compound surnames ("Marta van der Berg") stay whole. A single
    word is treated as a last name, leaving First Name empty rather than
    inventing a split.

    Args:
        contact_name (str): The extracted contact name.

    Returns:
        tuple[str, str]: First name, last name; either may be empty.
    """
    parts = contact_name.split()
    if len(parts) < 2:
        return "", " ".join(parts)
    return parts[0], " ".join(parts[1:])


def set_debtor_identity(
    editor: Any,
    company: str | None = None,
    contact_name: str | None = None,
    salutation: str | None = None,
) -> None:
    """Fill the debtor's company and name (step 2.6).

    The proposed Customer ID is Fakturama's to allocate and is never touched.
    Salutation is left at "---" unless the document supplied one, which is
    what the step asks for -- so passing None asserts it rather than setting
    it.

    Args:
        editor (Any): The debtor editor from `open_new_debtor`.
        company (str | None): Extracted company, if any.
        contact_name (str | None): Extracted contact, split into the two name
            boxes by `split_contact_name`.
        salutation (str | None): Salutation to select, or None to leave the
            editor's "---" in place.

    Raises:
        ValueError: If a field does not hold what we entered, or Salutation is
            on something other than "---" when none was supplied.
    """
    if company:
        with tracing.step(f"enter company {company!r}"):
            _set_text(editor, "Company", company)

    if contact_name:
        first, last = split_contact_name(contact_name)
        with tracing.step(f"enter name {first!r} {last!r}"):
            if first:
                _set_text(editor, NAME_LABEL, first, index=0)
            _set_text(editor, NAME_LABEL, last, index=1)

    combo = _field(editor, "Salutation", control_type="ComboBox")
    if salutation:
        with tracing.step(f"set Salutation to {salutation!r}"):
            tracing.point_at(combo)
            combo.select(salutation)
        shown = _field(editor, "Salutation", control_type="ComboBox").selected_text()
        if shown != salutation:
            raise ValueError(f"Salutation is {shown!r} after selecting {salutation!r}.")
    elif combo.selected_text() != NO_SALUTATION:
        raise ValueError(
            f"Salutation is {combo.selected_text()!r}, expected {NO_SALUTATION!r} "
            "for a document that supplies none."
        )


def select_tab(editor: Any, name: str) -> Any:
    """Bring an editor's tab to the front, and return it.

    Args:
        editor (Any): An editor pane.
        name (str): The tab's label, e.g. "Main address".

    Returns:
        Any: The tab item.

    Raises:
        LookupError: If the editor has no such tab.
    """
    tabs = [t for t in editor.descendants(control_type="TabItem") if t.element_info.name == name]
    if not tabs:
        raise LookupError(f"No {name!r} tab in this editor.")
    tab = tabs[0]
    if not tab.is_selected():
        with tracing.step(f"open the {name} tab"):
            activate_tab(tab)
    return tab


def activate_tab(tab: Any) -> None:
    """Bring a tab to the front.

    SWT's tabs answer `is_selected` but not the SelectionItem pattern that
    would select them -- calling `select()` raises "Member not found". They do
    respond to a click on the tab strip, so we post one to the tab folder's
    own window at the tab's position within it. Client coordinates, so this
    still needs neither the pointer nor the foreground.

    Args:
        tab (Any): The tab item to select.

    Raises:
        LookupError: If the tab folder has no window handle to click.
        TimeoutError: If the tab does not come to the front.
    """
    tracing.point_at(tab)
    folder = tab.parent()
    handle = getattr(folder.element_info, "handle", None)
    if not handle:
        raise LookupError(f"The tab folder holding {tab.element_info.name!r} has no window handle.")
    left, top, _, _ = win32gui.GetWindowRect(handle)
    middle = tab.rectangle().mid_point()
    position = win32api.MAKELONG(middle.x - left, middle.y - top)
    win32gui.PostMessage(handle, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, position)
    win32gui.PostMessage(handle, win32con.WM_LBUTTONUP, 0, position)

    deadline = time.monotonic() + TAB_TIMEOUT
    while not tab.is_selected():
        if time.monotonic() >= deadline:
            raise TimeoutError(f"The {tab.element_info.name!r} tab did not come to the front.")
        time.sleep(0.1)


def activate_editor(win: WindowSpecification, title_re: str, what: str) -> Any:
    """Bring an editor to the front and return its body.

    An editor pane can only be read while its tab is selected -- Fakturama
    puts just the visible editor's controls in the tree -- so any step that
    goes back to an editor it opened earlier has to activate it first. Where
    several match, the newest (rightmost) is the one this run opened.

    Args:
        win (WindowSpecification): The main window.
        title_re (str): Pattern the tab label must match.
        what (str): The editor's name, for error messages.

    Returns:
        Any: The editor pane.

    Raises:
        LookupError: If no such editor is open.
    """
    pattern = re.compile(title_re)
    tabs = [t for t in win.descendants(control_type="TabItem") if pattern.match(t.element_info.name or "")]
    if not tabs:
        raise LookupError(f"No {what} editor is open.")
    tab = max(tabs, key=lambda t: t.rectangle().left)
    if not tab.is_selected():
        with tracing.step(f"switch to the {what} editor"):
            activate_tab(tab)
    return wait_for_editor(win, title_re, what)


def set_main_address(
    editor: Any,
    street: str | None = None,
    postcode: str | None = None,
    city: str | None = None,
    country: str | None = None,
    email: str | None = None,
    phone: str | None = None,
    additional_name: str | None = None,
    address_specification: str | None = None,
    district: str | None = None,
) -> None:
    """Fill Addresses > Main address from the billing address (step 2.7).

    Every argument is optional and only written when given: the step asks for
    additional name, address specification and district to be filled *only*
    when the document supplies them, and the same rule is the safe one for the
    rest -- a field we have nothing for keeps the editor's own value.

    Args:
        editor (Any): The debtor editor from `open_new_debtor`.
        street (str | None): Street and house number.
        postcode (str | None): Postal code, the left half of "ZIP - City".
        city (str | None): City, the right half.
        country (str | None): Country to select; the editor defaults to
            United States, so an import from anywhere else must set it.
        email (str | None): E-Mail address.
        phone (str | None): Telephone number.
        additional_name (str | None): Only if the document supplies one.
        address_specification (str | None): Only if the document supplies one.
        district (str | None): Only if the document supplies one.

    Raises:
        ValueError: If a field does not hold what we entered.
    """
    select_tab(editor, ADDRESSES_TAB)
    select_tab(editor, MAIN_ADDRESS_TAB)

    with tracing.step("fill the main address"):
        if street:
            _set_text(editor, "Street", street)
        if postcode:
            _set_text(editor, POSTCODE_CITY_LABEL, postcode, index=0)
        if city:
            _set_text(editor, POSTCODE_CITY_LABEL, city, index=1)
        if email:
            _set_text(editor, "E-Mail", email)
        if phone:
            _set_text(editor, "Telephone", phone)
        if additional_name:
            _set_text(editor, "additional name", additional_name)
        if address_specification:
            _set_text(editor, "Address specification", address_specification)
        if district:
            _set_text(editor, "district", district)

    if country:
        combo = _field(editor, "Country", control_type="ComboBox")
        if combo.selected_text() != country:
            with tracing.step(f"set Country to {country!r}"):
                tracing.point_at(combo)
                combo.select(country)
        shown = _field(editor, "Country", control_type="ComboBox").selected_text()
        if shown != country:
            raise ValueError(f"Country is {shown!r} after selecting {country!r}.")


def set_address_roles(editor: Any, invoice: bool = True, delivery: bool = False) -> None:
    """Give the main address its roles (step 2.8).

    The roles live behind the "address type" button, in a popup of two
    checkboxes that closes as soon as it loses focus -- so it is opened, set
    and dismissed in one go, with nothing in between.

    Setting the delivery role here is what avoids a second address: when the
    document's billing and delivery addresses are the same place, this one
    address plays both parts.

    Args:
        editor (Any): The debtor editor from `open_new_debtor`.
        invoice (bool): Give it the invoice address role.
        delivery (bool): Give it the delivery address role too, which the
            caller decides by comparing the two extracted addresses.

    Raises:
        LookupError: If the address type control or its popup is not found.
        ValueError: If the roles do not read back as asked for.
    """
    wanted = {ROLE_INVOICE: invoice, ROLE_DELIVERY: delivery}
    labels = [c for c in editor.descendants(control_type="Text") if c.element_info.name == ADDRESS_TYPE_LABEL]
    if not labels:
        raise LookupError(f"No {ADDRESS_TYPE_LABEL!r} row in the debtor editor.")
    anchor = labels[0].rectangle()

    def beside(control_type: str) -> Any:
        matches = [
            c
            for c in editor.descendants(control_type=control_type)
            if c.rectangle().left > anchor.right
            and abs(c.rectangle().mid_point().y - anchor.mid_point().y) <= LINE_TOLERANCE
        ]
        if not matches:
            raise LookupError(f"No {control_type} beside the {ADDRESS_TYPE_LABEL!r} label.")
        return min(matches, key=lambda c: c.rectangle().left)

    roles = ", ".join(name for name, on in wanted.items() if on) or "no roles"
    with tracing.step(f"set the address type to {roles}"):
        button = beside("Button")
        tracing.point_at(button)
        button.iface_invoke.Invoke()
        popup = _wait_for_role_popup(editor)
        boxes = {c.element_info.name: c for c in popup.descendants(control_type="CheckBox")}
        missing = set(wanted) - set(boxes)
        if missing:
            raise LookupError(f"The address type popup has no {sorted(missing)} checkbox.")
        for name, should_be_on in wanted.items():
            if bool(boxes[name].get_toggle_state()) != should_be_on:
                boxes[name].iface_toggle.Toggle()
        _close_popup(popup)

    shown = beside("Edit").get_value()
    for name, should_be_on in wanted.items():
        if (name in shown) != should_be_on:
            raise ValueError(f"Address type reads {shown!r}, expected {roles}.")


def _wait_for_role_popup(editor: Any, timeout: int = DIALOG_TIMEOUT) -> WindowSpecification:
    """Wait for the untitled popup the address type button opens.

    Args:
        editor (Any): The debtor editor, to identify the process.
        timeout (int): Seconds to wait.

    Returns:
        WindowSpecification: The popup.

    Raises:
        TimeoutError: If it does not appear.
    """
    _, process_id = win32process.GetWindowThreadProcessId(editor.element_info.handle)
    deadline = time.monotonic() + timeout
    while True:
        found: list[int] = []

        def visit(handle: int, _: Any) -> None:
            _, pid = win32process.GetWindowThreadProcessId(handle)
            if (
                pid == process_id
                and win32gui.IsWindowVisible(handle)
                and win32gui.GetClassName(handle) == DIALOG_CLASS
                and not win32gui.GetWindowText(handle)
            ):
                found.append(handle)

        win32gui.EnumWindows(visit, None)
        if found:
            handle = found[0]
            return Application(backend="uia").connect(handle=handle).window(handle=handle)
        if time.monotonic() >= deadline:
            raise TimeoutError(f"The {ADDRESS_TYPE_LABEL!r} popup did not open within {timeout}s.")
        time.sleep(0.1)


def _close_popup(popup: WindowSpecification) -> None:
    """Dismiss a focus-follows popup with Escape.

    Args:
        popup (WindowSpecification): The popup to close.
    """
    handle = popup.element_info.handle
    win32gui.PostMessage(handle, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
    win32gui.PostMessage(handle, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)


def _select_option(combo: Any, value: str) -> None:
    """Pick `value` in a combo, allowing for a list that is not ready yet.

    A combo dropped down a moment after its editor was built can report that
    it has no such item, and have it a moment later. Retried rather than
    trusted first time.

    Args:
        combo (Any): The combo box.
        value (str): The option's text, exactly as the app spells it.

    Raises:
        Exception: Whatever the last attempt raised.
    """
    for attempt in range(COMBO_ATTEMPTS):
        try:
            combo.select(value)
            return
        except Exception:
            if attempt == COMBO_ATTEMPTS - 1:
                raise
            time.sleep(RESULTS_POLL_SECONDS)


def _set_combo(editor: Any, label: str, value: str) -> None:
    """Put a combo on `value` and check it stuck.

    Args:
        editor (Any): An editor pane.
        label (str): The combo's label.
        value (str): The option to select.

    Raises:
        ValueError: If the combo does not show `value` afterwards.
    """
    combo = _field(editor, label, control_type="ComboBox")
    if combo.selected_text().strip() != value:
        tracing.point_at(combo)
        # Some of these lists label their options with trailing spaces, so the
        # string to select is whichever option reads as `value`.
        options = combo_options(editor, label)
        matching = [o for o in options if o.strip() == value]
        if not matching:
            raise LookupError(f"{label!r} has no {value!r} option; it offers {options}.")
        _select_option(combo, matching[0])
    shown = _field(editor, label, control_type="ComboBox").selected_text().strip()
    if shown != value:
        raise ValueError(f"{label!r} is {shown!r} after selecting {value!r}.")


def set_debtor_miscellaneous(
    editor: Any,
    alias: str | None = None,
    discount: str = NO_DISCOUNT,
    price_mode: str = NET_PRICE_MODE,
) -> None:
    """Fill the debtor's Miscellaneous tab (step 2.9).

    Discount and Net/Gross are settings of the debtor, not of this document:
    they decide how every future order for this customer is priced, which is
    why the step pins them rather than leaving Fakturama's defaults. Discount
    already reads 0% on a new debtor and Net/Gross reads "---", so in practice
    only the latter is written -- but both are asserted afterwards.

    Args:
        editor (Any): The debtor editor from `open_new_debtor`.
        alias (str | None): The extracted customer alias, if any.
        discount (str): Discount to leave the debtor on.
        price_mode (str): "Net" or "Gross".

    Raises:
        ValueError: If a field does not hold what we entered.
    """
    select_tab(editor, MISCELLANEOUS_TAB)
    with tracing.step(f"set alias {alias!r}, discount {discount}, prices {price_mode}"):
        if alias:
            _set_text(editor, ALIAS_LABEL, alias)
        if _field(editor, DISCOUNT_LABEL).get_value() != discount:
            _set_text(editor, DISCOUNT_LABEL, discount)
        _set_combo(editor, NET_GROSS_LABEL, price_mode)


class PaymentMethodUnavailable(LookupError):
    """Raised when the document's payment method is not in Fakturama yet.

    Carries the methods that do exist, since the caller's next move is to
    create the missing one (steps 2.10.1-2.10.6) and it needs to know what it
    is choosing between.
    """

    def __init__(self, method: str, available: list[str]) -> None:
        super().__init__(
            f"No payment method named {method!r}; this install has {available}."
        )
        self.method = method
        self.available = available


def combo_options(editor: Any, label: str) -> list[str]:
    """List what a combo offers.

    The options only exist in the tree while the combo is open, so this drops
    it down and closes it again.

    Args:
        editor (Any): An editor pane.
        label (str): The combo's label.

    Returns:
        list[str]: The option labels, in the order shown.
    """
    combo = _field(editor, label, control_type="ComboBox")
    combo.expand()
    try:
        deadline = time.monotonic() + COMBO_TIMEOUT
        options: list[str] = []
        while time.monotonic() < deadline:
            options = [i.element_info.name for i in combo.descendants(control_type="ListItem")]
            if options:
                break
            time.sleep(0.1)
        return options
    finally:
        combo.collapse()


def set_debtor_payment(editor: Any, method: str) -> None:
    """Select the debtor's payment method (step 2.10).

    The method has to match the document exactly -- "Bank Transfer" is a
    different thing from "Bank transfer fee" -- so this never falls back to a
    near match. When the exact one is missing it raises rather than picking
    something, leaving the editor untouched and open for the caller to create
    the method and come back.

    Args:
        editor (Any): The debtor editor from `open_new_debtor`.
        method (str): The extracted payment method, e.g. "Bank Transfer".

    Raises:
        PaymentMethodUnavailable: If no method of exactly that name exists.
        ValueError: If the combo does not show `method` afterwards.
    """
    select_tab(editor, MISCELLANEOUS_TAB)
    available = combo_options(editor, PAYMENT_LABEL)
    if method not in available:
        raise PaymentMethodUnavailable(method, available)
    with tracing.step(f"set the payment method to {method!r}"):
        _set_combo(editor, PAYMENT_LABEL, method)


class ManualReviewRequired(Exception):
    """Raised when the app's data is ambiguous and a person has to look.

    The steps are explicit that an ambiguous or conflicting match is not
    something to resolve by picking one -- so this carries the rows that
    caused it and leaves every editor open where it was.
    """

    def __init__(self, what: str, rows: list[table.Row]) -> None:
        super().__init__(f"{what}: {[r.cells for r in rows]}")
        self.rows = rows


def open_list_view(win: WindowSpecification, name: str, timeout: int = EDITOR_TIMEOUT) -> Any:
    """Open one of the left-hand Data lists and return its view (step 2.10.1).

    Args:
        win (WindowSpecification): The main window.
        name (str): The entry's label, e.g. "terms of payment".
        timeout (int): Seconds to wait for the view.

    Returns:
        Any: The view pane.

    Raises:
        LookupError: If the left panel has no such entry.
        TimeoutError: If the view does not open.
    """
    view = win.child_window(title=name, control_type="Tab")
    if view.exists():
        return activate_editor(win, f"^{re.escape(name)}$", name)

    links = [c for c in win.descendants(control_type="Text") if c.element_info.name == name]
    if not links:
        raise LookupError(f"No {name!r} entry in the left panel.")
    with tracing.step(f"open Data > {name}"):
        tracing.point_at(links[0])
        _post_click(links[0])
        editor = wait_for_editor(win, f"^{re.escape(name)}$", name, timeout)
    tracing.point_at(editor, colour=tracing.CONFIRM)
    return editor


def find_payment_method(win: WindowSpecification, method: str) -> table.Row | None:
    """Look a payment method up in terms of payment (steps 2.10.1-2.10.2).

    A method counts only when both its Name and its Description are exactly
    the extracted string, which is how 2.10.3 writes them. A row that matches
    one but not the other is a conflicting definition of the same thing, and a
    second exact row is an ambiguity -- neither is ours to resolve.

    Args:
        win (WindowSpecification): The main window.
        method (str): The extracted payment method, e.g. "Bank Transfer".

    Returns:
        table.Row | None: The matching row, or None if there is none and it
            has to be created.

    Raises:
        ManualReviewRequired: If several exact rows, or a conflicting one,
            come back.
    """
    view = open_list_view(win, TERMS_OF_PAYMENT_VIEW)
    rows = search_list(view, method)
    exact = [r for r in rows if r.get("Name") == method and r.get("Description") == method]
    conflicting = [r for r in rows if (r.get("Name") == method) != (r.get("Description") == method)]
    if len(exact) > 1:
        raise ManualReviewRequired(f"Several payment methods called {method!r}", exact)
    if conflicting:
        raise ManualReviewRequired(f"A conflicting definition of {method!r}", conflicting)
    return exact[0] if exact else None


def create_payment_method(win: WindowSpecification, method: str) -> Any:
    """Fill in a new term of payment for `method` (steps 2.10.3-2.10.5).

    Leaves the editor open and unsaved: saving is step 2.10.6, and it is the
    first thing in this flow that writes to the database.

    Name and Description both get the exact extracted string, which is what
    makes `find_payment_method` able to recognise it later -- Fakturama shows
    the Description in the debtor's Payment list and the Name in the terms of
    payment list, so a method that reads the same in both places is the one
    unambiguous choice.

    Account is left blank, the three texts are left blank, and "Set as
    standard" is not touched: this method is for one customer, not the
    install's default.

    Args:
        win (WindowSpecification): The main window.
        method (str): The extracted payment method, e.g. "Bank Transfer".

    Returns:
        Any: The unsaved editor.

    Raises:
        LookupError: If the method has no payment code in `PAYMENT_CODES`, or
            the view has no create button.
        ValueError: If a field does not hold what we entered.
    """
    if method not in PAYMENT_CODES:
        raise LookupError(
            f"No payment code mapped for {method!r}; known methods are {sorted(PAYMENT_CODES)}."
        )

    view = open_list_view(win, TERMS_OF_PAYMENT_VIEW)
    buttons = [c for c in view.descendants(control_type="Button") if c.element_info.name == NEW_TERM_BUTTON]
    if not buttons:
        raise LookupError(f"No {NEW_TERM_BUTTON!r} button in the {TERMS_OF_PAYMENT_VIEW!r} view.")

    with tracing.step(f"create the {method!r} term of payment"):
        tracing.point_at(buttons[0])
        buttons[0].iface_invoke.Invoke()
        editor = wait_for_editor(win, NEW_TERM_TAB_RE, "New Term of Payment")

        # 2.10.3: the method's name, twice; Account stays blank.
        _set_text(editor, "Name", method)
        _set_text(editor, "Description", method)

        # 2.10.4: the code Fakturama files this kind of payment under.
        _set_combo(editor, PAYMENT_CODE_LABEL, PAYMENT_CODES[method])

        # 2.10.5: no discount, no terms. These are the editor's defaults, so
        # they are checked rather than rewritten.
        for label, zero in ((CASH_DISCOUNT_LABEL, NO_CASH_DISCOUNT), *((d, ZERO_DAYS) for d in DAY_LABELS)):
            if _field(editor, label).get_value() != zero:
                _set_text(editor, label, zero)

    for label in BLANK_TEXT_LABELS:
        text = _field(editor, label).get_value()
        if text:
            raise ValueError(f"{label!r} should be blank on a new term of payment, but holds {text!r}.")
    tracing.point_at(editor, colour=tracing.CONFIRM)
    return editor


def save_editor(win: WindowSpecification, title_re: str, what: str, timeout: int = EDITOR_TIMEOUT) -> None:
    """Save one editor with the toolbar's Save, once (step 2.10.6).

    Save acts on whichever editor is in front, so the editor is brought
    forward first -- with several unsaved tabs open, saving the wrong one is
    the easy mistake. Eclipse's unsaved marker is the receipt: the tab loses
    its leading star when the record is written, so that is what we wait for
    rather than a fixed pause, and clicking twice cannot happen.

    Args:
        win (WindowSpecification): The main window.
        title_re (str): Pattern matching the editor's tab.
        what (str): The editor's name, for messages.
        timeout (int): Seconds to wait for the save to land.

    Raises:
        LookupError: If the toolbar has no Save button.
        TimeoutError: If the editor is still unsaved afterwards.
    """
    activate_editor(win, title_re, what)
    buttons = [c for c in win.descendants(control_type="Button") if c.element_info.name == SAVE_BUTTON]
    if not buttons:
        raise LookupError(f"No {SAVE_BUTTON!r} button in the toolbar.")

    # Hold on to this editor's own tab. Other unsaved editors of the same kind
    # may be open -- watching "no dirty tab of this name" would wait for them
    # too -- and a saved record renames its tab (to "Bank Transfer", say), so
    # the name is no help either. The element stays the same through both.
    pattern = re.compile(title_re)
    tabs = [
        t
        for t in win.descendants(control_type="TabItem")
        if pattern.match(t.element_info.name or "") and t.is_selected()
    ]
    if not tabs:
        raise LookupError(f"The {what} editor is not the one in front.")
    tab = tabs[0]

    # Save acts on the part with the focus, which is not the same as the tab
    # that happens to be showing: a run that read a list view a moment ago
    # leaves the focus there, and Save then does nothing at all. Clicking the
    # tab -- even when it is already the visible one -- hands the editor the
    # focus, so Save has something to write.
    activate_tab(tab)
    if not buttons[0].is_enabled():
        raise RuntimeError(f"{SAVE_BUTTON!r} is disabled; the {what} editor does not have the focus.")

    with tracing.step(f"save the {what}"):
        tracing.point_at(buttons[0])
        buttons[0].iface_invoke.Invoke()
        deadline = time.monotonic() + timeout
        while (tab.element_info.name or "").startswith(DIRTY_MARKER):
            if time.monotonic() >= deadline:
                raise TimeoutError(f"The {what} was still unsaved {timeout}s after clicking {SAVE_BUTTON!r}.")
            time.sleep(0.25)


def select_row(container: Any, row: table.Row) -> None:
    """Click a row of a drawn list (step 2.12).

    The list paints its rows, so there is no row object to select -- only a
    position within the table's window, which is what `table.Row` carries.
    The click is posted to that window, and the list is captured either side
    of it: a row that highlights proves the click landed, and an unchanged
    picture means we aimed at nothing.

    Args:
        container (Any): The dialog or view holding the list.
        row (table.Row): The row to select, from `search_list`.

    Raises:
        LookupError: If the list pane has no window handle.
        RuntimeError: If the list does not react to the click.
    """
    pane = _list_pane(container)
    handle = getattr(pane.element_info, "handle", None)
    if not handle:
        raise LookupError("The result list has no window handle to click.")

    before = _grab(pane).tobytes()
    rectangle = pane.rectangle()
    position = win32api.MAKELONG(int(rectangle.width() * ROW_CLICK_X), int(row.y))
    with tracing.step(f"select {row.cells}"):
        win32gui.PostMessage(handle, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, position)
        win32gui.PostMessage(handle, win32con.WM_LBUTTONUP, 0, position)
        deadline = time.monotonic() + SELECTION_TIMEOUT
        while _grab(pane).tobytes() == before:
            if time.monotonic() >= deadline:
                raise RuntimeError(f"The list did not react to a click on row {row.cells}.")
            time.sleep(RESULTS_POLL_SECONDS)


def choose_address(dialog: WindowSpecification, row: table.Row) -> None:
    """Select a debtor in the address dialog and confirm it (step 2.12).

    Args:
        dialog (WindowSpecification): The Select the address dialog.
        row (table.Row): The debtor's row.
    """
    select_row(dialog, row)
    dismiss_dialog(dialog, "OK")


def order_addresses(editor: Any) -> dict[str, str]:
    """Read the addresses the order has been given (step 2.13).

    Args:
        editor (Any): The order editor.

    Returns:
        dict[str, str]: Address tab name -> the text filled in under it. The
            editor shows one tab per role it has an address for, so an order
            with the same address for both shows one tab named for both.
    """
    addresses: dict[str, str] = {}
    deadline = time.monotonic() + EDITOR_TIMEOUT
    while not any(
        ROLE_INVOICE.lower() in (t.element_info.name or "").lower()
        or ROLE_DELIVERY.lower() in (t.element_info.name or "").lower()
        for t in editor.descendants(control_type="TabItem")
    ):
        if time.monotonic() >= deadline:
            return addresses
        time.sleep(0.25)

    for tab in editor.descendants(control_type="TabItem"):
        name = tab.element_info.name or ""
        if ROLE_INVOICE.lower() in name.lower() or ROLE_DELIVERY.lower() in name.lower():
            select_tab(editor, name)
            boxes = [
                c
                for c in editor.descendants(control_type="Edit")
                if not c.element_info.name and c.rectangle().top > tab.rectangle().bottom
            ]
            addresses[name] = boxes[0].get_value() if boxes else ""
    return addresses


def _cell_matches(visible: str, expected: str) -> bool:
    """Whether a listed cell shows `expected`.

    The list is read as a person reads it, so the comparison has to allow for
    how it displays: a column too narrow for its value ends in an ellipsis,
    and the recogniser writes that variously as "...", ".." or ".". A cell cut
    off that way is matched on the part that is visible, since that is all the
    evidence there is -- and it is still enough to tell one customer from
    another when the rest of the row matches too.

    Args:
        visible (str): What the cell shows.
        expected (str): The extracted value it should be.

    Returns:
        bool: True when the cell agrees with the extracted value.
    """
    shown, wanted = visible.strip(), expected.strip()
    if not shown:
        return False
    if shown.endswith(TRUNCATION_MARKS):
        prefix = shown.rstrip(".… ")
        return bool(prefix) and wanted.casefold().startswith(prefix.casefold())
    return shown.casefold() == wanted.casefold()


def find_debtor(dialog: WindowSpecification, term: str, **expected: str | None) -> table.Row | None:
    """Search the address list and decide what it found (steps 2.2-2.3).

    A row counts as the document's customer only when every field the document
    gives -- company, first name, name, ZIP, city -- reads the same in the
    list. Anything less is not a match: two customers at one company differ
    only by name, and two branches of one company only by address.

    Args:
        dialog (WindowSpecification): The Select the address dialog.
        term (str): What to search for, usually the company.
        **expected (str | None): Extracted values by column name, as spelled
            in `DEBTOR_MATCH_COLUMNS`. Fields the document did not yield are
            passed as None and left out of the comparison.

    Returns:
        table.Row | None: The one matching row, or None when nothing matches
            and the debtor has to be created.

    Raises:
        ManualReviewRequired: If several rows match, or a row is the same
            company recorded with different details.
    """
    wanted = {
        DEBTOR_MATCH_COLUMNS[field]: value
        for field, value in expected.items()
        if value and field in DEBTOR_MATCH_COLUMNS
    }
    rows = search_list(dialog, term)
    exact = [row for row in rows if all(_cell_matches(row.get(c), v) for c, v in wanted.items())]
    if len(exact) > 1:
        raise ManualReviewRequired(f"{len(exact)} debtors match the document", exact)
    if exact:
        return exact[0]

    company = wanted.get(COMPANY_COLUMN)
    conflicting = [row for row in rows if company and _cell_matches(row.get(COMPANY_COLUMN), company)]
    if conflicting:
        raise ManualReviewRequired(
            f"{len(conflicting)} debtor(s) at {company!r} with different details", conflicting
        )
    return None


def _squash(text: str) -> str:
    """Reduce text to letters and digits, for comparing what two forms show."""
    return _NON_ALNUM_RE.sub("", text).casefold()


def confirm_order_addresses(editor: Any, parts_by_role: dict[str, list[str]]) -> dict[str, str]:
    """Check the order's addresses carry what the document said (step 2.4).

    The order does not show the document's wording back: Fakturama composes
    the block from the debtor -- adding the contact's name as a line, writing
    the postcode as "DE-10117" -- so the check is that every part the document
    gave appears, not that the text matches character for character. Spacing
    and punctuation are ignored for the same reason.

    Args:
        editor (Any): The order editor.
        parts_by_role (dict[str, list[str]]): For each address role, the
            extracted values that must appear in it.

    Returns:
        dict[str, str]: What each address role on the order now holds.

    Raises:
        ManualReviewRequired: If a required role is missing from the order, or
            an address is missing something the document gave.
    """
    shown = order_addresses(editor)
    for role, parts in parts_by_role.items():
        text = next((value for name, value in shown.items() if role.casefold() in name.casefold()), None)
        if text is None:
            raise ManualReviewRequired(
                f"The order has no {role!r}; it shows {sorted(shown)}", []
            )
        missing = [part for part in parts if part and _squash(part) not in _squash(text)]
        if missing:
            raise ManualReviewRequired(
                f"The order's {role!r} is missing {missing} -- it reads {text!r}", []
            )
    return shown


def dismiss_dialog(dialog: WindowSpecification, button: str) -> None:
    """Close a modal with one of its buttons, e.g. "OK" or "Cancel".

    Args:
        dialog (WindowSpecification): The dialog to close.
        button (str): The button's label.

    Raises:
        LookupError: If the dialog has no such button.
    """
    matches = [c for c in dialog.descendants(control_type="Button") if c.element_info.name == button]
    if not matches:
        raise LookupError(f"No {button!r} button in the {dialog.window_text()!r} dialog.")
    title = dialog.window_text()
    handle = dialog.element_info.handle
    with tracing.step(f"click {button} in {title!r}"):
        tracing.point_at(matches[0])
        matches[0].iface_invoke.Invoke()
        # Wait for it to actually go: the caller's next move is usually to read
        # what the dialog did to the editor behind it, and that editor is still
        # mid-update while the modal is on screen.
        deadline = time.monotonic() + DIALOG_TIMEOUT
        while win32gui.IsWindow(handle) and win32gui.IsWindowVisible(handle):
            if time.monotonic() >= deadline:
                raise TimeoutError(f"The {title!r} dialog stayed open after clicking {button!r}.")
            time.sleep(0.1)


def open_new_debtor(win: WindowSpecification, timeout: int = EDITOR_TIMEOUT) -> Any:
    """Click New Contact in the left New panel and wait for the editor (step 2.5).

    The order editor is left exactly as it is: this opens a second editor
    alongside it, which is what the later steps need -- the debtor is created
    and saved, then picked up from the still-open order.

    Args:
        win (WindowSpecification): The main window.
        timeout (int): Seconds to wait for the editor.

    Returns:
        Any: The New Debtor editor pane.

    Raises:
        LookupError: If the New panel has no New Contact entry.
        TimeoutError: If no new debtor editor opens.
    """
    links = [c for c in win.descendants(control_type="Text") if c.element_info.name == NEW_CONTACT_LINK]
    if not links:
        raise LookupError(f"No {NEW_CONTACT_LINK!r} entry in the left panel.")

    # Fakturama reuses the tab name for every unsaved debtor, so "did a new one
    # open" is a count, not a name.
    before = _tab_count(win, NEW_DEBTOR_TAB_RE)
    with tracing.step(f"click {NEW_CONTACT_LINK!r} in the New panel"):
        tracing.point_at(links[0])
        _post_click(links[0])
        deadline = time.monotonic() + timeout
        while _tab_count(win, NEW_DEBTOR_TAB_RE) <= before:
            if time.monotonic() >= deadline:
                raise TimeoutError(
                    f"No new debtor editor opened within {timeout}s of clicking {NEW_CONTACT_LINK!r}."
                )
            time.sleep(0.25)
        editor = wait_for_editor(win, NEW_DEBTOR_TAB_RE, "New Debtor", timeout)
    tracing.point_at(editor, colour=tracing.CONFIRM)
    return editor


def _post_click(ctrl: Any) -> None:
    """Click `ctrl` by posting mouse messages to its own window handle.

    For controls with an Invoke pattern we invoke; these address icons are
    SWT labels with a mouse listener and expose no pattern at all, so a click
    is the only way to work them. We post WM_LBUTTONDOWN/UP straight to the
    control's HWND rather than moving the real pointer: a pointer click needs
    the window foreground and unobscured, takes the mouse away from whoever is
    using the machine, and lands in whatever app happens to be in front if the
    window is not where we last saw it (verified the hard way -- a stale rect
    put a click into a browser window).

    Args:
        ctrl (Any): The control to click; must have a native window handle.

    Raises:
        LookupError: If the control has no window handle to post to.
    """
    handle = getattr(ctrl.element_info, "handle", None)
    if not handle:
        raise LookupError("Control has no window handle; cannot post a click to it.")
    rect = ctrl.rectangle()
    # Client coordinates, so the widget sees the press over its own middle.
    position = win32api.MAKELONG((rect.right - rect.left) // 2, (rect.bottom - rect.top) // 2)
    win32gui.PostMessage(handle, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, position)
    win32gui.PostMessage(handle, win32con.WM_LBUTTONUP, 0, position)


def _find_dialog(process_id: int, title: str) -> int | None:
    """Return the handle of a visible top-level `title` window, or None.

    Fakturama's modals are plain Win32 dialogs owned by the main window, and
    pywinauto's `Application.windows()` does not list them, so we enumerate
    the desktop ourselves and keep to the one process.

    Args:
        process_id (int): The Fakturama process.
        title (str): Exact window caption.

    Returns:
        int | None: The window handle if it is open.
    """
    found: list[int] = []

    def visit(handle: int, _: Any) -> None:
        _, pid = win32process.GetWindowThreadProcessId(handle)
        if pid == process_id and win32gui.IsWindowVisible(handle) and win32gui.GetWindowText(handle) == title:
            found.append(handle)

    win32gui.EnumWindows(visit, None)
    return found[0] if found else None


def wait_for_dialog(win: WindowSpecification, title: str, timeout: int = DIALOG_TIMEOUT) -> WindowSpecification:
    """Wait for a modal dialog of the given title and return it.

    Args:
        win (WindowSpecification): The main window, to identify the process.
        title (str): Exact window caption.
        timeout (int): Seconds to wait.

    Returns:
        WindowSpecification: The dialog.

    Raises:
        TimeoutError: If it does not open in time.
    """
    _, process_id = win32process.GetWindowThreadProcessId(win.element_info.handle)
    deadline = time.monotonic() + timeout
    while True:
        handle = _find_dialog(process_id, title)
        if handle:
            return Application(backend="uia").connect(handle=handle).window(handle=handle)
        if time.monotonic() >= deadline:
            raise TimeoutError(f"No {title!r} dialog appeared within {timeout}s.")
        time.sleep(0.25)


def open_address_selector(editor: Any, win: WindowSpecification) -> WindowSpecification:
    """Open Select the address from the order's Addresses row (step 2.1).

    Clicks the upper icon, which picks an existing contact. The lower green +
    beside it starts a brand new debtor and must not be touched here, so we
    locate the pair and take the topmost rather than matching on an icon.

    Args:
        editor (Any): The editor pane from `open_new_order`.
        win (WindowSpecification): The main window, to find the dialog.

    Returns:
        WindowSpecification: The open dialog.

    Raises:
        LookupError: If the Addresses row does not hold exactly the two icons
            we expect -- safer than guessing which one creates a debtor.
    """
    labels = [c for c in editor.descendants(control_type="Text") if c.element_info.name == "Addresses"]
    if not labels:
        raise LookupError("No 'Addresses' label in the order editor.")
    icons = sorted(labels[0].parent().children(control_type="Image"), key=lambda c: c.rectangle().top)
    if len(icons) != ADDRESS_ICON_COUNT:
        raise LookupError(
            f"Expected {ADDRESS_ICON_COUNT} icons beside Addresses, found {len(icons)}; "
            "refusing to guess which one picks an existing contact."
        )

    with tracing.step(f"open {ADDRESS_DIALOG_TITLE!r} from the Addresses row"):
        tracing.point_at(icons[0])
        # An editor built moments ago can take the click before its icon is
        # listening, and then nothing happens at all -- so the click is
        # repeated, against a freshly located icon, until the dialog shows up.
        for attempt in range(DIALOG_ATTEMPTS):
            _post_click(icons[0])
            try:
                dialog = wait_for_dialog(win, ADDRESS_DIALOG_TITLE, timeout=DIALOG_RETRY_SECONDS)
                break
            except TimeoutError:
                if attempt == DIALOG_ATTEMPTS - 1:
                    raise
                icons = sorted(
                    labels[0].parent().children(control_type="Image"),
                    key=lambda c: c.rectangle().top,
                )
    tracing.point_at(dialog, colour=tracing.CONFIRM)
    return dialog


def _grab(ctrl: Any) -> Image.Image:
    """Capture a control's own pixels, even when another window covers it.

    A screen grab of a background window returns whatever is on top of it, so
    we ask the control to render itself into a bitmap instead. That keeps the
    run independent of what else is on screen -- and of where the window is.

    Args:
        ctrl (Any): A control with a native window handle.

    Returns:
        Image.Image: The control as an RGB image.

    Raises:
        LookupError: If the control has no window handle.
        RuntimeError: If Windows declines to render it.
    """
    handle = getattr(ctrl.element_info, "handle", None)
    if not handle:
        raise LookupError("Control has no window handle; cannot capture it.")
    left, top, right, bottom = win32gui.GetWindowRect(handle)
    window_dc = win32gui.GetWindowDC(handle)
    source = win32ui.CreateDCFromHandle(window_dc)
    memory = source.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    try:
        bitmap.CreateCompatibleBitmap(source, right - left, bottom - top)
        memory.SelectObject(bitmap)
        if not ctypes.windll.user32.PrintWindow(handle, memory.GetSafeHdc(), PW_RENDERFULLCONTENT):
            raise RuntimeError(f"PrintWindow failed for window {handle}.")
        info = bitmap.GetInfo()
        return Image.frombuffer(
            "RGB", (info["bmWidth"], info["bmHeight"]), bitmap.GetBitmapBits(True), "raw", "BGRX", 0, 1
        )
    finally:
        win32gui.DeleteObject(bitmap.GetHandle())
        memory.DeleteDC()
        source.DeleteDC()
        win32gui.ReleaseDC(handle, window_dc)


def _search_box(container: Any) -> Any:
    """Return a dialog's or list view's search box.

    Args:
        container (Any): A dialog or list view holding a searchable list.

    Returns:
        Any: The search Edit.

    Raises:
        LookupError: If the dialog has no search box.
    """
    labels = [c for c in container.descendants(control_type="Text") if c.element_info.name == SEARCH_LABEL]
    edits = container.descendants(control_type="Edit")
    if not labels or not edits:
        raise LookupError(f"No {SEARCH_LABEL!r} box in {container.element_info.name!r}.")
    anchor = labels[0].rectangle()
    beside = [e for e in edits if abs(e.rectangle().mid_point().y - anchor.mid_point().y) <= LINE_TOLERANCE]
    return min(beside or edits, key=lambda e: abs(e.rectangle().left - anchor.right))


def _list_pane(container: Any) -> Any:
    """Return the pane holding the result rows.

    The list is a drawn table (a canvas, not a tree of controls), so it has no
    rows in the accessibility tree at all -- only this one pane. We take the
    pane between the search row and the buttons, which is the table's own
    window and therefore capturable on its own.

    Args:
        container (Any): A dialog or list view holding a searchable list.

    Returns:
        Any: The results pane.

    Raises:
        LookupError: If no such pane is there.
    """
    top = _search_box(container).rectangle().bottom
    bottom = min(
        (c.rectangle().top for c in container.descendants(control_type="Button") if c.element_info.name == "OK"),
        default=container.rectangle().bottom,
    )
    panes = [
        c
        for c in container.descendants(control_type="Pane")
        if getattr(c.element_info, "handle", None)
        and c.rectangle().top >= top
        and c.rectangle().bottom <= bottom
    ]
    if not panes:
        raise LookupError("No result list pane between the search box and the buttons.")
    return max(panes, key=lambda c: c.rectangle().width() * c.rectangle().height())


def wait_for_results(container: Any, timeout: int = RESULTS_TIMEOUT) -> Image.Image:
    """Wait until the result list stops redrawing, and return its picture.

    Args:
        container (Any): The dialog or list view holding the list.
        timeout (int): Seconds to wait for it to settle.

    Returns:
        Image.Image: The settled list, for reading the rows off.

    Raises:
        TimeoutError: If it is still changing when the time is up.
    """
    pane = _list_pane(container)
    deadline = time.monotonic() + timeout
    previous: bytes | None = None
    identical = 1
    while True:
        shot = _grab(pane)
        current = shot.tobytes()
        identical = identical + 1 if current == previous else 1
        if identical >= RESULTS_STABLE_FRAMES:
            return shot
        if time.monotonic() >= deadline:
            raise TimeoutError(f"The result list was still changing after {timeout}s.")
        previous = current
        time.sleep(RESULTS_POLL_SECONDS)


def search_list(container: Any, term: str) -> list[table.Row]:
    """Search a list and read back what it shows (steps 2.2, 2.10.1).

    Args:
        container (Any): The dialog or list view holding the list.
        term (str): What to search for.

    Returns:
        list[table.Row]: The rows the list settled on, each keyed by the
            column names it draws and carrying where it sits in the list.

    Raises:
        ValueError: If the search box does not hold `term` afterwards.
    """
    with tracing.step(f"search for {term!r}"):
        tracing.point_at(_search_box(container))
        # A view opened a moment ago accepts the value and then clears it as
        # it finishes building, so the box is re-found and rewritten until it
        # keeps what we put in it.
        for attempt in range(SEARCH_ATTEMPTS):
            _search_box(container).iface_value.SetValue(term)
            shown = _search_box(container).get_value()
            if shown == term:
                break
            if attempt == SEARCH_ATTEMPTS - 1:
                raise ValueError(f"Search box holds {shown!r} after entering {term!r}.")
            time.sleep(RESULTS_POLL_SECONDS)
        settled = wait_for_results(container)
    tracing.point_at(_list_pane(container), colour=tracing.CONFIRM)
    return table.read(settled)


#: Values the smoke run below drives the UI with. They are what
#: `extract_sales_order("invoice.pdf")` returns, hardcoded so that running
#: this module exercises the app without paying for OCR first; `main.py` runs
#: the same steps on freshly extracted data.
DEMO_ORDER_DATE = "2026-07-14"
DEMO_EXTERNAL_REFERENCE = "WEB-2026-0714-A17"
DEMO_COMPANY = "Northstar Office GmbH"
DEMO_CONTACT_NAME = "Marta Klein"
DEMO_ALIAS = "NORTHSTAR-BERLIN"
DEMO_PAYMENT_METHOD = "Bank Transfer"
DEMO_STREET = "Friedrichstrasse 88"
DEMO_POSTCODE = "10117"
DEMO_CITY = "Berlin"
DEMO_COUNTRY = "Germany"
DEMO_EMAIL = "marta.klein@example.test"
DEMO_PHONE = "+49 30 5550 1420"


if __name__ == "__main__":
    # A smoke run of every step implemented so far, against the running app:
    #   python fakturama.py
    # It leaves an order and a debtor editor open and unsaved, which is where
    # steps 2.6 onwards pick up.
    tracing.configure(visual=True)
    try:
        window = connect()
        print("connected:", window.element_info.name)

        # 1.3-1.7: a new order, with the header filled in.
        order_editor = open_new_order(window)
        print("editor:", order_editor.element_info.name)
        set_date(order_editor, DEMO_ORDER_DATE)
        print("date:", _field(order_editor, "Date").get_value())
        set_customer_reference(order_editor, DEMO_EXTERNAL_REFERENCE)
        print("cust.ref:", _field(order_editor, CUSTOMER_REFERENCE_LABEL).get_value())
        set_price_mode(order_editor)
        set_vat_mode(order_editor)
        print("price mode:", _price_mode_combo(order_editor).selected_text())
        print("vat:", _field_combo(order_editor, "VAT").selected_text())

        # 2.1-2.3: look the debtor up from the order, and decide.
        first_name, last_name = split_contact_name(DEMO_CONTACT_NAME)
        criteria = dict(
            company=DEMO_COMPANY,
            first_name=first_name,
            last_name=last_name,
            postcode=DEMO_POSTCODE,
            city=DEMO_CITY,
        )
        dialog = open_address_selector(order_editor, window)
        existing_debtor = find_debtor(dialog, DEMO_COMPANY, **criteria)
        print("existing debtor:", existing_debtor.cells if existing_debtor else "none -- create one")
        if existing_debtor:
            # 2.4: it is already there, so use it, then check the order really
            # got the document's address before going on to the products.
            choose_address(dialog, existing_debtor)
            order_editor = activate_editor(window, NEW_ORDER_TAB_RE, "New Order")
            filled = confirm_order_addresses(
                order_editor,
                {ROLE_INVOICE: [DEMO_COMPANY, DEMO_STREET, DEMO_POSTCODE, DEMO_CITY, DEMO_COUNTRY]},
            )
            for role, text in filled.items():
                print(f"{role}: {text!r}")
            print("addresses confirmed against the document")
            raise SystemExit(0)
        dismiss_dialog(dialog, "Cancel")

        # 2.5: the debtor editor, opened beside the still-open order.
        debtor_editor = open_new_debtor(window)
        print("debtor editor:", debtor_editor.element_info.name)

        # 2.6: identity, leaving the proposed Customer ID and "---" alone.
        set_debtor_identity(debtor_editor, company=DEMO_COMPANY, contact_name=DEMO_CONTACT_NAME)
        print("customer id:", _field(debtor_editor, "Customer ID").get_value())
        print("company:", _field(debtor_editor, "Company").get_value())
        print("name:", [c.get_value() for c in _fields(debtor_editor, NAME_LABEL)])
        print("salutation:", _field(debtor_editor, "Salutation", control_type="ComboBox").selected_text())

        # 2.7: the billing address. `main.py` splits the extracted one-liner
        # into these parts with `models.parse_postal_address`.
        set_main_address(
            debtor_editor,
            street=DEMO_STREET,
            postcode=DEMO_POSTCODE,
            city=DEMO_CITY,
            country=DEMO_COUNTRY,
            email=DEMO_EMAIL,
            phone=DEMO_PHONE,
        )
        print("street:", _field(debtor_editor, "Street").get_value())
        print("zip/city:", [c.get_value() for c in _fields(debtor_editor, POSTCODE_CITY_LABEL)])
        print("country:", _field(debtor_editor, "Country", control_type="ComboBox").selected_text())
        print("email:", _field(debtor_editor, "E-Mail").get_value())
        print("phone:", _field(debtor_editor, "Telephone").get_value())

        # 2.8: this address is the invoice address. It would take the delivery
        # role too if the document's two addresses were the same place --
        # `main.py` decides that with `models.same_address`.
        set_address_roles(debtor_editor, invoice=True, delivery=False)
        print("address type:", _field(debtor_editor, ADDRESS_TYPE_LABEL).get_value())

        # 2.9: alias, no discount, prices net.
        set_debtor_miscellaneous(debtor_editor, alias=DEMO_ALIAS)
        print("alias:", _field(debtor_editor, ALIAS_LABEL).get_value())
        print("discount:", _field(debtor_editor, DISCOUNT_LABEL).get_value())
        print("net or gross:", _field(debtor_editor, NET_GROSS_LABEL, control_type="ComboBox").selected_text())

        # 2.10: the extracted payment method, if this install has it. Creating
        # a missing one is steps 2.10.1-2.10.6, and leaves this editor open.
        try:
            set_debtor_payment(debtor_editor, DEMO_PAYMENT_METHOD)
            print("payment:", _field(debtor_editor, PAYMENT_LABEL, control_type="ComboBox").selected_text())
        except PaymentMethodUnavailable as exc:
            print("payment: needs creating --", exc)
            # 2.10.1-2.10.2: is it there under a different debtor's nose?
            existing = find_payment_method(window, DEMO_PAYMENT_METHOD)
            print("terms of payment says:", existing.cells if existing else "no exact row -- create it")
            if not existing:
                # 2.10.3-2.10.5: fill it in. Saving is 2.10.6.
                term_editor = create_payment_method(window, DEMO_PAYMENT_METHOD)
                print("new term name:", _field(term_editor, "Name").get_value())
                print("new term description:", _field(term_editor, "Description").get_value())
                print(
                    "payment code:",
                    _field(term_editor, PAYMENT_CODE_LABEL, control_type="ComboBox").selected_text().strip(),
                )
                print(
                    "zeros:",
                    [_field(term_editor, label).get_value() for label in (CASH_DISCOUNT_LABEL, *DAY_LABELS)],
                )

                # 2.10.6: save it, then go back to the debtor and select it.
                save_editor(window, NEW_TERM_TAB_RE, "New Term of Payment")
                print("saved; terms of payment now:", find_payment_method(window, DEMO_PAYMENT_METHOD).cells)
                debtor_editor = activate_editor(window, NEW_DEBTOR_TAB_RE, "New Debtor")
                set_debtor_payment(debtor_editor, DEMO_PAYMENT_METHOD)
                print("payment:", _field(debtor_editor, PAYMENT_LABEL, control_type="ComboBox").selected_text())

        # 2.11: save the debtor, once.
        save_editor(window, NEW_DEBTOR_TAB_RE, "New Debtor")
        print("debtor saved as:", _field(debtor_editor, "Customer ID").get_value())

        # 2.12: back to the order, and pick the debtor we just saved.
        order_editor = activate_editor(window, NEW_ORDER_TAB_RE, "New Order")
        dialog = open_address_selector(order_editor, window)
        found = search_list(dialog, DEMO_COMPANY)
        print(f"searched again; {len(found)} row(s): {[r.cells for r in found]}")
        if len(found) != 1:
            dismiss_dialog(dialog, "Cancel")
            raise ManualReviewRequired(f"{len(found)} debtors match {DEMO_COMPANY!r}", found)
        choose_address(dialog, found[0])

        # 2.13: the order should now carry the debtor's address.
        order_editor = activate_editor(window, NEW_ORDER_TAB_RE, "New Order")
        for role, text in order_addresses(order_editor).items():
            print(f"{role}: {text!r}")
        print("order tabs still open:", _tab_count(window, NEW_ORDER_TAB_RE))
    finally:
        tracing.stop()
