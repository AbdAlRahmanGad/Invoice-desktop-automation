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
from layout import normalize
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
ADDRESSES_LABEL = "Addresses"
ADDRESS_ICON_COUNT = 2

#: Modal that the upper icon opens.
ADDRESS_DIALOG_TITLE = "Select the address"

#: The Items table's four icons, top to bottom: pick a product, add a line
#: (green +), delete a line, duplicate a line. Step 3.2 wants the first and
#: explicitly not the second.
ITEMS_LABEL = "Items"
ITEM_ICON_COUNT = 4
PRODUCT_DIALOG_TITLE = "Select a product"

#: The product list's item-number column, which is what a SKU is matched
#: against (step 3.3).
SKU_COLUMN = "Item No."

#: How long to give the product picker to close itself on a unique match
#: before concluding that it is waiting for us instead.
PICK_SETTLE_SECONDS = 1.5

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

#: How much of a cut-short cell has to be showing before its opening counts
#: as recognising the value. Without a floor, "Ltd" would match half an
#: address book.
MIN_PREFIX_MATCH = 6

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

#: The order's totals, its own charges, and the number it is filed under
#: (steps 4.2-4.4).
ORDER_TOTAL_LABELS = {"net": "Total Net", "vat": "VAT", "gross": "Total"}
DOCUMENT_NUMBER_LABEL = "No."
SHIPPING_LABEL = "Shipping"
FREE_SHIPPING = "Free of shipping costs"

#: Data > Documents, where a saved order is checked (step 4.5).
DOCUMENTS_VIEW = "Documents"
DOCUMENT_COLUMN = "Document"
DOCUMENT_MATCH_COLUMNS = {
    "date": "Date",
    "reference": "Cust.Ref.",
    "state": "State",
    "total": "Total",
}

#: What Documents shows in State: an order that has been saved and not yet
#: turned into anything, and an invoice recorded as settled.
OPEN_STATE = "open"
PAID_STATE = "paid"

#: Raising the next document from the order itself (steps 4.6-4.7). The
#: follow-up controls are tall image buttons, which is what tells them from
#: the identically-named toolbar button that would start an unrelated invoice.
FOLLOW_UP_INVOICE = "Invoice"
FOLLOW_UP_MIN_HEIGHT = 40
FOLLOW_UP_TAB_RE = r"^\*?New {kind}$"

#: The invoice's payment row (steps 5.2-5.3): a checkbox, then an unnamed
#: combo, a value and a date, none of which carry names of their own.
PAID_LABEL = "paid"
PAID_AT_LABEL = "at"
PAID_VALUE_LABEL = "Value"

#: The invoice raised from an order (step 5.1). Its own No., Invoice Date and
#: Service date are proposed by Fakturama and left as they are; everything
#: below is what the follow-up should have brought across from the order.
NEW_INVOICE_TAB_RE = r"^\*?New Invoice$"
ORDER_DATE_LABEL = "Order Date"
INHERITED_LINE_COLUMNS = ("Item No", "Qty", "U.Price", "Discount", "Price")

#: Columns of the item table used to find and check a line.
ITEM_NUMBER_COLUMN = "Item No"
NAME_COLUMN = "Name"

#: Seconds to let a line's own total catch up with a quantity or discount
#: that was just typed into it, and how many times to enter them before
#: giving up on the app taking them.
LINE_SETTLE_TIMEOUT = 10
LINE_ATTEMPTS = 3

#: Seconds to let a resized window settle before reading it.
WINDOW_SETTLE_SECONDS = 1.0

#: A window showing no more children than this has not published its contents
#: to the accessibility layer yet.
THIN_TREE_CHILDREN = 3

#: How far the item table's mouse canvas is inset inside the pane that draws
#: it. The pair is what identifies the table in an editor.
TABLE_INSET = 2

#: How far below its tab an address box starts. Bounded, because the Remarks
#: box further down the form is also unnamed and is the larger of the two.
ADDRESS_BOX_REACH = 40

#: The order's item table sits between these two labels, and its columns are
#: named by their headers (steps 3.13-3.16). "Qty" and "Pos" lose the dot
#: Fakturama prints after them.
REMARKS_LABEL = "Remarks"
QTY_COLUMN = "Qty"
UNIT_PRICE_COLUMN = "U.Price"
LINE_DISCOUNT_COLUMN = "Discount"
LINE_PRICE_COLUMN = "Price"
LINE_VAT_COLUMN = "VAT"

#: Creating a product (steps 3.7-3.11). The gross price and cost price boxes
#: carry no accessible name, so they are found beside their labels.
NEW_PRODUCT_BUTTON = "Create a new product"
NEW_PRODUCT_TAB_RE = r"^\*?New product$"
GROSS_PRICE_LABEL = "Price (gross)"
COST_PRICE_LABEL = "cost price (net)"
STOCK_LABEL = "Stock"
ZERO_AMOUNT = "0.00"

#: The Data list holding VAT rates, and what a rate for an imported item is
#: called there (steps 3.4-3.6). This install ships "MwSt. 19%", which is the
#: same tax under a different name -- the steps want one named for its rate.
VATS_VIEW = "VATs"
VAT_NAME_PREFIX = "VAT"
NEW_VAT_BUTTON = "Create a new tax rate"
NEW_VAT_TAB_RE = r"^\*?New TAX Rate$"
VAT_VALUE_LABEL = "Value"

#: The e-invoice code a normal rate carries, which a new rate already has.
VAT_CODE_LABEL = "VAT code (E-Invoice)"
STANDARD_VAT_CODE = "S (Standard rate)"

#: A value that is only a number, so it should be compared as one however the
#: app decorates it ("678.30" against "$678.30").
_AMOUNT_RE = re.compile(r"[-+]?\d[\d.,]*")

#: A number in displayed text, for reading a percentage back.
_NUMBER_RE = re.compile(r"[-+]?\d*\.?\d+")

#: The Data list holding payment methods (step 2.10.1).
TERMS_OF_PAYMENT_VIEW = "terms of payment"

#: The debtor's payment method (step 2.10). This install keeps it on the
#: Miscellaneous tab rather than a Payment tab of its own.
PAYMENT_LABEL = "Payment"

#: How many times to repeat a UI Automation call that failed transiently, and
#: how long to leave between tries. A dialog that has just appeared can refuse
#: to be looked up for about a second, so the budget covers several of those.
UIA_ATTEMPTS = 8
UIA_RETRY_SECONDS = 0.5

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

#: Seconds to wait for a tab to come to the front, and where along it to
#: click -- well left of the close button on its right-hand end.
TAB_TIMEOUT = 5
TAB_CLICK_X = 0.25


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
        WindowSpecification: The main window.

    Raises:
        RuntimeError: If no Fakturama window is open.
    """
    handle = _retry(_main_window_handle)
    if not handle:
        raise RuntimeError(
            "No running Fakturama window found. Start Fakturama first; this "
            "module attaches to a running instance rather than launching one."
        )
    win = Application(backend="uia").connect(handle=handle).window(handle=handle)
    # "visible", not "ready": a modal dialog disables the main window, and a
    # run that has to reopen the address selector legitimately attaches while
    # one is open.
    win.wait("visible", timeout=EDITOR_TIMEOUT)

    clear_leftover_pickers(win)
    wake(win)
    return win


def clear_leftover_pickers(win: WindowSpecification) -> None:
    """Cancel a picker left open by a run that stopped part way.

    A modal disables the main window, and a disabled window publishes almost
    nothing to the accessibility layer -- so the next run cannot find the
    toolbar, cannot wake the window either, and fails with a puzzle instead of
    a reason. Pickers are ours and hold nothing: Cancel is their do-nothing
    button, so they are simply closed.

    Any other dialog is left alone and reported: it may be the application
    asking something that matters, and answering it is not ours to guess.

    Args:
        win (WindowSpecification): The main window.

    Raises:
        RuntimeError: If a dialog that is not a picker is open.
    """
    _, process_id = win32process.GetWindowThreadProcessId(win.element_info.handle)
    for title in (ADDRESS_DIALOG_TITLE, PRODUCT_DIALOG_TITLE):
        handle = _find_dialog(process_id, title)
        if handle:
            # Closed with WM_CLOSE rather than by its Cancel button: a dialog
            # from an abandoned run publishes as little of itself as the
            # disabled window behind it, so there is no button to find. For a
            # picker this is the same as Cancel -- it holds nothing to lose.
            print(f"closing a {title!r} dialog left open by an earlier run", flush=True)
            win32gui.PostMessage(handle, win32con.WM_CLOSE, 0, 0)
            deadline = time.monotonic() + DIALOG_TIMEOUT
            while win32gui.IsWindow(handle) and win32gui.IsWindowVisible(handle):
                if time.monotonic() >= deadline:
                    raise RuntimeError(f"A {title!r} dialog is open and will not close.")
                time.sleep(0.1)

    if not win32gui.IsWindowEnabled(win.element_info.handle):
        raise RuntimeError(
            f"A dialog is open in Fakturama ({_open_dialog_titles(process_id)}) and the main "
            "window is disabled. Close it and run again."
        )


def _open_dialog_titles(process_id: int) -> list[str]:
    """The captions of the process's visible dialogs, for an error message.

    Args:
        process_id (int): The Fakturama process.

    Returns:
        list[str]: What is open.
    """
    found: list[str] = []

    def visit(handle: int, _: Any) -> None:
        _, pid = win32process.GetWindowThreadProcessId(handle)
        if pid == process_id and win32gui.IsWindowVisible(handle) and win32gui.GetClassName(handle) == DIALOG_CLASS:
            found.append(win32gui.GetWindowText(handle))

    win32gui.EnumWindows(visit, None)
    return found


def wake(win: WindowSpecification, force: bool = False) -> None:
    """Make sure the window has published its contents to the tree.

    Fakturama shows only its title bar and menu until something gives it the
    focus -- every control below that is simply absent, so a lookup for a
    toolbar button finds nothing and the app looks like it has no controls at
    all. It can lapse back into that state between steps, so this is checked
    wherever a run reaches for the window afresh, not only on attaching.

    The focus is not forced while a dialog is up: the main window is disabled
    then, and its tree is already built anyway.

    Args:
        win (WindowSpecification): The main window.
        force (bool): Focus it even though its tree looks populated. A partly
            published tree has children and still lacks the control being
            looked for, so a caller that cannot find one asks for this.
    """
    handle = win.element_info.handle
    if win32gui.IsWindowEnabled(handle) and (force or len(win.children()) <= THIN_TREE_CHILDREN):
        with tracing.step("wake the window's controls"):
            win.set_focus()
            time.sleep(WINDOW_SETTLE_SECONDS)


def _main_window_handle() -> int | None:
    """The Fakturama window's handle, found by its caption.

    Looked up through the window list rather than through UI Automation:
    searching UIA by title intermittently fails to see a window that is
    plainly on screen, and the caption is the same either way.

    Returns:
        int | None: The handle, or None when the app is not running.
    """
    found: list[int] = []

    def visit(handle: int, _: Any) -> None:
        if win32gui.IsWindowVisible(handle) and re.match(APP_TITLE_RE, win32gui.GetWindowText(handle)):
            found.append(handle)

    win32gui.EnumWindows(visit, None)
    return found[0] if found else None


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
        wake(win)
        button = _find_waking(
            win, lambda: win.child_window(title=NEW_ORDER_BUTTON, control_type="Button").wrapper_object()
        )
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


def _type_date(field: Any, wanted: date) -> None:
    """Type a date into one of Fakturama's date widgets.

    They take digits one date-field at a time and advance by themselves; a
    written value is accepted and then thrown away (see `set_date`).

    Args:
        field (Any): The date widget.
        wanted (date): The date it should hold.
    """
    digits = {"day": f"{wanted.day:02d}", "month": f"{wanted.month:02d}", "year": f"{wanted.year:04d}"}
    tracing.point_at(field)
    field.set_focus()
    send_keys("{HOME}" + "".join(digits[name] for name in DATE_FIELD_ORDER))


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

    with tracing.step(f"set Date to {wanted.isoformat()}"):
        _type_date(_field(editor, "Date"), wanted)

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

    shown = _combo_settles_on(lambda: _price_mode_combo(editor).selected_text(), mode)
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

    shown = _combo_settles_on(lambda: _field_combo(editor, "VAT").selected_text(), mode)
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
    # Aim at the tab's left end, not its middle: a selected tab carries a
    # close button on its right, and on a crowded strip the middle is close
    # enough to it to hit it -- which closes the editor and puts up a "Save
    # Parts" prompt instead of switching to it.
    rectangle = tab.rectangle()
    position = win32api.MAKELONG(
        int(rectangle.left + rectangle.width() * TAB_CLICK_X) - left,
        rectangle.mid_point().y - top,
    )
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
        # An empty tree is not an empty application: it can mean the window
        # has stopped publishing its contents.
        wake(win)
        tabs = [t for t in win.descendants(control_type="TabItem") if pattern.match(t.element_info.name or "")]
    if not tabs:
        raise LookupError(f"No {what} editor is open.")
    # One of them may already be in front, in which case that is the one being
    # worked on -- switching away from it to the newest would be wrong, and
    # clicking a tab that the tab strip has scrolled out of reach fails anyway.
    tab = next((t for t in tabs if t.is_selected()), max(tabs, key=lambda t: t.rectangle().left))
    if not tab.is_selected():
        with tracing.step(f"switch to the {what} editor"):
            activate_tab(tab)

    editor = wait_for_editor(win, title_re, what)
    if not editor.children():
        # In front, but with nothing published below it -- the same lapse
        # `wake` deals with, and it would otherwise read as an editor that has
        # no fields in it.
        wake(win)
        editor = wait_for_editor(win, title_re, what)
    return editor


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


def _combo_settles_on(read: Any, value: str) -> str:
    """Wait for a combo to report the option that was just chosen.

    The choice is applied a moment after it is made, so reading straight back
    can still see the old value -- which looks exactly like a refused
    selection. Read until it agrees or the time is up.

    Args:
        read (Any): A callable returning what the combo currently shows.
        value (str): What it should come to show.

    Returns:
        str: What it shows at the end.
    """
    deadline = time.monotonic() + COMBO_TIMEOUT
    shown = read()
    while shown != value and time.monotonic() < deadline:
        time.sleep(RESULTS_POLL_SECONDS)
        shown = read()
    return shown


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
    shown = _combo_settles_on(
        lambda: _field(editor, label, control_type="ComboBox").selected_text().strip(), value
    )
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
            f"No payment method named {method!r}; "
            + (f"this install has {available}." if available else "and its list could not be read.")
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
    return _combo_options(_field(editor, label, control_type="ComboBox"))


def _combo_options(combo: Any) -> list[str]:
    """List what a combo offers.

    Args:
        combo (Any): The combo box.

    Returns:
        list[str]: The option labels, in the order shown.
    """
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
    rows = search_view(win, TERMS_OF_PAYMENT_VIEW, method)
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


def choose_row(dialog: Any, row: table.Row) -> None:
    """Select a row in a picker dialog and confirm it (steps 2.12, 3.3).

    Args:
        dialog (WindowSpecification): The picker dialog.
        row (table.Row): The row to choose.
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
            # Under the tab, not merely below it: an invoice has dated fields
            # away to the right that would otherwise be read as the address.
            anchor = tab.rectangle()
            boxes = [
                c
                for c in editor.descendants(control_type="Edit")
                if not c.element_info.name
                and anchor.bottom <= c.rectangle().top <= anchor.bottom + ADDRESS_BOX_REACH
                and c.rectangle().left <= anchor.right
            ]
            addresses[name] = (
                max(boxes, key=lambda c: c.rectangle().width() * c.rectangle().height()).get_value()
                if boxes
                else ""
            )
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
    if shown.casefold() == wanted.casefold():
        return True

    # A cell too narrow for its value is cut short, and the ellipsis marking
    # that is not always read back -- "Northstar Office ..." comes back as
    # "Northstar Office" whenever the crop lands a pixel tighter. So a shorter
    # cell that begins the expected value counts, provided enough of it shows
    # to mean something. The caller compares several columns, so an opening
    # never decides a match on its own.
    prefix = shown.rstrip(". …")
    return (
        len(prefix) >= MIN_PREFIX_MATCH
        and len(prefix) < len(wanted)
        and wanted.casefold().startswith(prefix.casefold())
    )


def find_debtor(dialog: Any, term: str, **expected: str | None) -> table.Row | None:
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
        # Step 2.3 asks for manual review here; the instruction for this
        # install is to take the first instead. Rows that match on all five
        # fields describe the same customer entered twice, so either one
        # points at the same person -- but it is said out loud, because the
        # duplicate itself is worth someone's attention.
        print(
            f"note: {len(exact)} debtors match the document; using the first, "
            f"{[row.get('No.') for row in exact]}",
            flush=True,
        )
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


def vat_name(rate: float) -> str:
    """The name a VAT rate must carry to be reusable (step 3.5).

    Args:
        rate (float): The extracted VAT rate, e.g. 19.0.

    Returns:
        str: "VAT 19%" -- whole rates without a decimal part, since that is
            how the step writes them and how a person would type them.
    """
    rounded = round(rate, 2)
    figure = int(rounded) if rounded == int(rounded) else rounded
    return f"{VAT_NAME_PREFIX} {figure}%"


def percentage(text: str) -> float | None:
    """Read a rate out of what a field or a list cell shows.

    Args:
        text (str): Displayed value, e.g. "19.00 %" or "19%".

    Returns:
        float | None: The number, or None if there is none.
    """
    match = _NUMBER_RE.search(text.replace(",", "."))
    return float(match.group()) if match else None


def open_row(win: WindowSpecification, view: Any, row: table.Row, name: str) -> Any:
    """Open a list row's record, and return its editor.

    A drawn list has no rows to invoke, so the record is opened the way a
    person opens it: a double-click on the row. The editor it opens is named
    after the record.

    Args:
        win (WindowSpecification): The main window.
        view (Any): The list view holding the row.
        row (table.Row): The row to open.
        name (str): The record's name, which its tab takes.

    Returns:
        Any: The record's editor.

    Raises:
        LookupError: If the list pane has no window handle.
        TimeoutError: If no editor opens.
    """
    pane = _list_pane(view)
    handle = getattr(pane.element_info, "handle", None)
    if not handle:
        raise LookupError("The list has no window handle to click.")
    position = win32api.MAKELONG(int(pane.rectangle().width() * ROW_CLICK_X), int(row.y))
    with tracing.step(f"open {name!r}"):
        for message in (win32con.WM_LBUTTONDOWN, win32con.WM_LBUTTONUP, win32con.WM_LBUTTONDBLCLK, win32con.WM_LBUTTONUP):
            flags = win32con.MK_LBUTTON if message != win32con.WM_LBUTTONUP else 0
            win32gui.PostMessage(handle, message, flags, position)
        return wait_for_editor(win, f"^\\*?{re.escape(name)}$", name)


def find_vat(win: WindowSpecification, rate: float) -> table.Row | None:
    """Look up the VAT rate an item needs (steps 3.4-3.5).

    A rate is only reusable when everything about it agrees with the
    document: it is named for its percentage, its value *is* that
    percentage, and it is filed as the standard rate. A rate of the right name whose value
    or code says something else is a disagreement about tax, which is not
    something to resolve by picking one.

    Note that a rate with the right value under another name -- this install
    ships "MwSt. 19%" -- is not a match and not a conflict either: step 3.6
    creates the named one alongside it.

    Args:
        win (WindowSpecification): The main window.
        rate (float): The extracted VAT percentage.

    Returns:
        table.Row | None: The matching rate, or None if it has to be created.

    Raises:
        ManualReviewRequired: If a rate of that name disagrees about the
            rate or the e-invoice code, or several carry the name.
    """
    name = vat_name(rate)
    rows = search_view(win, VATS_VIEW, name)
    named = [row for row in rows if _cell_matches(row.get("Name"), name)]
    if not named:
        return None
    if len(named) > 1:
        raise ManualReviewRequired(f"{len(named)} VAT rates are called {name!r}", named)

    row = named[0]
    value = percentage(row.get("Value"))
    if value != rate:
        raise ManualReviewRequired(f"{name!r} is worth {row.get('Value')!r}, not {rate}%", [row])

    editor = open_row(win, open_list_view(win, VATS_VIEW), row, name)
    code = _field(editor, VAT_CODE_LABEL, control_type="ComboBox").selected_text()
    if code != STANDARD_VAT_CODE:
        raise ManualReviewRequired(f"{name!r} is coded {code!r}, not {STANDARD_VAT_CODE!r}", [row])
    return row


def create_vat(win: WindowSpecification, rate: float) -> Any:
    """Fill in a new VAT rate for `rate` (step 3.6).

    Leaves the editor unsaved for the caller to save, and does not touch the
    displayed Standard rate: this rate is one an imported document needs, not
    a new default for the install.

    Args:
        win (WindowSpecification): The main window.
        rate (float): The extracted VAT percentage.

    Returns:
        Any: The unsaved editor.

    Raises:
        LookupError: If the view has no create button.
        ValueError: If a field does not hold what we entered.
    """
    name = vat_name(rate)
    view = open_list_view(win, VATS_VIEW)
    buttons = [c for c in view.descendants(control_type="Button") if c.element_info.name == NEW_VAT_BUTTON]
    if not buttons:
        raise LookupError(f"No {NEW_VAT_BUTTON!r} button in the {VATS_VIEW!r} view.")

    with tracing.step(f"create the {name!r} rate"):
        tracing.point_at(buttons[0])
        buttons[0].iface_invoke.Invoke()
        editor = wait_for_editor(win, NEW_VAT_TAB_RE, "New TAX Rate")

        _set_text(editor, "Name", name)
        _set_text(editor, "Description", name)
        _set_text(editor, VAT_VALUE_LABEL, f"{rate:g}%")

    code = _field(editor, VAT_CODE_LABEL, control_type="ComboBox").selected_text()
    if code != STANDARD_VAT_CODE:
        raise ValueError(f"A new rate is coded {code!r}, expected {STANDARD_VAT_CODE!r}.")
    shown = percentage(_field(editor, VAT_VALUE_LABEL).get_value())
    if shown != rate:
        raise ValueError(f"Value reads {shown!r} after entering {rate}%.")
    tracing.point_at(editor, colour=tracing.CONFIRM)
    return editor


def create_product(
    win: WindowSpecification,
    sku: str,
    description: str,
    price: float,
    vat: str,
) -> Any:
    """Fill in a new product for an item line (steps 3.7-3.10).

    Opened from the toolbar's product button rather than the New panel's
    "New product" link: they run the same command, but the toolbar button
    answers Invoke while the link is a bare label whose click landed
    elsewhere when the window was part off-screen.

    The VAT rate has to exist before this runs -- the editor's VAT list is
    filled when it opens, so a rate created afterwards would not be offered.

    Leaves the editor unsaved for the caller to save. Category, GTIN,
    supplier code, allowance, the picture and the user-defined fields are
    left exactly as the editor proposes them.

    Args:
        win (WindowSpecification): The main window.
        sku (str): The extracted item number.
        description (str): The extracted item description, used for both the
            product's name and its description.
        price (float): The gross price, from `models.gross_price`.
        vat (str): The name of the VAT rate to select, e.g. "VAT 19%".

    Returns:
        Any: The unsaved editor.

    Raises:
        LookupError: If the toolbar has no product button.
        ValueError: If a field does not hold what we entered.
    """
    buttons = [c for c in win.descendants(control_type="Button") if c.element_info.name == NEW_PRODUCT_BUTTON]
    if not buttons:
        raise LookupError(f"No {NEW_PRODUCT_BUTTON!r} button in the toolbar.")

    with tracing.step(f"create the product {sku!r}"):
        tracing.point_at(buttons[0])
        buttons[0].iface_invoke.Invoke()
        editor = wait_for_editor(win, NEW_PRODUCT_TAB_RE, "New product")

        # 3.8: the item number, and the description under both its names.
        _set_text(editor, "Item Number", sku)
        _set_text(editor, "Name", description)
        _set_text(editor, "Description", description)

        # 3.9-3.10: the master price, nothing for cost or stock, and the rate.
        _set_text(editor, GROSS_PRICE_LABEL, f"{price:.2f}")
        _set_text(editor, COST_PRICE_LABEL, ZERO_AMOUNT)
        _set_text(editor, STOCK_LABEL, ZERO_AMOUNT)
        _set_combo(editor, "VAT", vat)

    shown = money(_field(editor, GROSS_PRICE_LABEL).get_value())
    if shown != round(price, 2):
        raise ValueError(f"{GROSS_PRICE_LABEL} reads {shown!r} after entering {price:.2f}.")
    tracing.point_at(editor, colour=tracing.CONFIRM)
    return editor


def money(text: str) -> float | None:
    """Read an amount out of a displayed value like "$297.50".

    Args:
        text (str): What the field shows.

    Returns:
        float | None: The amount, or None if there is none.
    """
    match = _NUMBER_RE.search(text.replace(",", ""))
    return float(match.group()) if match else None


def items_table(editor: Any) -> tuple[Any, Any]:
    """The document's item table: the pane to read, and the canvas to click.

    Found by its shape rather than by the label beside it -- the invoice
    editor has no "Items" caption at all, though its table is built the same
    way. That table is always a pane with a second pane inset a couple of
    pixels inside it: the outer one draws the table, the inner one takes the
    mouse. Clicks posted to the outer are ignored, which is why both come back.

    Args:
        editor (Any): An order or invoice editor.

    Returns:
        tuple[Any, Any]: The pane to capture, and the canvas to click.

    Raises:
        LookupError: If no such pair is there.
    """
    panes = [c for c in editor.descendants(control_type="Pane") if getattr(c.element_info, "handle", None)]
    pairs = [
        (outer, inner)
        for outer in panes
        for inner in panes
        if inner is not outer
        and inner.rectangle().left - outer.rectangle().left == TABLE_INSET
        and inner.rectangle().top - outer.rectangle().top == TABLE_INSET
    ]
    if not pairs:
        raise LookupError("No item table in this editor.")
    return max(pairs, key=lambda pair: pair[1].rectangle().width() * pair[1].rectangle().height())


def scroll_items(editor: Any, to_end: bool = False) -> None:
    """Scroll the item table sideways so the wanted columns are on screen.

    The table is wider than its pane, and a column that is scrolled out of
    view can be neither read nor clicked.

    Args:
        editor (Any): The order editor.
        to_end (bool): Scroll to the right-hand end instead of the left.
    """
    outer, canvas = items_table(editor)
    # Only one of the table's nested windows carries the scroll pattern, and
    # which one is not fixed -- a silently skipped scroll leaves the leftmost
    # columns (Qty., Item No.) off screen, where they can be neither read nor
    # clicked.
    for pane in (canvas, outer):
        try:
            pane.iface_scroll.SetScrollPercent(100.0 if to_end else 0.0, -1.0)
        except Exception:
            continue
        time.sleep(RESULTS_POLL_SECONDS)
        return
    # Neither pane scrolls, which means there is nothing to scroll: a
    # maximized window shows the whole table at once.


def read_items(editor: Any, to_end: bool = False) -> tuple[list[table.Row], dict[str, float], Any]:
    """Read the order's item lines.

    Args:
        editor (Any): The order editor.
        to_end (bool): Read the right-hand columns instead of the left-hand
            ones; the table is too wide to show both at once.

    Returns:
        tuple[list[table.Row], dict[str, float], Any]: The lines, where each
            column's header sits, and the capture they were read from.
    """
    scroll_items(editor, to_end)
    outer, _ = items_table(editor)
    picture = _grab(outer)
    return table.read(picture), table.header_positions(picture), picture


def set_item_cell(editor: Any, line: table.Row, column: str, value: str) -> None:
    """Type a value into one cell of an item line (steps 3.13, 3.15).

    A drawn table has no cell to address, so the cell is opened the way a
    person opens it -- a double-click at the column's x and the line's y --
    which puts a real text box over it. That box is then typed into and
    committed with Return, exactly as if it had been filled by hand.

    Args:
        editor (Any): The order editor.
        line (table.Row): The line to edit, from `read_items`.
        column (str): The column's header, e.g. "Qty".
        value (str): What to put in the cell.

    Raises:
        LookupError: If the column is not on screen.
        TimeoutError: If no cell editor opens.
        ValueError: If the cell does not end up holding `value`.
    """
    outer, canvas = items_table(editor)
    # The table is wider than its pane: a column may be off to the right, so
    # look at both ends before giving up on it. Column names come from reading
    # the header, so they are matched the same forgiving way as list columns
    # ("U.Price" comes back as "U.Prig" often enough).
    columns = {}
    for to_end in (False, True):
        _, found, _ = read_items(editor, to_end=to_end)
        columns = found
        position = _column_x(found, column)
        if position is not None:
            break
    else:
        position = None
    if position is None:
        raise LookupError(f"No {column!r} column in the item table; it shows {sorted(columns)}.")

    outer_rect, canvas_rect = outer.rectangle(), canvas.rectangle()
    left, top, _, _ = win32gui.GetWindowRect(canvas.element_info.handle)
    x = int(position + outer_rect.left - left)
    y = int(line.y + outer_rect.top - top)
    position = win32api.MAKELONG(x, y)

    def open_and_write() -> None:
        """Open the cell and put the value in it, from scratch."""
        handle = canvas.element_info.handle
        for message, flags in (
            (win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON),
            (win32con.WM_LBUTTONUP, 0),
            (win32con.WM_LBUTTONDBLCLK, win32con.MK_LBUTTON),
            (win32con.WM_LBUTTONUP, 0),
        ):
            win32gui.PostMessage(handle, message, flags, position)

        box = _wait_for_cell_editor(editor, canvas_rect)
        # Written, not typed. The cell's editor is created for the click and
        # ignores posted characters -- they land as raw text ("111000") or not
        # at all -- while a written value is picked up and recalculated. This
        # is the opposite of the debtor's boxes, which ignore a written value.
        box.iface_value.SetValue(value)
        win32gui.PostMessage(box.element_info.handle, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        win32gui.PostMessage(box.element_info.handle, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        time.sleep(RESULTS_POLL_SECONDS)

    def write_and_check() -> None:
        """Fill the cell, and read the table back to see that it took."""
        open_and_write()
        shown = _cell_now(editor, line, column)
        wanted = money(value)
        if shown is None or wanted is None or abs(shown) != abs(wanted):
            raise ValueError(f"{column} reads {shown!r} after entering {value!r}.")

    # The editor belongs to the click that made it: once it goes, so does the
    # handle we were writing through. And a value can land in the cell as raw
    # text without the app taking it -- the table still says "10" while the
    # line total ignores it. So the whole gesture is repeated until the table
    # itself shows the number: click, find, write, read back.
    with tracing.step(f"set {column} to {value}"):
        _retry(write_and_check)


def _cell_now(editor: Any, line: table.Row, column: str) -> float | None:
    """Re-read one cell of an item line as a number.

    Args:
        editor (Any): The order or invoice editor.
        line (table.Row): The line, for its position.
        column (str): The column to read.

    Returns:
        float | None: What the cell shows, or None if it cannot be read.
    """
    lines, _, _ = read_items(editor, to_end=True)
    if not lines:
        return None
    nearest = min(lines, key=lambda row: abs(row.y - line.y))
    return money(nearest.get(column))


def _column_x(columns: dict[str, float], name: str) -> float | None:
    """Find a column's x by name, ignoring case and punctuation.

    Args:
        columns (dict[str, float]): Header positions, as read from the table.
        name (str): The column wanted.

    Returns:
        float | None: Its x, or None when the table does not show it.
    """
    wanted = normalize(name)
    for header, x in columns.items():
        if normalize(header) == wanted:
            return x
    return None


def _wait_for_cell_editor(editor: Any, area: Any, timeout: int = DIALOG_TIMEOUT) -> Any:
    """Wait for the text box a cell opens when it is double-clicked.

    Args:
        editor (Any): The order editor.
        area (Any): The table's rectangle; the box appears inside it.
        timeout (int): Seconds to wait.

    Returns:
        Any: The cell's text box.

    Raises:
        TimeoutError: If no box appears.
    """
    deadline = time.monotonic() + timeout
    while True:
        boxes = [
            c
            for c in editor.descendants(control_type="Edit")
            if area.top <= c.rectangle().top and c.rectangle().bottom <= area.bottom
            and area.left <= c.rectangle().left and c.rectangle().right <= area.right
        ]
        if boxes:
            return boxes[0]
        if time.monotonic() >= deadline:
            raise TimeoutError("No cell editor opened for the item line.")
        time.sleep(RESULTS_POLL_SECONDS)


def maximize(win: WindowSpecification) -> None:
    """Make the window as large as the screen allows.

    The order's item table is wider than a windowed Fakturama can show: its
    Price column stays clipped however far the table is scrolled, and a value
    that cannot be seen cannot be checked. Maximizing is the difference
    between reading a line's price and guessing it.

    Args:
        win (WindowSpecification): The main window.
    """
    handle = win.element_info.handle
    placement = win32gui.GetWindowPlacement(handle)
    if placement[1] != win32con.SW_SHOWMAXIMIZED:
        win32gui.ShowWindow(handle, win32con.SW_MAXIMIZE)
        time.sleep(WINDOW_SETTLE_SECONDS)


def find_item_line(editor: Any, sku: str) -> table.Row | None:
    """The order line carrying `sku`, read from the item table.

    Args:
        editor (Any): The order editor.
        sku (str): The item number to look for.

    Returns:
        table.Row | None: The line, or None if the table does not show it.
    """
    lines, _, _ = read_items(editor)
    for line in lines:
        if _cell_matches(line.get(ITEM_NUMBER_COLUMN), sku):
            return line
    return None


def complete_item_line(
    editor: Any,
    sku: str,
    qty: float | None = None,
    discount: float | None = None,
    total: float | None = None,
) -> table.Row:
    """Enter a line's quantity and discount, and see that the line took them
    (steps 3.13-3.16).

    The line's own total is the only honest proof that a cell was accepted: a
    quantity or a discount can sit in its cell, read back correctly, and leave
    the total untouched, which means the app never took it. So the values are
    entered, the line is given time to recalculate, and where it still does
    not come to what it should they are entered again.

    Args:
        editor (Any): The order or invoice editor.
        sku (str): The line's item number.
        qty (float | None): The quantity the document gives.
        discount (float | None): The line's discount, if it has one.
        total (float | None): What the line should come to, which is what
            makes this checkable at all.

    Returns:
        table.Row: The line, as it reads once it has settled.

    Raises:
        ManualReviewRequired: If the line is not there, or never comes to
            `total` however often the values are entered.
    """
    filled = None
    for _ in range(LINE_ATTEMPTS):
        line = find_item_line(editor, sku)
        if line is None:
            raise ManualReviewRequired(f"No order line for {sku!r} to complete", [])

        if qty is not None:
            set_item_cell(editor, line, QTY_COLUMN, f"{qty:g}")
        if discount:
            set_item_cell(editor, line, LINE_DISCOUNT_COLUMN, f"{discount:g}")

        filled = settled_item_line(editor, sku, total)
        if filled is None:
            raise ManualReviewRequired(f"The line for {sku!r} disappeared while filling it", [])
        if total is None or money(filled.get(LINE_PRICE_COLUMN)) == round(total, 2):
            return filled
        print(
            f"line {sku}: came to {filled.get(LINE_PRICE_COLUMN)!r}, entering it again",
            flush=True,
        )

    raise ManualReviewRequired(
        f"The line for {sku!r} comes to {filled.get(LINE_PRICE_COLUMN)!r}, "
        f"expected {total:.2f}",
        [filled] if filled else [],
    )


def settled_item_line(editor: Any, sku: str, total: float | None, timeout: int = LINE_SETTLE_TIMEOUT) -> table.Row | None:
    """Read a line once the app has finished working out what it comes to.

    A quantity or a discount lands in the cell before the line's own total
    catches up, so reading straight afterwards can show the price the line had
    a moment ago -- for a quantity of three at 40.00 that is 40.00, the price
    of one, which looks exactly like a quantity that never took. Read until it
    agrees with what the line should come to, or until the time is up: a line
    that never gets there is a real disagreement, and the caller says so.

    Args:
        editor (Any): The order or invoice editor.
        sku (str): The line's item number.
        total (float | None): What the line should come to, if it is known.
        timeout (int): Seconds to let it settle.

    Returns:
        table.Row | None: The line, settled if it settles.
    """
    deadline = time.monotonic() + timeout
    line = item_line(editor, sku)
    while total is not None and time.monotonic() < deadline:
        if line is not None and money(line.get(LINE_PRICE_COLUMN)) == round(total, 2):
            return line
        time.sleep(RESULTS_POLL_SECONDS)
        line = item_line(editor, sku)
    return line


def item_line(editor: Any, sku: str) -> table.Row | None:
    """A line's values, with every column on screen at once.

    Args:
        editor (Any): The order editor.
        sku (str): The item number of the line wanted.

    Returns:
        table.Row | None: The line, or None if it is not there.
    """
    lines, _, _ = read_items(editor, to_end=True)
    for line in lines:
        if _cell_matches(line.get(ITEM_NUMBER_COLUMN), sku) or _cell_matches(line.get(NAME_COLUMN), sku):
            return line
    return None


def shown_date(iso_date: str) -> str:
    """An ISO date as Fakturama prints it in a list ("2026-07-14" -> "Jul 14, 2026").

    Args:
        iso_date (str): The date in ISO form.

    Returns:
        str: The same date, formatted the way the app displays it.
    """
    return date.fromisoformat(iso_date).strftime(DATE_DISPLAY_FORMAT)


def field_value(editor: Any, label: str, index: int = 0) -> str:
    """What a labelled box on a form currently holds.

    The reading half of `_set_text`, for confirming a form rather than
    filling it -- which is what the verification steps and the smoke run do.

    Args:
        editor (Any): An editor pane.
        label (str): The box's label.
        index (int): Which box, when one label covers several.

    Returns:
        str: The value, as the app displays it.
    """
    return _fields(editor, label)[index].get_value()


def field_values(editor: Any, label: str) -> list[str]:
    """What every box under one label holds, left to right.

    Args:
        editor (Any): An editor pane.
        label (str): The label they share, e.g. "First Name Last Name".

    Returns:
        list[str]: Their values.
    """
    return [field.get_value() for field in _fields(editor, label)]


def combo_value(editor: Any, label: str) -> str:
    """Which option a labelled combo is on.

    Args:
        editor (Any): An editor pane.
        label (str): The combo's label.

    Returns:
        str: The selected option.
    """
    return _field(editor, label, control_type="ComboBox").selected_text()


def price_mode(editor: Any) -> str:
    """Whether a document is priced Net or Gross.

    Args:
        editor (Any): An order or invoice editor.

    Returns:
        str: The header row's price mode.
    """
    return _price_mode_combo(editor).selected_text()


def count_editors(win: WindowSpecification, title_re: str) -> int:
    """How many editors of one kind are open.

    Args:
        win (WindowSpecification): The main window.
        title_re (str): Pattern matching their tabs.

    Returns:
        int: The number open.
    """
    return _tab_count(win, title_re)


def order_totals(editor: Any) -> dict[str, float | None]:
    """What the order says it comes to (step 4.3).

    Args:
        editor (Any): The order editor.

    Returns:
        dict[str, float | None]: Net, VAT and gross, as the editor shows them.
    """
    return {
        name: money(_field(editor, label).get_value())
        for name, label in ORDER_TOTAL_LABELS.items()
    }


def confirm_order_totals(
    editor: Any, net: float | None = None, vat: float | None = None, gross: float | None = None
) -> dict[str, float | None]:
    """Check the order's totals against the document's (step 4.3).

    Totals are the one number on the page that depends on everything else, so
    they are the check worth making before saving: if the lines, the VAT rate
    and the discounts are all right, these agree, and if any of them is wrong,
    these do not.

    Args:
        editor (Any): The order editor.
        net (float | None): The document's net total, if it gives one.
        vat (float | None): Its VAT total.
        gross (float | None): Its gross total.

    Returns:
        dict[str, float | None]: The totals the order shows.

    Raises:
        ManualReviewRequired: If a total the document gives disagrees.
    """
    shown = order_totals(editor)
    wanted = {"net": net, "vat": vat, "gross": gross}
    complaints = [
        f"{name} reads {shown[name]}, document says {value}"
        for name, value in wanted.items()
        if value is not None and shown[name] != round(value, 2)
    ]
    if complaints:
        raise ManualReviewRequired("The order's totals do not match: " + "; ".join(complaints), [])
    return shown


def confirm_order_charges(editor: Any) -> None:
    """Check the order carries no charges of its own (step 4.2).

    The document prices its lines and nothing else, so an order-level discount
    or a shipping charge here would be money the document never mentioned.
    They are Fakturama's defaults, so this confirms rather than sets them.

    Args:
        editor (Any): The order editor.

    Raises:
        ManualReviewRequired: If either has been given a value.
    """
    discount = _field(editor, DISCOUNT_LABEL).get_value()
    shipping = _field(editor, SHIPPING_LABEL, control_type="ComboBox").selected_text()
    charge = money(_fields(editor, SHIPPING_LABEL)[-1].get_value()) if _fields(editor, SHIPPING_LABEL) else None
    if percentage(discount) or shipping != FREE_SHIPPING or charge:
        raise ManualReviewRequired(
            f"The order carries charges the document does not: discount {discount!r}, "
            f"shipping {shipping!r} at {charge!r}",
            [],
        )


def save_document(win: WindowSpecification, editor: Any, title_re: str, what: str) -> str:
    """Save a document once, and return the number it was filed under.

    The number is read first: saving renames the tab to it, so afterwards
    there is no "New Order" or "New Invoice" left to look under.

    Args:
        win (WindowSpecification): The main window.
        editor (Any): The document's editor.
        title_re (str): Pattern matching its tab while unsaved.
        what (str): The editor's name, for messages.

    Returns:
        str: The document's No.
    """
    number = _field(editor, DOCUMENT_NUMBER_LABEL).get_value()
    save_editor(win, title_re, what)
    return number


def save_order(win: WindowSpecification, editor: Any) -> str:
    """Save the order once (step 4.4).

    Args:
        win (WindowSpecification): The main window.
        editor (Any): The order editor.

    Returns:
        str: The order's No.
    """
    return save_document(win, editor, NEW_ORDER_TAB_RE, "New Order")


def save_invoice(win: WindowSpecification, editor: Any) -> str:
    """Save the invoice once (step 5.4).

    Args:
        win (WindowSpecification): The main window.
        editor (Any): The invoice editor.

    Returns:
        str: The invoice's No.
    """
    return save_document(win, editor, NEW_INVOICE_TAB_RE, "New Invoice")


def _view_button(view: Any, name: str) -> Any | None:
    """One of a view's own window buttons, if it has it.

    Args:
        view (Any): A list view.
        name (str): "Maximize" or "Restore".

    Returns:
        Any | None: The button, or None.
    """
    buttons = [c for c in view.descendants(control_type="Button") if c.element_info.name == name]
    return buttons[0] if buttons else None


def close_view(win: WindowSpecification, name: str) -> None:
    """Close a list view, by the close button on its own tab.

    Args:
        win (WindowSpecification): The main window.
        name (str): The view's name, e.g. "Documents".

    Raises:
        TimeoutError: If the view does not go away.
    """
    tabs = [t for t in win.descendants(control_type="TabItem") if t.element_info.name == name]
    if not tabs:
        return
    tab = tabs[0]
    activate_tab(tab)

    folder = tab.parent()
    handle = getattr(folder.element_info, "handle", None)
    if not handle:
        raise LookupError(f"The tab folder holding {name!r} has no window handle.")
    left, top, _, _ = win32gui.GetWindowRect(handle)
    middle = tab.rectangle().mid_point()
    position = win32api.MAKELONG(middle.x - left, middle.y - top)
    # Middle-clicked, not aimed at the little close cross: anywhere on the tab
    # will do, where the cross has to be hit exactly and ignores a click that
    # lands a few pixels off it.
    win32gui.PostMessage(handle, win32con.WM_MBUTTONDOWN, win32con.MK_MBUTTON, position)
    win32gui.PostMessage(handle, win32con.WM_MBUTTONUP, 0, position)

    deadline = time.monotonic() + TAB_TIMEOUT
    while any(t.element_info.name == name for t in win.descendants(control_type="TabItem")):
        if time.monotonic() >= deadline:
            raise TimeoutError(f"The {name!r} view did not close.")
        time.sleep(0.1)


def show_whole_list(view: Any) -> None:
    """Clear a view's category filter, so its search covers everything.

    The Documents view files its rows under a little tree -- Invoices into
    paid and unpaid, Orders into shipped and not -- and whichever branch was
    last clicked keeps filtering the list afterwards. Left on "Invoices /
    unpaid", a search for an order number finds nothing at all, which reads
    exactly like a document that was never saved. The first node is the whole
    list; views without such a tree are unaffected.

    Args:
        view (Any): A list view.
    """
    items = view.descendants(control_type="TreeItem")
    if not items:
        return
    try:
        items[0].iface_selection_item.Select()
        time.sleep(RESULTS_POLL_SECONDS)
    except Exception:
        return


def search_view(win: WindowSpecification, name: str, term: str) -> list[table.Row]:
    """Search a Data list, reopening it if it comes back empty.

    A view holds the rows it was opened with, so anything saved since is in
    the database but not in the list. Nothing found is therefore ambiguous --
    it means "not there" or "not there yet" -- and the two have opposite
    consequences: one says create the record, the other says one exists
    already and a second would be a duplicate. Reopening settles it.

    Args:
        win (WindowSpecification): The main window.
        name (str): The view's name, e.g. "VATs".
        term (str): What to search for.

    Returns:
        list[table.Row]: What the list shows, after a reopen if it had nothing.
    """
    view = open_list_view(win, name)
    show_whole_list(view)
    rows = search_list(view, term)
    if rows:
        return rows

    view = refresh_view(win, name)
    show_whole_list(view)
    return search_list(view, term)


def refresh_view(win: WindowSpecification, name: str) -> Any:
    """Close a list view and open it again, so it reads the database afresh.

    A view holds the rows it was opened with. A document saved after that is
    in the database but not in the list, and searching for it finds nothing --
    which looks exactly like a save that did not happen. Reopening the view is
    what tells the two apart.

    Args:
        win (WindowSpecification): The main window.
        name (str): The view's name.

    Returns:
        Any: The reopened view.
    """
    with tracing.step(f"reopen the {name} list"):
        close_view(win, name)
        return open_list_view(win, name)


def widen_view(win: WindowSpecification, name: str) -> Any:
    """Give a list view the whole window, and return it.

    A view sharing the window with an editor is narrow enough that its columns
    are cut short -- "$678.30" arrives as "$6..." -- and a value that cannot
    be read cannot be checked. Views offer Maximize for exactly this, and a
    view already pushed aside offers Restore first.

    Args:
        win (WindowSpecification): The main window.
        name (str): The view's name, e.g. "Documents".

    Returns:
        Any: The view, as wide as it goes.
    """
    view = open_list_view(win, name)
    for label in ("Restore", "Maximize"):
        button = _view_button(view, label)
        if button:
            button.iface_invoke.Invoke()
            time.sleep(WINDOW_SETTLE_SECONDS)
            view = open_list_view(win, name)
    return view


def unwiden_view(win: WindowSpecification, name: str) -> None:
    """Give the window back to the editors.

    Args:
        win (WindowSpecification): The main window.
        name (str): The view's name.
    """
    button = _view_button(open_list_view(win, name), "Restore")
    if button:
        button.iface_invoke.Invoke()
        time.sleep(WINDOW_SETTLE_SECONDS)


def confirm_documents(win: WindowSpecification, expected: dict[str, dict[str, str | None]]) -> dict[str, table.Row]:
    """Check several documents in one look at the list (step 5.5).

    The invoice and the order it came from are checked together, because what
    the step is really asking is that raising and paying the invoice left the
    order alone: same reference, same total, still open.

    Args:
        win (WindowSpecification): The main window.
        expected (dict[str, dict[str, str | None]]): Document number -> the
            columns it should show, keyed as in `DOCUMENT_MATCH_COLUMNS`.

    Returns:
        dict[str, table.Row]: The row found for each number.

    Raises:
        ManualReviewRequired: If any of them is missing or disagrees.
    """
    widen_view(win, DOCUMENTS_VIEW)
    try:
        return {
            number: confirm_document_row(win, number, widen=False, **columns)
            for number, columns in expected.items()
        }
    finally:
        unwiden_view(win, DOCUMENTS_VIEW)


def confirm_document_row(
    win: WindowSpecification, number: str, widen: bool = True, **expected: str | None
) -> table.Row:
    """Find the saved order in Data > Documents and check it (step 4.5).

    Args:
        win (WindowSpecification): The main window.
        number (str): The order's number, as `save_order` returned it.
        widen (bool): Give the list the whole window first. On by default,
            because a list sharing the window truncates its columns to the
            point where even the document number does not match; pass False
            when the caller has already widened it.
        **expected (str | None): Columns to check, keyed as in
            `DOCUMENT_MATCH_COLUMNS` -- date, reference, state, total.

    Returns:
        table.Row: The row the list shows for this order.

    Raises:
        ManualReviewRequired: If the order is not listed once, or a column
            disagrees with what was entered.
    """
    widened = widen
    view = widen_view(win, DOCUMENTS_VIEW) if widen else open_list_view(win, DOCUMENTS_VIEW)
    try:
        show_whole_list(view)
        rows = [row for row in search_list(view, number) if _cell_matches(row.get(DOCUMENT_COLUMN), number)]
        if not rows:
            # Nothing found does not yet mean the document is not there: the
            # view holds the rows it was opened with, so anything saved since
            # is missing from it. Look again at a freshly opened one.
            # Widen again whatever the caller did: a reopened view comes back
            # at its ordinary size, and a narrow one truncates the document
            # number itself, so nothing would match however fresh it is.
            refresh_view(win, DOCUMENTS_VIEW)
            view = widen_view(win, DOCUMENTS_VIEW)
            widened = True
            show_whole_list(view)
            rows = [row for row in search_list(view, number) if _cell_matches(row.get(DOCUMENT_COLUMN), number)]
    finally:
        if widened:
            unwiden_view(win, DOCUMENTS_VIEW)
    if len(rows) != 1:
        raise ManualReviewRequired(f"{len(rows)} documents are numbered {number!r}", rows)

    row = rows[0]
    complaints = [
        f"{DOCUMENT_MATCH_COLUMNS[field]} reads {row.get(DOCUMENT_MATCH_COLUMNS[field])!r}, expected {value!r}"
        for field, value in expected.items()
        if value and field in DOCUMENT_MATCH_COLUMNS
        and not _column_agrees(row.get(DOCUMENT_MATCH_COLUMNS[field]), value)
    ]
    if complaints:
        raise ManualReviewRequired(f"The saved order {number!r} does not read back: " + "; ".join(complaints), [row])
    return row


def _column_agrees(shown: str, expected: str) -> bool:
    """Whether a listed cell says the same thing as `expected`.

    Amounts are compared as numbers, because the list writes them with the
    currency it is configured for ("$678.30") while the document gives a bare
    figure. Everything else is compared as text, allowing for truncation.

    Args:
        shown (str): What the cell shows.
        expected (str): What it should say.

    Returns:
        bool: True when they agree.
    """
    if _AMOUNT_RE.fullmatch(expected.strip()):
        return money(shown) == money(expected)
    return _cell_matches(shown, expected)


def create_follow_up(win: WindowSpecification, editor: Any, kind: str = FOLLOW_UP_INVOICE) -> Any:
    """Raise a follow-up document from the saved order (steps 4.6-4.7).

    From the order's own "Create a follow-up document" area, not the toolbar:
    the toolbar's Invoice button starts an unrelated invoice, while this one
    carries the order's lines, its customer and the link back to it.

    Args:
        win (WindowSpecification): The main window.
        editor (Any): The saved order's editor.
        kind (str): Which follow-up to raise, e.g. "Invoice".

    Returns:
        Any: The new document's editor.

    Raises:
        LookupError: If the order editor offers no such follow-up.
        TimeoutError: If no editor opens.
    """
    buttons = [
        c
        for c in editor.descendants(control_type="Button")
        if c.element_info.name == kind and c.rectangle().height() > FOLLOW_UP_MIN_HEIGHT
    ]
    if not buttons:
        raise LookupError(f"No {kind!r} follow-up on this order.")

    with tracing.step(f"create the follow-up {kind}"):
        tracing.point_at(buttons[0])
        buttons[0].iface_invoke.Invoke()
        follow_up = wait_for_editor(win, FOLLOW_UP_TAB_RE.format(kind=kind), f"New {kind}")
    tracing.point_at(follow_up, colour=tracing.CONFIRM)
    return follow_up


def document_summary(editor: Any) -> dict[str, Any]:
    """Everything about a document that a follow-up should inherit.

    Read from the editor rather than from the database, because what the
    steps ask to confirm is what a person would see on screen.

    Args:
        editor (Any): An order or invoice editor.

    Returns:
        dict[str, Any]: Reference, addresses, VAT mode, totals and lines.
    """
    lines, _, _ = read_items(editor, to_end=True)
    return {
        "reference": _field(editor, CUSTOMER_REFERENCE_LABEL).get_value(),
        "addresses": order_addresses(editor),
        "vat_mode": _field_combo(editor, "VAT").selected_text(),
        "totals": order_totals(editor),
        "lines": [
            tuple(line.get(column) for column in INHERITED_LINE_COLUMNS)
            for line in lines
            if line.get(ITEM_NUMBER_COLUMN)
        ],
    }


def confirm_invoice_from_order(win: WindowSpecification, number: str) -> dict[str, Any]:
    """Confirm the invoice carries the order's content (step 5.1).

    The follow-up is supposed to copy the order, so the order is what it is
    checked against -- reading both editors and comparing, rather than
    re-checking each field against the document a second time. The invoice's
    own No., Invoice Date and Service date are Fakturama's to propose and are
    left alone, so they are not compared.

    Args:
        win (WindowSpecification): The main window.
        number (str): The saved order's number, which is its tab's name.

    Returns:
        dict[str, Any]: What the invoice holds.

    Raises:
        ManualReviewRequired: If anything the follow-up should have inherited
            differs from the order.
    """
    order_editor = activate_editor(win, rf"^\*?{re.escape(number)}$", number)
    from_order = document_summary(order_editor)
    from_order["order_date"] = _field(order_editor, "Date").get_value()

    invoice_editor = activate_editor(win, NEW_INVOICE_TAB_RE, "New Invoice")
    from_invoice = document_summary(invoice_editor)
    from_invoice["order_date"] = _field(invoice_editor, ORDER_DATE_LABEL).get_value()

    complaints = [
        f"{name} reads {from_invoice[name]!r}, the order has {value!r}"
        for name, value in from_order.items()
        if from_invoice[name] != value
    ]
    if complaints:
        raise ManualReviewRequired(
            f"The invoice does not carry {number}'s content: " + "; ".join(complaints), []
        )
    return from_invoice


def _payment_combo(editor: Any) -> Any:
    """The invoice's payment method combo.

    It carries no accessible name, so it is found by what it stands next to:
    the "paid" checkbox, on the same line of the form.

    Args:
        editor (Any): The invoice editor.

    Returns:
        Any: The combo box.

    Raises:
        LookupError: If the row is not there.
    """
    boxes = [c for c in editor.descendants(control_type="CheckBox") if c.element_info.name == PAID_LABEL]
    if not boxes:
        raise LookupError(f"No {PAID_LABEL!r} checkbox in the invoice editor.")
    anchor = boxes[0].rectangle()
    beside = [
        c
        for c in editor.descendants(control_type="ComboBox")
        if c.rectangle().left >= anchor.right
        and abs(c.rectangle().mid_point().y - anchor.mid_point().y) <= LINE_TOLERANCE
    ]
    if not beside:
        raise LookupError(f"No payment method combo beside {PAID_LABEL!r}.")
    return min(beside, key=lambda c: c.rectangle().left)


def set_invoice_payment(editor: Any, method: str) -> str:
    """Set or confirm the invoice's payment method (step 5.2).

    The follow-up usually arrives with the debtor's method already on it, so
    this is normally a confirmation. Where it is not, the method has to be one
    Fakturama offers -- inventing one here would be a payment term nobody
    agreed to, so a missing method stops the run instead.

    Args:
        editor (Any): The invoice editor.
        method (str): The extracted payment method.

    Returns:
        str: The method the invoice ends up on.

    Raises:
        PaymentMethodUnavailable: If the invoice cannot offer that method.
        ValueError: If the combo does not end up on it.
    """
    combo = _payment_combo(editor)
    if combo.selected_text() == method:
        return method

    # Some of these combos do not put their items in the accessibility tree
    # at all, so an empty list is "could not read", not "has none". Where the
    # list is readable it decides; otherwise the attempt to select does.
    options = _combo_options(combo)
    if options and method not in options:
        raise PaymentMethodUnavailable(method, options)
    with tracing.step(f"set the invoice's payment method to {method!r}"):
        tracing.point_at(combo)
        try:
            _select_option(combo, method)
        except Exception as exc:
            raise PaymentMethodUnavailable(method, options) from exc

    shown = _payment_combo(editor).selected_text()
    if shown != method:
        raise ValueError(f"The invoice's payment method is {shown!r} after selecting {method!r}.")
    return shown


def set_invoice_paid(
    editor: Any, paid: bool, payment_date: str | date | None = None, value: float | None = None
) -> dict[str, Any]:
    """Record whether the invoice has been paid (step 5.3).

    Ticking "paid" changes the row: the due-days and pay-until fields it shows
    while unpaid are replaced by the date the payment was made and its value.
    So the box is ticked first, and only then are the other two touched.

    An unpaid invoice is left alone entirely -- no date, no value. The
    document says nothing about either, and a plausible-looking guess in a
    payment record is worse than a blank.

    Args:
        editor (Any): The invoice editor.
        paid (bool): Whether the document says it has been paid.
        payment_date (str | date | None): When, if it has.
        value (float | None): How much; the whole invoice, per the step.

    Returns:
        dict[str, Any]: What the payment row holds afterwards.

    Raises:
        ValueError: If the date or the value does not read back.
    """
    box = [c for c in editor.descendants(control_type="CheckBox") if c.element_info.name == PAID_LABEL][0]
    if bool(box.get_toggle_state()) != paid:
        with tracing.step(f"{'tick' if paid else 'clear'} {PAID_LABEL!r}"):
            tracing.point_at(box)
            box.iface_toggle.Toggle()
            time.sleep(RESULTS_POLL_SECONDS)
    if not paid:
        return {"paid": False}

    if payment_date:
        wanted = date.fromisoformat(payment_date) if isinstance(payment_date, str) else payment_date
        with tracing.step(f"set the payment date to {wanted.isoformat()}"):
            _type_date(_payment_date_field(editor), wanted)
        shown = _payment_date_field(editor).get_value()
        if datetime.strptime(shown, DATE_DISPLAY_FORMAT).date() != wanted:
            raise ValueError(f"The payment date reads {shown!r} after entering {wanted.isoformat()}.")

    if value is not None and money(_field(editor, PAID_VALUE_LABEL).get_value()) != round(value, 2):
        with tracing.step(f"set the paid value to {value:.2f}"):
            _set_text(editor, PAID_VALUE_LABEL, f"{value:.2f}")
        if money(_field(editor, PAID_VALUE_LABEL).get_value()) != round(value, 2):
            raise ValueError(
                f"{PAID_VALUE_LABEL} reads {_field(editor, PAID_VALUE_LABEL).get_value()!r}, expected {value:.2f}."
            )

    return {
        "paid": True,
        "method": _payment_combo(editor).selected_text(),
        "date": _payment_date_field(editor).get_value(),
        "value": _field(editor, PAID_VALUE_LABEL).get_value(),
    }


def _payment_date_field(editor: Any) -> Any:
    """The date beside "at" on a paid invoice's payment row.

    Args:
        editor (Any): The invoice editor.

    Returns:
        Any: The date widget.

    Raises:
        LookupError: If the row does not show one, which it does not until
            the invoice is marked paid.
    """
    return _field(editor, PAID_AT_LABEL)


def invoice_payment(editor: Any) -> dict[str, Any]:
    """What an invoice's payment row records.

    Args:
        editor (Any): The invoice editor.

    Returns:
        dict[str, Any]: Whether it is paid, and by what method, when and for
            how much. The date and value only exist once it is marked paid.
    """
    box = [c for c in editor.descendants(control_type="CheckBox") if c.element_info.name == PAID_LABEL][0]
    paid = bool(box.get_toggle_state())
    record: dict[str, Any] = {"paid": paid, "method": _payment_combo(editor).selected_text()}
    if paid:
        record["date"] = _payment_date_field(editor).get_value()
        record["value"] = _field(editor, PAID_VALUE_LABEL).get_value()
    return record


def confirm_saved_invoice(
    win: WindowSpecification,
    number: str,
    method: str | None = None,
    paid: bool = False,
    payment_date: str | date | None = None,
    value: float | None = None,
) -> dict[str, Any]:
    """Confirm what the saved invoice actually holds (step 5.6).

    Reopened only if it has to be: while its editor is still open, that editor
    is the saved record and reading it proves the same thing. Once it has been
    closed, the record is fetched back from Data > Documents -- which is the
    stronger check of the two, since it comes from the database rather than
    from the form we typed into.

    Args:
        win (WindowSpecification): The main window.
        number (str): The invoice's number.
        method (str | None): The payment method it should carry.
        paid (bool): Whether it should be marked paid.
        payment_date (str | date | None): The date it should record.
        value (float | None): The amount it should record.

    Returns:
        dict[str, Any]: The payment row as saved.

    Raises:
        ManualReviewRequired: If the saved invoice does not hold what was
            entered, or cannot be found at all.
    """
    try:
        editor = activate_editor(win, rf"^\*?{re.escape(number)}$", number)
    except LookupError:
        with tracing.step(f"reopen {number}"):
            view = widen_view(win, DOCUMENTS_VIEW)
            rows = [r for r in search_list(view, number) if _cell_matches(r.get(DOCUMENT_COLUMN), number)]
            if not rows:
                raise ManualReviewRequired(f"The saved invoice {number!r} is not in {DOCUMENTS_VIEW}", [])
            editor = open_row(win, view, rows[0], number)
            unwiden_view(win, DOCUMENTS_VIEW)

    record = invoice_payment(editor)
    wanted: dict[str, Any] = {"paid": paid}
    if method:
        wanted["method"] = method
    if paid and payment_date:
        wanted["date"] = shown_date(payment_date) if isinstance(payment_date, str) else payment_date.strftime(DATE_DISPLAY_FORMAT)
    if paid and value is not None:
        wanted["value"] = value

    complaints = []
    for name, expected in wanted.items():
        shown = record.get(name)
        agrees = money(shown) == round(expected, 2) if name == "value" else shown == expected
        if not agrees:
            complaints.append(f"{name} reads {shown!r}, expected {expected!r}")
    if complaints:
        raise ManualReviewRequired(f"The saved invoice {number!r} does not hold: " + "; ".join(complaints), [])
    return record


def dismiss_dialog(dialog: Any, button: str) -> None:
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


def wait_for_dialog(win: WindowSpecification, title: str, timeout: int = DIALOG_TIMEOUT) -> Any:
    """Wait for a modal dialog of the given title and return it.

    Args:
        win (WindowSpecification): The main window, to identify the process.
        title (str): Exact window caption.
        timeout (int): Seconds to wait.

    Returns:
        Any: The dialog. It is a lazy specification, so it re-finds itself on
            every use -- which is what keeps it valid across the redraws a
            dialog goes through, at the cost of lookups that need retrying.

    Raises:
        TimeoutError: If it does not open in time.
    """
    _, process_id = win32process.GetWindowThreadProcessId(win.element_info.handle)
    deadline = time.monotonic() + timeout
    while True:
        handle = _find_dialog(process_id, title)
        if handle:
            dialog = Application(backend="uia").connect(handle=handle).window(handle=handle)
            # On screen is not the same as ready: a dialog can exist for a
            # moment with nothing published inside it, and returning then
            # leaves the caller looking for a search box that is not there yet.
            if _retry(lambda: bool(dialog.children())):
                return dialog
        if time.monotonic() >= deadline:
            raise TimeoutError(f"No {title!r} dialog appeared within {timeout}s.")
        time.sleep(0.25)


def open_picker(
    editor: Any, win: WindowSpecification, label: str, title: str, icons: int
) -> Any:
    """Open a picker dialog from the icons beside a section of the editor.

    Both pickers work the same way: a little column of icons sits beside the
    section -- Addresses, Items -- and the topmost one chooses something that
    already exists, while the green + below it starts a new record. We take
    the topmost by position rather than by recognising an icon, and insist on
    finding exactly as many as that section is known to have, so a changed
    toolbar fails a lookup instead of quietly clicking "new".

    Args:
        editor (Any): The order editor.
        label (str): The section's label, e.g. "Items".
        win (WindowSpecification): The main window, to find the dialog.
        title (str): The dialog's caption.
        icons (int): How many icons that section has.

    Returns:
        WindowSpecification: The open dialog.

    Raises:
        LookupError: If the section, or its expected icons, are not there.
        TimeoutError: If the dialog does not open.
    """
    labels = [c for c in editor.descendants(control_type="Text") if c.element_info.name == label]
    if not labels:
        raise LookupError(f"No {label!r} label in the order editor.")

    def buttons() -> list[Any]:
        return sorted(labels[0].parent().children(control_type="Image"), key=lambda c: c.rectangle().top)

    found = buttons()
    if len(found) != icons:
        raise LookupError(
            f"Expected {icons} icons beside {label}, found {len(found)}; "
            "refusing to guess which one picks an existing record."
        )

    with tracing.step(f"open {title!r} from the {label} row"):
        tracing.point_at(found[0])
        # An editor built moments ago can take the click before its icon is
        # listening, and then nothing happens at all -- so the click is
        # repeated, against a freshly located icon, until the dialog shows up.
        for attempt in range(DIALOG_ATTEMPTS):
            _post_click(found[0])
            try:
                dialog = wait_for_dialog(win, title, timeout=DIALOG_RETRY_SECONDS)
                break
            except TimeoutError:
                if attempt == DIALOG_ATTEMPTS - 1:
                    raise
                found = buttons()
    tracing.point_at(dialog, colour=tracing.CONFIRM)
    return dialog


def open_address_selector(editor: Any, win: WindowSpecification) -> Any:
    """Open Select the address from the order's Addresses row (step 2.1).

    Args:
        editor (Any): The editor pane from `open_new_order`.
        win (WindowSpecification): The main window, to find the dialog.

    Returns:
        WindowSpecification: The open dialog.
    """
    return open_picker(editor, win, ADDRESSES_LABEL, ADDRESS_DIALOG_TITLE, ADDRESS_ICON_COUNT)


def open_product_selector(editor: Any, win: WindowSpecification) -> Any:
    """Open Select a product from the order's Items table (step 3.2).

    Args:
        editor (Any): The order editor.
        win (WindowSpecification): The main window, to find the dialog.

    Returns:
        WindowSpecification: The open dialog.
    """
    return open_picker(editor, win, ITEMS_LABEL, PRODUCT_DIALOG_TITLE, ITEM_ICON_COUNT)


def select_product(editor: Any, win: WindowSpecification, sku: str) -> bool:
    """Give the order's next line a product, by item number (steps 3.2-3.3).

    The product picker does its own selecting: narrow its search to a single
    product and it adds that product to the order and closes itself, with no
    row to click and no OK to press. So the dialog closing is the success
    signal, and a dialog still standing means the search did not identify one
    product -- either none, or several to choose between.

    (The address picker, which looks identical, does not behave this way: it
    waits to be told. The difference is why this cannot reuse `search_list`,
    which insists on reading its search box back afterwards.)

    Args:
        editor (Any): The order editor.
        win (WindowSpecification): The main window.
        sku (str): The extracted item number.

    Returns:
        bool: True if the product was added to the order; False if no product
            carries that item number and it has to be created first.

    Raises:
        ManualReviewRequired: If the item number matches several products.
    """
    dialog = open_product_selector(editor, win)
    handle = dialog.element_info.handle

    with tracing.step(f"search products for {sku!r}"):
        # The whole term at once, never character by character. This picker
        # filters on every keystroke and accepts the moment one row is left,
        # so typing "CHR-ERG-01" walks through prefixes that match something
        # else entirely -- "CH" alone picked "Sicherung 1A" and put it on the
        # order. Writing the value skips those intermediate states.
        _retry(lambda: _search_box(dialog).iface_value.SetValue(sku))
        settled = _wait_for_pick(handle, dialog)
    if settled is None:
        return True

    rows = table.read(settled)
    exact = [row for row in rows if _cell_matches(row.get(SKU_COLUMN), sku)]
    if len(exact) > 1:
        dismiss_dialog(dialog, "Cancel")
        raise ManualReviewRequired(f"{len(exact)} products carry the item number {sku!r}", exact)
    if exact:
        choose_row(dialog, exact[0])
        return True

    dismiss_dialog(dialog, "Cancel")
    return False


def _wait_for_pick(handle: int, dialog: Any, timeout: int = RESULTS_TIMEOUT) -> Image.Image | None:
    """Wait for the picker to either choose for us or settle on a list.

    Args:
        handle (int): The dialog's window.
        dialog (Any): The dialog.
        timeout (int): Seconds to wait for the list to settle.

    Returns:
        Image.Image | None: The settled list, or None if the dialog picked a
            product and closed itself.
    """
    deadline = time.monotonic() + PICK_SETTLE_SECONDS
    while time.monotonic() < deadline:
        if not (win32gui.IsWindow(handle) and win32gui.IsWindowVisible(handle)):
            return None
        time.sleep(RESULTS_POLL_SECONDS)
    return wait_for_results(dialog, timeout)


def _grab(ctrl: Any) -> Image.Image:
    """Capture a control, retrying a refusal that means "not now".

    Args:
        ctrl (Any): A control with a native window handle.

    Returns:
        Image.Image: The control as an RGB image.
    """
    return _retry(lambda: _grab_once(ctrl))


def _grab_once(ctrl: Any) -> Image.Image:
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


def _find_waking(win: WindowSpecification, find: Any, attempts: int = UIA_ATTEMPTS) -> Any:
    """Look for a control, waking the window between tries.

    A plain retry does not help when the window has published a partial tree:
    the same query keeps missing the same control. Giving the window the focus
    is what makes it publish the rest, so that is done between attempts.

    Args:
        win (WindowSpecification): The main window.
        find (Any): A callable returning the control.
        attempts (int): How many times to look.

    Returns:
        Any: Whatever `find` returned.

    Raises:
        Exception: Whatever the last attempt raised.
    """
    for attempt in range(attempts):
        try:
            return find()
        except Exception:
            if attempt == attempts - 1:
                raise
            wake(win, force=True)


def _retry(action: Any, attempts: int = UIA_ATTEMPTS) -> Any:
    """Repeat a UI Automation call that failed for a passing reason.

    Resolving a control while the application is mid-redraw fails with a COM
    error ("An event was unable to invoke any of the subscribers") that means
    nothing except "not now" -- a moment later the same call succeeds. Only
    the last failure is allowed to escape.

    Args:
        action (Any): A callable taking no arguments.
        attempts (int): How many times to try it.

    Returns:
        Any: Whatever `action` returned.
    """
    for attempt in range(attempts):
        try:
            return action()
        except Exception:
            if attempt == attempts - 1:
                raise
            time.sleep(UIA_RETRY_SECONDS)


def _search_box(container: Any) -> Any:
    """Return a dialog's or list view's search box, retrying transient failures.

    Args:
        container (Any): A dialog or list view holding a searchable list.

    Returns:
        Any: The search Edit.
    """
    return _retry(lambda: _locate_search_box(container))


def _locate_search_box(container: Any) -> Any:
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
    """Return the pane holding the result rows, retrying transient failures.

    Args:
        container (Any): A dialog or list view holding a searchable list.

    Returns:
        Any: The results pane.
    """
    return _retry(lambda: _locate_list_pane(container))


def _locate_list_pane(container: Any) -> Any:
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
            _retry(lambda: _search_box(container).iface_value.SetValue(term))
            shown = _retry(lambda: _search_box(container).get_value())
            if shown == term:
                break
            if attempt == SEARCH_ATTEMPTS - 1:
                raise ValueError(f"Search box holds {shown!r} after entering {term!r}.")
            time.sleep(RESULTS_POLL_SECONDS)
        settled = wait_for_results(container)
    tracing.point_at(_list_pane(container), colour=tracing.CONFIRM)
    return table.read(settled)
