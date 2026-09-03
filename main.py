"""Read a sales-order document and enter it into a running Fakturama.

The seam between the two halves of the project: `sales_order` turns the file
into a `SalesOrder`, `fakturama` drives the application, and neither imports
the other. Which extracted field goes into which UI field is decided here.

Covers the order header (steps 1.3-1.7) and the debtor (steps 2.1-2.13):
select the customer if Fakturama already knows them, otherwise create them --
along with the payment method, if that is missing too -- and come back to the
order to select them. The order editor is left open and unsaved throughout, so
the products still have to be added before anything is filed.

Usage:
    python main.py [document]           # default: invoice.pdf
    python main.py --cursor             # watch it work, on-screen pointer
"""

import argparse

import fakturama
import tracing
from models import PostalAddress, SalesOrder, gross_price, line_total, parse_postal_address, same_address
from sales_order import extract_sales_order


#: Document to read when none is named on the command line.
DEFAULT_DOCUMENT = "invoice.pdf"


def enter_header(editor: object, order: SalesOrder) -> None:
    """Fill the New Order editor's header from `order` (steps 1.4-1.7).

    A field the document did not yield is left at whatever the editor
    proposed, and said so on stdout: for an import that is reviewed before
    saving, an obvious gap beats a guess.

    Args:
        editor (object): The editor pane from `fakturama.open_new_order`.
        order (SalesOrder): The extracted order.
    """
    # 1.4: the proposed No. is Fakturama's to allocate -- we never touch it.
    if order.order_date:
        fakturama.set_date(editor, order.order_date)
    else:
        print("no order date extracted; leaving the proposed Date", flush=True)

    if order.external_reference:
        fakturama.set_customer_reference(editor, order.external_reference)
    else:
        print("no external reference extracted; leaving Cust.Ref. empty", flush=True)

    fakturama.set_price_mode(editor)
    fakturama.set_vat_mode(editor)


def debtor_criteria(order: SalesOrder, billing: PostalAddress) -> dict[str, str | None]:
    """The values a listed debtor must show to be this order's customer (step 2.3).

    Args:
        order (SalesOrder): The extracted order.
        billing (PostalAddress): Its billing address, split up.

    Returns:
        dict[str, str | None]: Criteria for `fakturama.find_debtor`.
    """
    first_name, last_name = fakturama.split_contact_name(order.contact_name or "")
    return {
        "company": order.company,
        "first_name": first_name,
        "last_name": last_name,
        "postcode": billing.postcode,
        "city": billing.city,
    }


def search_term(order: SalesOrder) -> str:
    """What to type into the address search (step 2.2).

    Args:
        order (SalesOrder): The extracted order.

    Returns:
        str: The company, or the contact's name when there is no company.

    Raises:
        ValueError: If the document names neither.
    """
    term = order.company or order.contact_name
    if not term:
        raise ValueError("The document gives neither a company nor a contact name to search for.")
    return term


def create_debtor(win: object, order: SalesOrder, billing: PostalAddress) -> object:
    """Create and save the document's customer (steps 2.5-2.11).

    Args:
        win (object): The main window.
        order (SalesOrder): The extracted order.
        billing (PostalAddress): Its billing address, split up.

    Returns:
        object: The saved debtor's editor.
    """
    editor = fakturama.open_new_debtor(win)

    # 2.6: identity. The proposed Customer ID stays as Fakturama allocated it.
    fakturama.set_debtor_identity(editor, company=order.company, contact_name=order.contact_name)

    # 2.7-2.8: the billing address, and the roles it plays. One address covers
    # both roles only when the document says they are the same place.
    delivers_here = same_address(order.billing_address, order.delivery_address)
    fakturama.set_main_address(
        editor,
        street=billing.street,
        postcode=billing.postcode,
        city=billing.city,
        country=billing.country,
        email=order.email,
        phone=order.phone,
    )
    fakturama.set_address_roles(editor, invoice=True, delivery=delivers_here)
    if order.delivery_address and not delivers_here:
        print(
            "note: the document's delivery address differs from the billing one and is "
            f"not entered anywhere: {order.delivery_address!r}",
            flush=True,
        )

    # 2.9: alias, no discount, prices net.
    fakturama.set_debtor_miscellaneous(editor, alias=order.customer_alias)

    # 2.10: the payment method, creating it first if this install lacks it.
    if order.payment_method:
        set_payment_method(win, editor, order.payment_method)
    else:
        print("no payment method extracted; leaving the debtor's default", flush=True)

    # 2.11: save, once.
    fakturama.save_editor(win, fakturama.NEW_DEBTOR_TAB_RE, "New Debtor")
    return editor


def set_payment_method(win: object, editor: object, method: str) -> None:
    """Select the document's payment method, creating it if need be (step 2.10).

    Args:
        win (object): The main window.
        editor (object): The debtor editor.
        method (str): The extracted payment method.
    """
    try:
        fakturama.set_debtor_payment(editor, method)
        return
    except fakturama.PaymentMethodUnavailable as unavailable:
        print(f"{unavailable}; creating it", flush=True)

    # 2.10.1-2.10.2: it may exist without being offered to this debtor yet.
    if not fakturama.find_payment_method(win, method):
        # 2.10.3-2.10.6: fill it in and save it.
        fakturama.create_payment_method(win, method)
        fakturama.save_editor(win, fakturama.NEW_TERM_TAB_RE, "New Term of Payment")

    editor = fakturama.activate_editor(win, fakturama.NEW_DEBTOR_TAB_RE, "New Debtor")
    fakturama.set_debtor_payment(editor, method)


def select_debtor(win: object, order_editor: object, order: SalesOrder, billing: PostalAddress) -> object:
    """Give the order its customer (steps 2.1-2.4, 2.12-2.13).

    Looks the customer up from the order first, and only creates one when the
    address list has nothing that matches. Either way the order ends up
    pointing at a saved debtor, with the addresses it filled in checked
    against the document.

    Args:
        win (object): The main window.
        order_editor (object): The order editor.
        order (SalesOrder): The extracted order.
        billing (PostalAddress): Its billing address, split up.

    Returns:
        object: The order editor, back in front.

    Raises:
        fakturama.ManualReviewRequired: If the list is ambiguous, or the
            addresses the order picked up do not match the document.
    """
    term = search_term(order)
    criteria = debtor_criteria(order, billing)

    dialog = fakturama.open_address_selector(order_editor, win)
    debtor = fakturama.find_debtor(dialog, term, **criteria)
    if not debtor:
        # 2.5-2.11: not there, so create them -- with the order left open.
        fakturama.dismiss_dialog(dialog, "Cancel")
        create_debtor(win, order, billing)

        # 2.12: back to the order, and pick the debtor we just saved.
        order_editor = fakturama.activate_editor(win, fakturama.NEW_ORDER_TAB_RE, "New Order")
        dialog = fakturama.open_address_selector(order_editor, win)
        debtor = fakturama.find_debtor(dialog, term, **criteria)
        if not debtor:
            fakturama.dismiss_dialog(dialog, "Cancel")
            raise fakturama.ManualReviewRequired(
                f"The debtor was saved but does not come back when searching for {term!r}", []
            )

    print(f"debtor: {debtor.cells}", flush=True)
    fakturama.choose_row(dialog, debtor)

    # 2.4 / 2.13: the order should now carry the document's address.
    order_editor = fakturama.activate_editor(win, fakturama.NEW_ORDER_TAB_RE, "New Order")
    expected = [order.company, billing.street, billing.postcode, billing.city, billing.country]
    filled = fakturama.confirm_order_addresses(
        order_editor, {fakturama.ROLE_INVOICE: [part for part in expected if part]}
    )
    for role, text in filled.items():
        print(f"{role}: {text!r}", flush=True)
    return order_editor


def select_products(win: object, order_editor: object, order: SalesOrder) -> object:
    """Give every item line its product, in source order (steps 3.1-3.3).

    Args:
        win (object): The main window.
        order_editor (object): The order editor.
        order (SalesOrder): The extracted order.

    Returns:
        object: The order editor, back in front.

    Raises:
        fakturama.ManualReviewRequired: If a line has no SKU to search for, if
            the product list is ambiguous, or if the product does not exist --
            creating one is steps 3.4-3.12, which are not built yet.
    """
    for position, item in enumerate(order.items, start=1):
        if not item.sku:
            raise fakturama.ManualReviewRequired(f"Item {position} has no SKU to search for", [])

        if not fakturama.select_product(order_editor, win, item.sku):
            # 3.4-3.11: no such product, so make one -- the order stays open.
            create_product(win, item)

            # 3.12: back to the order and pick the product we just saved.
            order_editor = fakturama.activate_editor(win, fakturama.NEW_ORDER_TAB_RE, "New Order")
            if not fakturama.select_product(order_editor, win, item.sku):
                raise fakturama.ManualReviewRequired(
                    f"Product {item.sku!r} was saved but the picker still does not offer it", []
                )
        # 3.13-3.16: quantity, discount, and the checks on what it comes to.
        order_editor = fakturama.activate_editor(win, fakturama.NEW_ORDER_TAB_RE, "New Order")
        complete_line(win, order_editor, item)
        print(f"item {position}: {item.sku} added and completed", flush=True)
    return order_editor


def ensure_vat(win: object, percentage: float) -> str:
    """Make sure the VAT rate an item needs exists (steps 3.4-3.6).

    Args:
        win (object): The main window.
        percentage (float): The extracted VAT percentage.

    Returns:
        str: The rate's name, for selecting in the product editor.
    """
    name = fakturama.vat_name(percentage)
    if fakturama.find_vat(win, percentage):
        return name

    fakturama.create_vat(win, percentage)
    fakturama.save_editor(win, fakturama.NEW_VAT_TAB_RE, "New TAX Rate")
    print(f"created the {name} rate", flush=True)
    return name


def create_product(win: object, item: object) -> None:
    """Create the product an item line needs (steps 3.4-3.11).

    Args:
        win (object): The main window.
        item (object): The extracted line item.

    Raises:
        fakturama.ManualReviewRequired: If the line gives no price or VAT to
            work out the product's price from.
    """
    if item.unit_price is None or item.vat_pct is None:
        raise fakturama.ManualReviewRequired(
            f"Item {item.sku!r} has no unit price or VAT percentage to price a product with", []
        )

    # 3.4-3.7: the rate first, so the product editor offers it.
    vat = ensure_vat(win, item.vat_pct)

    fakturama.create_product(
        win,
        sku=item.sku,
        description=item.description or item.sku,
        price=gross_price(item.unit_price, item.vat_pct),
        vat=vat,
    )
    # 3.11: save, once.
    fakturama.save_editor(win, fakturama.NEW_PRODUCT_TAB_RE, "New product")
    print(f"created product {item.sku}", flush=True)


def complete_line(win: object, order_editor: object, item: object) -> None:
    """Fill in and check the order line for `item` (steps 3.13-3.16).

    Args:
        win (object): The main window.
        order_editor (object): The order editor.
        item (object): The extracted line item.

    Raises:
        fakturama.ManualReviewRequired: If the line cannot be found, if what
            the product brought with it disagrees with the document, or if the
            line total does not come to what the document says it should.
    """
    line = fakturama.find_item_line(order_editor, item.sku)
    if line is None:
        raise fakturama.ManualReviewRequired(f"No order line for {item.sku!r} to complete", [])

    # 3.13 and 3.15: the quantity and the discount this transaction was given.
    if item.qty is not None:
        fakturama.set_item_cell(order_editor, line, fakturama.QTY_COLUMN, f"{item.qty:g}")
    if item.discount_pct:
        fakturama.set_item_cell(order_editor, line, fakturama.LINE_DISCOUNT_COLUMN, f"{item.discount_pct:g}")

    # 3.14 and 3.16: what the product brought with it, and what it comes to.
    filled = fakturama.item_line(order_editor, item.sku)
    if filled is None:
        raise fakturama.ManualReviewRequired(f"The line for {item.sku!r} disappeared while filling it", [])
    check_line(filled, item)


def check_line(line: object, item: object) -> None:
    """Check a filled-in line against the document (steps 3.14, 3.16).

    Args:
        line (object): The line as the order shows it.
        item (object): The extracted line item.

    Raises:
        fakturama.ManualReviewRequired: On any disagreement.
    """
    complaints = []
    unit = fakturama.money(line.get(fakturama.UNIT_PRICE_COLUMN))
    if item.unit_price is not None and unit != round(item.unit_price, 2):
        complaints.append(f"unit price reads {line.get(fakturama.UNIT_PRICE_COLUMN)!r}, document says {item.unit_price}")

    vat = fakturama.percentage(line.get(fakturama.LINE_VAT_COLUMN))
    if item.vat_pct is not None and vat != item.vat_pct:
        complaints.append(f"VAT reads {line.get(fakturama.LINE_VAT_COLUMN)!r}, document says {item.vat_pct}%")

    # The discount is shown as a negative adjustment, so it is compared by size.
    discount = fakturama.percentage(line.get(fakturama.LINE_DISCOUNT_COLUMN))
    if item.discount_pct is not None and discount is not None and abs(discount) != item.discount_pct:
        complaints.append(f"discount reads {line.get(fakturama.LINE_DISCOUNT_COLUMN)!r}, document says {item.discount_pct}%")

    price = fakturama.money(line.get(fakturama.LINE_PRICE_COLUMN))
    expected = line_total(item)
    if expected is not None and price != expected:
        complaints.append(f"line price reads {line.get(fakturama.LINE_PRICE_COLUMN)!r}, expected {expected:.2f}")

    if complaints:
        raise fakturama.ManualReviewRequired(f"The line for {item.sku!r} does not match: " + "; ".join(complaints), [])


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Read the command line.

    Args:
        argv (list[str] | None): Arguments to parse; defaults to sys.argv[1:].

    Returns:
        argparse.Namespace: `document` and `cursor`, the latter None when the
            flag was not given, so FAKTURAMA_TRACE still decides.
    """
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "document",
        nargs="?",
        default=DEFAULT_DOCUMENT,
        help=f"sales-order PDF or image to read (default: {DEFAULT_DOCUMENT})",
    )
    parser.add_argument(
        "--cursor",
        action=argparse.BooleanOptionalAction,
        default=None,
        help=(
            "draw the on-screen pointer showing which control is being driven; "
            "off by default, and settable with FAKTURAMA_TRACE=visual"
        ),
    )
    return parser.parse_args(argv)


def main(path: str) -> None:
    """Extract `path` and enter it into Fakturama.

    Args:
        path (str): The sales-order PDF or image to read.
    """
    with tracing.step(f"extract {path}"):
        order = extract_sales_order(path)
    print(order.model_dump_json(indent=2))

    window = fakturama.connect()
    # The item table's right-hand columns are only readable at full width.
    fakturama.maximize(window)
    order_editor = fakturama.open_new_order(window)
    enter_header(order_editor, order)

    billing = parse_postal_address(order.billing_address, order.company)
    order_editor = select_debtor(window, order_editor, order, billing)
    order_editor = select_products(window, order_editor, order)
    print("header, customer and products entered; the New Order tab is left open and unsaved.", flush=True)


if __name__ == "__main__":
    args = parse_args()
    tracing.configure(visual=args.cursor)
    try:
        main(args.document)
    except fakturama.ManualReviewRequired as needs_review:
        raise SystemExit(f"Stopped for manual review: {needs_review}")
    finally:
        tracing.stop()
