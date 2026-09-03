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
from models import PostalAddress, SalesOrder, parse_postal_address, same_address
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
    fakturama.choose_address(dialog, debtor)

    # 2.4 / 2.13: the order should now carry the document's address.
    order_editor = fakturama.activate_editor(win, fakturama.NEW_ORDER_TAB_RE, "New Order")
    expected = [order.company, billing.street, billing.postcode, billing.city, billing.country]
    filled = fakturama.confirm_order_addresses(
        order_editor, {fakturama.ROLE_INVOICE: [part for part in expected if part]}
    )
    for role, text in filled.items():
        print(f"{role}: {text!r}", flush=True)
    return order_editor


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
    order_editor = fakturama.open_new_order(window)
    enter_header(order_editor, order)

    billing = parse_postal_address(order.billing_address, order.company)
    select_debtor(window, order_editor, order, billing)
    print("header and customer entered; the New Order tab is left open and unsaved.", flush=True)


if __name__ == "__main__":
    args = parse_args()
    tracing.configure(visual=args.cursor)
    try:
        main(args.document)
    except fakturama.ManualReviewRequired as needs_review:
        raise SystemExit(f"Stopped for manual review: {needs_review}")
    finally:
        tracing.stop()
