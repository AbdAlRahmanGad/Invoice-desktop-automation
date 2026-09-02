"""Read a sales-order document and enter it into a running Fakturama.

The seam between the two halves of the project: `sales_order` turns the file
into a `SalesOrder`, `fakturama` drives the application, and neither imports
the other. Which extracted field goes into which UI field is decided here.

Covers the header of the order (steps 1.3-1.7). The editor is deliberately
left open and unsaved afterwards (step 1.8), because the debtor, payment
method, VAT rate and products still have to be resolved into it.

Usage:
    python main.py [document]           # default: invoice.pdf
    python main.py --cursor             # watch it work, on-screen pointer
"""

import argparse

import fakturama
import tracing
from models import SalesOrder
from sales_order import extract_sales_order


#: Document to read when none is named on the command line.
DEFAULT_DOCUMENT = "invoice.pdf"


def enter_header(editor: object, order: SalesOrder) -> None:
    """Fill the New Order editor's header from `order`.

    A field the document did not yield is left at whatever the editor
    proposed, and said so on stdout: for an import that is reviewed before
    saving, an obvious gap beats a guess.

    Args:
        editor (object): The editor pane from `fakturama.open_new_order`.
        order (SalesOrder): The extracted order.
    """
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
    editor = fakturama.open_new_order(window)
    enter_header(editor, order)
    print("header entered; the New Order tab is left open and unsaved.", flush=True)


if __name__ == "__main__":
    args = parse_args()
    tracing.configure(visual=args.cursor)
    try:
        main(args.document)
    finally:
        tracing.stop()
