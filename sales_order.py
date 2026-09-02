"""Extracting the TJM Labs "Sales Order Input" layout into a `SalesOrder`.

This is the document-specific layer: the label->field map below, and the rules
for walking its label/value grid and items table. A differently laid-out
document needs its own map (or a proper invoice2data YAML template); the
geometry in `layout.py` and the OCR in `ocr.py` stay as they are.
"""

from typing import Any

import ocr
from layout import MAX_DIST_RATIO, Y_TOL_RATIO, TextBox, cluster_rows, nearest_label, normalize
from models import SalesOrder


LABEL_FIELDS = {
    "EXTERNAL REFERENCE": "external_reference",
    "ORDER DATE": "order_date",
    "CUSTOMER ID": "customer_id",
    "CURRENCY": "currency",
    "COMPANY": "company",
    "CONTACT NAME": "contact_name",
    "CUSTOMER ALIAS": "customer_alias",
    "EMAIL": "email",
    "PHONE": "phone",
    "BILLING ADDRESS": "billing_address",
    "DELIVERY ADDRESS": "delivery_address",
    "PAYMENT METHOD": "payment_method",
    "PAID STATUS": "paid_status",
    "PAYMENT DATE": "payment_date",
}
#: Multi-line values: collected until the next label or section instead of
#: being read from a single row.
ADDRESS_FIELDS = {"billing_address", "delivery_address"}
SECTION_HEADERS = {"ORDER", "CUSTOMER AND CONTACT", "ADDRESSES", "PAYMENT", "ITEMS"}
TOTALS_LABELS = {"NET TOTAL": "net_total", "VAT TOTAL": "vat_total", "GROSS TOTAL": "gross_total"}
#: Items table columns, left to right, matched positionally to the header row.
ITEM_COLUMNS = ["sku", "description", "qty", "unit", "unit_price", "discount_pct", "vat_pct", "line_total"]

LABEL_BY_NORM = {normalize(label): field for label, field in LABEL_FIELDS.items()}
SECTION_BY_NORM = {normalize(h) for h in SECTION_HEADERS}
TOTALS_BY_NORM = {normalize(label): field for label, field in TOTALS_LABELS.items()}

# Render resolution is a genuine trade-off for this layout, so we OCR twice:
#   2x  - correct word spacing ("Friedrichstrasse 88", "+49 30 5550 1420"),
#         but small table cells (single-digit Qty, "0%") fall below the text
#         detector's sensitivity and are dropped silently rather than misread.
#   4x  - recovers those cells, but the recognizer starts losing inter-word
#         spaces ("Friedrichstrasse88", "+493055501420").
# 3x was measured too and fixes neither fully.
HEADER_SCALE = 2
ITEMS_SCALE = 4

#: Below this many matched header labels we assume the layout/quality is wrong
#: rather than returning a silently-empty SalesOrder.
MIN_MATCHED_FIELDS = 5


class UnrecognizedLayoutError(ValueError):
    """Raised when too few labels match to trust the result."""


def _is_section(row: list[TextBox]) -> bool:
    return len(row) == 1 and normalize(row[0].text) in SECTION_BY_NORM


def _parse_header_fields(rows: list[list[TextBox]], max_dist: float) -> tuple[dict[str, Any], int]:
    """Walk the label/value grid above the items table.

    Args:
        rows (list[list[TextBox]]): Clustered rows for the page.
        max_dist (float): Column tolerance for label matching.

    Returns:
        tuple[dict[str, Any], int]: Field values, and the row index of "ITEMS"
            (or the row count if the section is absent).
    """
    # Box-by-box rather than whole-row: the two side-by-side label/value
    # columns don't share a vertical rhythm, so row clustering sometimes
    # merges one column's value with the other column's next label.
    data: dict[str, Any] = {}
    pending: dict[str, tuple[float, float]] = {}
    i = 0
    while i < len(rows):
        row = rows[i]

        if _is_section(row):
            if normalize(row[0].text) == "ITEMS":
                return data, i
            pending = {}
            i += 1
            continue

        new_labels = {
            LABEL_BY_NORM[normalize(b.text)]: (b.x, b.y) for b in row if normalize(b.text) in LABEL_BY_NORM
        }
        addr_labels = {f: xy for f, xy in new_labels.items() if f in ADDRESS_FIELDS}
        pending.update({f: xy for f, xy in new_labels.items() if f not in ADDRESS_FIELDS})

        if addr_labels:
            collected: dict[str, list[str]] = {k: [] for k in addr_labels}
            i += 1
            while i < len(rows):
                nrow = rows[i]
                if _is_section(nrow) or any(normalize(b.text) in LABEL_BY_NORM for b in nrow):
                    break
                for b in nrow:
                    field = nearest_label(addr_labels, b, max_dist)
                    if field:
                        collected[field].append(b.text)
                i += 1
            for field, parts in collected.items():
                data[field] = ", ".join(parts) if parts else None
            continue

        remaining = dict(pending)
        for b in row:
            if normalize(b.text) in LABEL_BY_NORM:
                continue
            field = nearest_label(remaining, b, max_dist)
            if field:
                data[field] = b.text
                del remaining[field]
        pending = remaining
        i += 1
    return data, len(rows)


def _parse_items(
    rows: list[list[TextBox]], items_row: int, max_dist: float
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Read the items table and the totals row that ends it.

    Args:
        rows (list[list[TextBox]]): Clustered rows for the page.
        items_row (int): Index of the "ITEMS" section header row.
        max_dist (float): Column tolerance for matching totals to their labels.

    Returns:
        tuple[list[dict[str, Any]], dict[str, Any]]: Line items, and totals.
    """
    if items_row + 1 >= len(rows):
        return [], {}
    header_row = sorted(rows[items_row + 1], key=lambda b: b.x)
    columns = list(zip(ITEM_COLUMNS, [b.x for b in header_row]))

    items: list[dict[str, Any]] = []
    totals: dict[str, Any] = {}
    i = items_row + 2
    while i < len(rows):
        row = rows[i]
        normed = [normalize(b.text) for b in row]
        if normed and all(u in TOTALS_BY_NORM for u in normed):
            label_x = {TOTALS_BY_NORM[u]: (b.x, b.y) for u, b in zip(normed, row)}
            i += 1
            if i < len(rows):
                for b in rows[i]:
                    field = nearest_label(label_x, b, max_dist)
                    if field:
                        totals[field] = b.text
            break
        item: dict[str, Any] = {}
        for b in row:
            field = min(columns, key=lambda c: abs(c[1] - b.x))[0]
            item[field] = b.text
        items.append(item)
        i += 1
    return items, totals


def _rows_for(path: str, scale: int) -> tuple[list[list[TextBox]], float]:
    """OCR `path` at `scale` and cluster it into rows, with the column tolerance."""
    page = ocr.read_page(path, scale)
    return cluster_rows(page.boxes, page.width * Y_TOL_RATIO), page.width * MAX_DIST_RATIO


def extract_sales_order(path: str) -> SalesOrder:
    """OCR `path` (first page, if a PDF) and parse it into a `SalesOrder`.

    Args:
        path (str): Path to a PDF or image of the sales-order document.

    Returns:
        SalesOrder: Parsed fields; anything not found on the page stays None.

    Raises:
        UnrecognizedLayoutError: If too few labels matched to trust the result.
    """
    header_rows, header_dist = _rows_for(path, HEADER_SCALE)
    data, _ = _parse_header_fields(header_rows, header_dist)

    matched = sum(1 for v in data.values() if v is not None)
    if matched < MIN_MATCHED_FIELDS:
        raise UnrecognizedLayoutError(
            f"Matched only {matched} of {len(LABEL_FIELDS)} field labels in {path!r}. "
            "Either this is not the expected sales-order layout, or the source is too "
            "low-resolution to OCR reliably (a ~385px-wide screenshot fails this way; "
            "prefer the vector PDF)."
        )

    # Second pass purely for the items table (see HEADER_SCALE/ITEMS_SCALE).
    item_rows, item_dist = _rows_for(path, ITEMS_SCALE)
    _, items_row = _parse_header_fields(item_rows, item_dist)
    items, totals = _parse_items(item_rows, items_row, item_dist)

    data["items"] = items
    data.update(totals)
    return SalesOrder(**data)
