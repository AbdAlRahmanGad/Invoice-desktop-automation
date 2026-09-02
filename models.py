"""Schema for an extracted sales order.

Pure data: no OCR, no geometry, no I/O.
"""

import re
from typing import Any

from pydantic import BaseModel, Field, field_validator


_NUMERIC_RE = re.compile(r"[-+]?\d*\.?\d+")


def _to_number(value: Any) -> float | None:
    """Pull a number out of an OCR string ("EUR678.30" -> 678.30, "10%" -> 10.0)."""
    if not isinstance(value, str):
        return value
    # "1,234.50" becomes "1234.50" but European "1.234,50" style not handled
    match = _NUMERIC_RE.search(value.replace(",", ""))
    return float(match.group()) if match else None


class LineItem(BaseModel):
    sku: str | None = None
    description: str | None = None
    qty: float | None = None
    unit: str | None = None
    unit_price: float | None = None
    discount_pct: float | None = None
    vat_pct: float | None = None
    line_total: float | None = None

    @field_validator("qty", "unit_price", "discount_pct", "vat_pct", "line_total", mode="before")
    @classmethod
    def _parse_number(cls, v: Any) -> float | None:
        return _to_number(v)


class SalesOrder(BaseModel):
    external_reference: str | None = None
    order_date: str | None = None
    customer_id: str | None = None
    currency: str | None = None
    company: str | None = None
    contact_name: str | None = None
    customer_alias: str | None = None
    email: str | None = None
    phone: str | None = None
    billing_address: str | None = None
    delivery_address: str | None = None
    payment_method: str | None = None
    paid_status: str | None = None
    payment_date: str | None = None
    items: list[LineItem] = Field(default_factory=list)
    net_total: float | None = None
    vat_total: float | None = None
    gross_total: float | None = None

    @field_validator("net_total", "vat_total", "gross_total", mode="before")
    @classmethod
    def _parse_number(cls, v: Any) -> float | None:
        return _to_number(v)
