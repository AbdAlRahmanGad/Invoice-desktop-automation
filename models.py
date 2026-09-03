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


_POSTCODE_CITY_RE = re.compile(r"^(?P<postcode>\d{4,6})\s+(?P<city>.+)$")
_ADDRESS_NOISE_RE = re.compile(r"[^0-9A-Za-z]+")


class PostalAddress(BaseModel):
    """A postal address split into the parts a form asks for."""

    street: str | None = None
    postcode: str | None = None
    city: str | None = None
    country: str | None = None


def parse_postal_address(text: str | None, company: str | None = None) -> PostalAddress:
    """Split a one-line address into street, postcode, city and country.

    Extraction yields an address as the document prints it, one line per comma
    ("Northstar Office GmbH, Friedrichstrasse 88, 10117 Berlin, Germany"),
    because that is what the page shows. A form wants the pieces separately,
    so we anchor on the one line that is unambiguous -- the postcode and city
    -- and read outwards: what precedes it is the street, what follows it is
    the country.

    Args:
        text (str | None): The address as extracted.
        company (str | None): Company name to drop if the address repeats it
            as its first line, which invoices usually do.

    Returns:
        PostalAddress: The parts that could be identified; the rest stay None.
            An address with no recognisable postcode line keeps everything in
            `street`, so nothing is silently dropped.
    """
    if not text:
        return PostalAddress()

    parts = [part.strip() for part in text.split(",") if part.strip()]
    if company and parts and parts[0].casefold() == company.casefold():
        parts = parts[1:]

    for index, part in enumerate(parts):
        match = _POSTCODE_CITY_RE.match(part)
        if match:
            return PostalAddress(
                street=", ".join(parts[:index]) or None,
                postcode=match.group("postcode"),
                city=match.group("city"),
                country=parts[index + 1] if index + 1 < len(parts) else None,
            )
    return PostalAddress(street=", ".join(parts) or None)


def same_address(first: str | None, second: str | None) -> bool:
    """Whether two extracted addresses describe the same place.

    Compared loosely on purpose: the two are OCR'd from different corners of
    the same page, so they differ in case, spacing and punctuation far more
    often than in meaning. Anything that is not a letter or a digit is dropped
    before comparing.

    Args:
        first (str | None): One address as extracted.
        second (str | None): The other.

    Returns:
        bool: True when both are present and match once normalised.
    """
    if not first or not second:
        return False
    return _ADDRESS_NOISE_RE.sub("", first).casefold() == _ADDRESS_NOISE_RE.sub("", second).casefold()


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
