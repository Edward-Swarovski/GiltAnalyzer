from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal


@dataclass(frozen=True)
class GiltMarketRow:
    """Canonical market row.

    `coupon_rate` is stored in percentage points, not decimal form:
    `Decimal("1.5")` means 1.5%, not 150% and not 0.015%.
    """

    isin: str
    gilt_name: str
    coupon_rate: Decimal
    maturity_date: date
    imported_price: Decimal | None
    imported_yield: Decimal | None
    valuation_date: date
    source_name: str


@dataclass(frozen=True)
class GiltUserInput:
    isin: str
    nominal_amount: Decimal
    override_price: Decimal | None = None
    override_yield: Decimal | None = None


@dataclass(frozen=True)
class GiltPriceQuote:
    """External quote row where coupon_rate is also stored in percentage points."""

    epic: str
    gilt_name: str
    coupon_rate: Decimal
    maturity_date: date
    clean_price: Decimal
    yield_to_maturity: Decimal


@dataclass(frozen=True)
class GiltRetailQuote:
    """DMO retail purchase/sale quote row.

    `coupon_rate` is in percentage points, same convention as `GiltMarketRow`.
    """

    isin: str
    gilt_name: str
    coupon_rate: Decimal
    purchase_clean_price: Decimal
    purchase_dirty_price: Decimal
    sale_clean_price: Decimal
    sale_dirty_price: Decimal
    redemption_date: date
    data_date: date
