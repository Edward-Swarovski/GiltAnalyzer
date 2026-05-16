from __future__ import annotations

from datetime import date
from decimal import Decimal

from app.models import GiltMarketRow


def normalize_market_row(raw: dict[str, str], *, source_name: str) -> GiltMarketRow:
    """Normalize a raw gilt row into the canonical model.

    The caller must supply `coupon_rate` in percentage points, so `"1.5"` means
    1.5% and `"0.125"` means 0.125%.
    """
    imported_yield = raw.get("yield")

    return GiltMarketRow(
        isin=raw["isin"].strip(),
        gilt_name=raw["gilt_name"].strip(),
        coupon_rate=Decimal(raw["coupon_rate"]),
        maturity_date=date.fromisoformat(raw["maturity_date"]),
        imported_price=Decimal(raw["price"]) if raw.get("price") else None,
        imported_yield=Decimal(imported_yield) if imported_yield else None,
        valuation_date=date.fromisoformat(raw["valuation_date"]),
        source_name=source_name,
    )
