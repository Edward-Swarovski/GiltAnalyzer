from __future__ import annotations

from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path

import pandas as pd

from app.models import GiltRetailQuote


_PLAUSIBLE_YIELD_MIN = Decimal("-0.05")
_PLAUSIBLE_YIELD_MAX = Decimal("0.20")

# Normalised header names (lowercased, whitespace/newlines collapsed) mapped to
# canonical keys used internally. This makes the parser robust to minor DMO
# formatting changes such as added newlines or punctuation shifts.
_HEADER_MAP = {
    "isin code": "isin",
    "isin": "isin",
    "gilt name": "gilt_name",
    "stock": "gilt_name",
    "purchase clean price": "purchase_clean_price",
    "purchase dirty price": "purchase_dirty_price",
    "sale clean price": "sale_clean_price",
    "sale dirty price": "sale_dirty_price",
    "redemption date": "redemption_date",
}

_REQUIRED_KEYS = {
    "isin",
    "gilt_name",
    "purchase_clean_price",
    "purchase_dirty_price",
    "sale_clean_price",
    "sale_dirty_price",
    "redemption_date",
}


def _normalise_header(value: str) -> str:
    import re
    return re.sub(r"[\s\(\)£\n]+", " ", value).strip().lower()


def parse_d10b_xls(path: Path) -> list[GiltRetailQuote]:
    raw = pd.read_excel(path, header=None)
    data_date = _parse_data_date(raw.iloc[0, 0])

    # Locate the header row: first row whose first non-null cell maps to "isin"
    header_row_idx = None
    for i, row in raw.iterrows():
        cell = str(row.iloc[0]) if not pd.isna(row.iloc[0]) else ""
        if _HEADER_MAP.get(_normalise_header(cell)) == "isin":
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError(f"Cannot find header row in D10B file: {path}")

    raw_headers = [str(v) if not pd.isna(v) else "" for v in raw.iloc[header_row_idx]]
    col: dict[str, int] = {}
    for idx, raw_h in enumerate(raw_headers):
        canonical = _HEADER_MAP.get(_normalise_header(raw_h))
        if canonical and canonical not in col:
            col[canonical] = idx

    missing = _REQUIRED_KEYS - col.keys()
    if missing:
        raise ValueError(
            f"D10B file is missing expected columns: {sorted(missing)}"
        )

    quotes: list[GiltRetailQuote] = []
    for row in raw.iloc[header_row_idx + 1:].itertuples(index=False):
        if pd.isna(row[col["isin"]]) or pd.isna(row[col["gilt_name"]]):
            continue
        isin_raw = str(row[col["isin"]]).strip()
        if not _is_valid_isin(isin_raw):
            continue
        gilt_name = str(row[col["gilt_name"]]).strip()
        coupon_rate = _parse_coupon_from_name(gilt_name)
        quotes.append(
            GiltRetailQuote(
                isin=isin_raw,
                gilt_name=gilt_name,
                coupon_rate=coupon_rate,
                purchase_clean_price=Decimal(str(row[col["purchase_clean_price"]])),
                purchase_dirty_price=Decimal(str(row[col["purchase_dirty_price"]])),
                sale_clean_price=Decimal(str(row[col["sale_clean_price"]])),
                sale_dirty_price=Decimal(str(row[col["sale_dirty_price"]])),
                redemption_date=_coerce_date(row[col["redemption_date"]]),
                data_date=data_date,
            )
        )
    return quotes


def approximate_retail_ask_yield(
    *,
    coupon_rate: Decimal,
    sale_dirty_price: Decimal,
    redemption_date: date,
    valuation_date: date,
) -> Decimal:
    """Approximate annual redemption yield implied by the DMO sale dirty price.

    Uses a bisection solver over a simplified semi-annual coupon cashflow model.

    Day count: days / 365.25 — this is a simplification. The exact DMO
    convention is Actual/Actual (ICMA). Results are labelled 'Approx' in the
    workbook accordingly; use QuantLib or exact coupon schedules for
    official-convention precision.
    """
    cashflows = _future_cashflows(coupon_rate, redemption_date, valuation_date)
    low = Decimal("-0.99")
    high = Decimal("1.00")
    for _ in range(120):
        mid = (low + high) / 2
        pv = sum(
            amount / ((Decimal("1") + mid / 2) ** (Decimal("2") * years))
            for years, amount in cashflows
        )
        if pv > sale_dirty_price:
            low = mid
        else:
            high = mid

    result = (low + high) / 2 * 100
    if not (_PLAUSIBLE_YIELD_MIN * 100 <= result <= _PLAUSIBLE_YIELD_MAX * 100):
        raise ValueError(
            f"Bisection yield {result:.4f}% is outside the plausible range "
            f"[{_PLAUSIBLE_YIELD_MIN * 100}%, {_PLAUSIBLE_YIELD_MAX * 100}%]. "
            f"Check inputs: coupon_rate={coupon_rate}, "
            f"sale_dirty_price={sale_dirty_price}, "
            f"redemption_date={redemption_date}, "
            f"valuation_date={valuation_date}."
        )
    return result


def _future_cashflows(
    coupon_rate: Decimal,
    redemption_date: date,
    valuation_date: date,
) -> list[tuple[Decimal, Decimal]]:
    coupon_amount = coupon_rate / 2
    coupon_dates: list[date] = []
    current = redemption_date
    while current > valuation_date:
        coupon_dates.append(current)
        current = _subtract_six_months(current)
    coupon_dates.sort()
    cashflows: list[tuple[Decimal, Decimal]] = []
    for coupon_date in coupon_dates:
        years = Decimal((coupon_date - valuation_date).days) / Decimal("365.25")
        amount = coupon_amount + (
            Decimal("100") if coupon_date == redemption_date else Decimal("0")
        )
        cashflows.append((years, amount))
    return cashflows


def _subtract_six_months(value: date) -> date:
    month = value.month - 6
    year = value.year
    if month <= 0:
        month += 12
        year -= 1
    day = min(value.day, _days_in_month(year, month))
    return date(year, month, day)


def _days_in_month(year: int, month: int) -> int:
    if month == 12:
        return 31
    return (date(year, month + 1, 1) - timedelta(days=1)).day


def _parse_data_date(value: object) -> date:
    return datetime.strptime(
        str(value).replace("Data Date: ", ""), "%d-%b-%Y"
    ).date()


def _coerce_date(value: object) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return pd.Timestamp(value).date()


def _is_valid_isin(value: str) -> bool:
    import re
    return bool(re.fullmatch(r"[A-Z]{2}[A-Z0-9]{10}", value))


def _parse_coupon_from_name(gilt_name: str) -> Decimal:
    """Extract coupon rate in percentage points from a gilt name string."""
    import re
    coupon_text = gilt_name.split("%", maxsplit=1)[0].strip()
    fraction_values = {
        "½": Decimal("0.5"),
        "¼": Decimal("0.25"),
        "¾": Decimal("0.75"),
        "⅛": Decimal("0.125"),
    }
    match = re.fullmatch(r"(?:(\d+))?([½¼¾⅛])", coupon_text)
    if match:
        whole = Decimal(match.group(1)) if match.group(1) else Decimal("0")
        return whole + fraction_values[match.group(2)]
    ascii_fraction = re.fullmatch(r"(?:(\d+)\s+)?(\d+)/(\d+)", coupon_text)
    if ascii_fraction:
        whole = Decimal(ascii_fraction.group(1)) if ascii_fraction.group(1) else Decimal("0")
        return whole + Decimal(ascii_fraction.group(2)) / Decimal(ascii_fraction.group(3))
    try:
        return Decimal(coupon_text)
    except Exception as exc:
        raise ValueError(
            f"Cannot parse coupon rate from gilt name: {gilt_name!r}"
        ) from exc
