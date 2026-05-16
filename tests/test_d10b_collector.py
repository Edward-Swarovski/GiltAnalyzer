from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from app.collectors.d10b import approximate_retail_ask_yield, parse_d10b_xls


_REAL_FILE = Path("data/20260515 - DMO Gilt Purchase and Sale Service Prices.xls")


@pytest.mark.skipif(not _REAL_FILE.exists(), reason="real D10B file not present")
def test_parse_d10b_xls_real_file() -> None:
    quotes = parse_d10b_xls(_REAL_FILE)
    assert len(quotes) >= 70
    assert quotes[0].isin == "GB00BYZW3G56"
    assert quotes[0].coupon_rate == Decimal("1.5")
    assert quotes[0].sale_clean_price == Decimal("99.61")
    assert quotes[0].sale_dirty_price == Decimal("100.09895")
    assert quotes[0].data_date == date(2026, 5, 15)


def test_approximate_retail_ask_yield_is_plausible() -> None:
    result = approximate_retail_ask_yield(
        coupon_rate=Decimal("0.125"),
        sale_dirty_price=Decimal("93.267638"),
        redemption_date=date(2028, 1, 31),
        valuation_date=date(2026, 5, 15),
    )
    assert Decimal("4.0") < result < Decimal("4.5")


def test_approximate_retail_ask_yield_raises_on_implausible_result() -> None:
    with pytest.raises(ValueError, match="outside the plausible range"):
        approximate_retail_ask_yield(
            coupon_rate=Decimal("5.0"),
            sale_dirty_price=Decimal("0.001"),
            redemption_date=date(2028, 1, 31),
            valuation_date=date(2026, 5, 15),
        )
