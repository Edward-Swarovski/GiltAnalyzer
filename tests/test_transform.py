from decimal import Decimal

from app.transform import normalize_market_row


def test_normalize_market_row() -> None:
    row = normalize_market_row(
        {
            "isin": " GB00TEST1234 ",
            "gilt_name": " 0.125% Treasury Gilt 2028 ",
            "coupon_rate": "0.125",
            "maturity_date": "2028-01-31",
            "price": "91.42",
            "yield": "4.10",
            "valuation_date": "2026-05-15",
        },
        source_name="dmo",
    )

    assert row.isin == "GB00TEST1234"
    assert row.gilt_name == "0.125% Treasury Gilt 2028"
    assert row.coupon_rate == Decimal("0.125")
    assert row.imported_price == Decimal("91.42")
    assert row.imported_yield == Decimal("4.10")
    assert row.source_name == "dmo"


def test_normalize_market_row_without_price() -> None:
    row = normalize_market_row(
        {
            "isin": "GB00TEST5678",
            "gilt_name": "1.5% Treasury Gilt 2030",
            "coupon_rate": "1.5",
            "maturity_date": "2030-01-31",
            "price": "",
            "yield": "",
            "valuation_date": "2026-05-15",
        },
        source_name="manual",
    )

    assert row.imported_price is None
    assert row.imported_yield is None
