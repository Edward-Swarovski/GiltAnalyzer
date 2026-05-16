from datetime import date
from decimal import Decimal

import pytest

from app.collectors.dividenddata import (
    DividendDataCollector,
    _parse_currency,
    enrich_with_quotes,
    parse_quotes_html,
)
from app.models import GiltMarketRow


def test_parse_quotes_html() -> None:
    quotes = parse_quotes_html(
        """
        <table id="main-table">
          <tbody>
            <tr>
              <td>TG28</td>
              <td>1 5/8% Treasury Gilt 2028</td>
              <td>1.625%</td>
              <td>22-Oct-2028</td>
              <td>2 years 160 days</td>
              <td>£93.61</td>
              <td>1.736%</td>
              <td>4.419%</td>
              <td></td>
            </tr>
          </tbody>
        </table>
        """
    )

    assert len(quotes) == 1
    assert quotes[0].coupon_rate == Decimal("1.625")
    assert quotes[0].clean_price == Decimal("93.61")
    assert quotes[0].yield_to_maturity == Decimal("4.419")


def test_enrich_with_quotes_matches_by_name_maturity_and_coupon() -> None:
    row = GiltMarketRow(
        isin="GB00TEST1234",
        gilt_name="0⅛% Treasury Gilt 2028",
        coupon_rate=Decimal("0.125"),
        maturity_date=date(2028, 1, 31),
        imported_price=None,
        imported_yield=None,
        valuation_date=date(2026, 5, 15),
        source_name="dmo_d1a",
    )
    quotes = parse_quotes_html(
        """
        <table id="main-table">
          <tbody>
            <tr>
              <td>TN28</td>
              <td>0.125% Treasury Gilt 2028</td>
              <td>0.125%</td>
              <td>31-Jan-2028</td>
              <td>1 year 260 days</td>
              <td>£93.20</td>
              <td>0.134%</td>
              <td>4.285%</td>
              <td></td>
            </tr>
          </tbody>
        </table>
        """
    )

    enriched = enrich_with_quotes([row], quotes)

    assert enriched[0].imported_price == Decimal("93.20")
    assert enriched[0].imported_yield == Decimal("4.285")
    assert enriched[0].source_name.endswith("+dividenddata")


def test_fetch_rejects_response_without_quote_table(monkeypatch) -> None:
    class FakeResponse:
        text = "<html><body>blocked</body></html>"

        @staticmethod
        def raise_for_status() -> None:
            return None

    monkeypatch.setattr(
        "app.collectors.dividenddata.requests.get",
        lambda *args, **kwargs: FakeResponse(),
    )

    with pytest.raises(RuntimeError, match="recognizable gilt quote table"):
        DividendDataCollector().fetch()


def test_enrich_with_quotes_logs_unmatched_rows(caplog) -> None:
    row = GiltMarketRow(
        isin="GB00UNMATCHED",
        gilt_name="9% Treasury Gilt 2099",
        coupon_rate=Decimal("9"),
        maturity_date=date(2099, 1, 1),
        imported_price=None,
        imported_yield=None,
        valuation_date=date(2026, 5, 15),
        source_name="dmo_d1a",
    )

    with caplog.at_level("WARNING"):
        enriched = enrich_with_quotes([row], [])

    assert enriched[0].imported_price is None
    assert "No DividendData quote matched DMO row" in caplog.text


def test_parse_currency_preserves_negative_sign() -> None:
    assert _parse_currency("-£93.61") == Decimal("-93.61")


def test_parse_currency_raises_on_empty_or_dash() -> None:
    with pytest.raises(ValueError):
        _parse_currency("£-")
    with pytest.raises(ValueError):
        _parse_currency("-")


def test_enrich_with_quotes_treats_coupon_mismatch_as_unmatched(caplog) -> None:
    row = GiltMarketRow(
        isin="GB00TEST1234",
        gilt_name="0⅛% Treasury Gilt 2028",
        coupon_rate=Decimal("0.125"),
        maturity_date=date(2028, 1, 31),
        imported_price=None,
        imported_yield=None,
        valuation_date=date(2026, 5, 15),
        source_name="dmo_d1a",
    )
    # Quote has same name and maturity but wrong coupon — should not match
    quotes = parse_quotes_html(
        """
        <table id="main-table">
          <tbody>
            <tr>
              <td>TN28</td>
              <td>0.125% Treasury Gilt 2028</td>
              <td>1.500%</td>
              <td>31-Jan-2028</td>
              <td>1 year 260 days</td>
              <td>£93.20</td>
              <td>0.134%</td>
              <td>4.285%</td>
              <td></td>
            </tr>
          </tbody>
        </table>
        """
    )

    with caplog.at_level("WARNING"):
        enriched = enrich_with_quotes([row], quotes)

    assert enriched[0].imported_price is None
    assert "No DividendData quote matched DMO row" in caplog.text
