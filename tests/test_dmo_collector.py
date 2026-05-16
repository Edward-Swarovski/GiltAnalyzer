from decimal import Decimal

from app.collectors.dmo import parse_d1a_xml


def test_parse_d1a_xml_keeps_only_conventional_gilts() -> None:
    rows = parse_d1a_xml(
        """
        <Data>
          <View_GILTS_IN_ISSUE
            CLOSE_OF_BUSINESS_DATE="2026-05-14T00:00:00"
            INSTRUMENT_TYPE="Conventional "
            INSTRUMENT_NAME="1½% Treasury Gilt 2029"
            REDEMPTION_DATE="2029-07-22T00:00:00"
            ISIN_CODE="GB00TEST0001" />
          <View_GILTS_IN_ISSUE
            CLOSE_OF_BUSINESS_DATE="2026-05-14T00:00:00"
            INSTRUMENT_TYPE="Index-linked "
            INSTRUMENT_NAME="0⅛% Index-linked Treasury Gilt 2039"
            REDEMPTION_DATE="2039-03-22T00:00:00"
            ISIN_CODE="GB00TEST0002" />
        </Data>
        """
    )

    assert len(rows) == 1
    assert rows[0].isin == "GB00TEST0001"
    assert rows[0].coupon_rate == Decimal("1.5")
    assert rows[0].imported_price is None


def test_parse_d1a_xml_parses_fractional_coupon() -> None:
    rows = parse_d1a_xml(
        """
        <Data>
          <View_GILTS_IN_ISSUE
            CLOSE_OF_BUSINESS_DATE="2026-05-14T00:00:00"
            INSTRUMENT_TYPE="Conventional "
            INSTRUMENT_NAME="0⅛% Treasury Gilt 2028"
            REDEMPTION_DATE="2028-01-31T00:00:00"
            ISIN_CODE="GB00TEST0003" />
        </Data>
        """
    )

    assert rows[0].coupon_rate == Decimal("0.125")


def test_parse_d1a_xml_parses_ascii_fractional_coupon() -> None:
    rows = parse_d1a_xml(
        """
        <Data>
          <View_GILTS_IN_ISSUE
            CLOSE_OF_BUSINESS_DATE="2026-05-14T00:00:00"
            INSTRUMENT_TYPE="Conventional "
            INSTRUMENT_NAME="0 3/8% Treasury Gilt 2026"
            REDEMPTION_DATE="2026-10-22T00:00:00"
            ISIN_CODE="GB00TEST0004" />
        </Data>
        """
    )

    assert rows[0].coupon_rate == Decimal("0.375")
