from __future__ import annotations

from datetime import date
from decimal import Decimal
import re
from xml.etree import ElementTree as ET

import requests

from app.collectors.base import BaseMarketRowCollector
from app.models import GiltMarketRow


class DmoCollector(BaseMarketRowCollector):
    """Collect conventional gilt reference data from the DMO D1A XML report."""

    D1A_URL = "https://www.dmo.gov.uk/data/XmlDataReport?reportCode=D1A"

    def fetch(self) -> list[GiltMarketRow]:
        response = requests.get(self.D1A_URL, timeout=30)
        response.raise_for_status()

        if "xml" not in response.headers.get("content-type", "").lower():
            raise RuntimeError(
                "DMO did not return XML. The site may be presenting an anti-bot page."
            )

        return parse_d1a_xml(response.text)


def parse_d1a_xml(xml_text: str) -> list[GiltMarketRow]:
    root = ET.fromstring(xml_text)
    rows: list[GiltMarketRow] = []

    for node in root.iter("View_GILTS_IN_ISSUE"):
        instrument_type = node.attrib.get("INSTRUMENT_TYPE", "").strip()
        if instrument_type.lower() != "conventional":
            continue

        instrument_name = node.attrib["INSTRUMENT_NAME"].strip()
        rows.append(
            GiltMarketRow(
                isin=node.attrib["ISIN_CODE"].strip(),
                gilt_name=instrument_name,
                coupon_rate=_parse_coupon_rate(instrument_name),
                maturity_date=_parse_dmo_date(node.attrib["REDEMPTION_DATE"]),
                imported_price=None,
                imported_yield=None,
                valuation_date=_parse_dmo_date(node.attrib["CLOSE_OF_BUSINESS_DATE"]),
                source_name="dmo_d1a",
            )
        )

    return rows


def _parse_dmo_date(raw: str) -> date:
    return date.fromisoformat(raw[:10])


def _parse_coupon_rate(instrument_name: str) -> Decimal:
    coupon_text = instrument_name.split("%", maxsplit=1)[0].strip()
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
        numerator = Decimal(ascii_fraction.group(2))
        denominator = Decimal(ascii_fraction.group(3))
        return whole + (numerator / denominator)

    try:
        return Decimal(coupon_text)
    except Exception as exc:
        raise ValueError(
            f"Cannot parse coupon rate from instrument name: {instrument_name!r}"
        ) from exc
