from __future__ import annotations

from datetime import datetime
from decimal import Decimal
import logging
import re

import requests
from bs4 import BeautifulSoup

from app.collectors.base import BasePriceQuoteCollector
from app.models import GiltMarketRow, GiltPriceQuote

logger = logging.getLogger(__name__)


class DividendDataCollector(BasePriceQuoteCollector):
    """Collect current conventional-gilt quotes from DividendData."""

    QUOTES_URL = (
        "https://www.dividenddata.co.uk/"
        "uk-gilts-prices-yields.py?showCompDetails=999"
    )

    def fetch(self) -> list[GiltPriceQuote]:
        response = requests.get(self.QUOTES_URL, timeout=30)
        response.raise_for_status()
        quotes = parse_quotes_html(response.text)
        if not quotes:
            raise RuntimeError(
                "DividendData did not return a recognizable gilt quote table."
            )
        return quotes


def parse_quotes_html(html_text: str) -> list[GiltPriceQuote]:
    soup = BeautifulSoup(html_text, "html.parser")
    quotes: list[GiltPriceQuote] = []

    for row in soup.select("table#main-table tbody tr"):
        cells = [cell.get_text(" ", strip=True) for cell in row.find_all("td")]
        if len(cells) < 8:
            continue

        try:
            quotes.append(
                GiltPriceQuote(
                    epic=cells[0],
                    gilt_name=cells[1],
                    coupon_rate=_parse_percentage(cells[2]),
                    maturity_date=datetime.strptime(cells[3], "%d-%b-%Y").date(),
                    clean_price=_parse_currency(cells[5]),
                    yield_to_maturity=_parse_percentage(cells[7]),
                )
            )
        except Exception:
            logger.warning(
                "Skipping malformed DividendData row: %s",
                cells,
            )

    return quotes


def enrich_with_quotes(
    rows: list[GiltMarketRow],
    quotes: list[GiltPriceQuote],
) -> list[GiltMarketRow]:
    quotes_by_key = {_quote_key(quote): quote for quote in quotes}
    enriched_rows: list[GiltMarketRow] = []

    for row in rows:
        quote = quotes_by_key.get(_market_row_key(row))
        if quote is None:
            logger.warning(
                "No DividendData quote matched DMO row: isin=%s name=%s maturity=%s coupon=%s",
                row.isin,
                row.gilt_name,
                row.maturity_date.isoformat(),
                row.coupon_rate,
            )
        enriched_rows.append(
            GiltMarketRow(
                isin=row.isin,
                gilt_name=row.gilt_name,
                coupon_rate=row.coupon_rate,
                maturity_date=row.maturity_date,
                imported_price=quote.clean_price if quote else row.imported_price,
                imported_yield=quote.yield_to_maturity if quote else row.imported_yield,
                valuation_date=row.valuation_date,
                source_name=(
                    f"{row.source_name}+dividenddata" if quote else row.source_name
                ),
            )
        )

    return enriched_rows


def _market_row_key(row: GiltMarketRow) -> tuple[str, str, Decimal]:
    return (_normalize_name(row.gilt_name), row.maturity_date.isoformat(), row.coupon_rate)


def _quote_key(quote: GiltPriceQuote) -> tuple[str, str, Decimal]:
    return (
        _normalize_name(quote.gilt_name),
        quote.maturity_date.isoformat(),
        quote.coupon_rate,
    )


def _normalize_name(value: str) -> str:
    normalized = value.replace("Treasury Stock", "Treasury Gilt")
    if "%" in normalized:
        normalized = normalized.split("%", maxsplit=1)[1]
    return re.sub(r"\s+", " ", normalized).strip().lower()


def _parse_currency(value: str) -> Decimal:
    cleaned = re.sub(r"[^\d.\-]", "", value)
    if not cleaned or cleaned == "-":
        raise ValueError(f"Cannot parse currency value: {value!r}")
    return Decimal(cleaned)


def _parse_percentage(value: str) -> Decimal:
    return Decimal(value.replace("%", "").strip())
