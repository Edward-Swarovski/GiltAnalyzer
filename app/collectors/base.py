from __future__ import annotations

from abc import ABC, abstractmethod

from app.models import GiltMarketRow, GiltPriceQuote


class BaseMarketRowCollector(ABC):
    @abstractmethod
    def fetch(self) -> list[GiltMarketRow]:
        """Fetch and normalize market rows from a data source."""


class BasePriceQuoteCollector(ABC):
    @abstractmethod
    def fetch(self) -> list[GiltPriceQuote]:
        """Fetch normalized quote rows used to enrich market rows."""
