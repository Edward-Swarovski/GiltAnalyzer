from dataclasses import dataclass
from decimal import Decimal


@dataclass(frozen=True)
class TaxScenario:
    name: str
    coupon_tax_rate: Decimal
    notes: str


DEFAULT_TAX_SCENARIOS = (
    TaxScenario(
        name="Basic rate",
        coupon_tax_rate=Decimal("0.20"),
        notes="Scenario rate only; does not model personal allowances.",
    ),
    TaxScenario(
        name="Higher rate",
        coupon_tax_rate=Decimal("0.40"),
        notes="Scenario rate only; does not model personal allowances.",
    ),
    TaxScenario(
        name="Additional rate",
        coupon_tax_rate=Decimal("0.45"),
        notes="Scenario rate only; does not model personal allowances.",
    ),
)
