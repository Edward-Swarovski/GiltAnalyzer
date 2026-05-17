from datetime import date
from decimal import Decimal

from app.excel_builder import build_workbook
from app.models import GiltMarketRow


def sample_row() -> GiltMarketRow:
    return GiltMarketRow(
        isin="GB00TEST1234",
        gilt_name="0.125% Treasury Gilt 2028",
        coupon_rate=Decimal("0.125"),
        maturity_date=date(2028, 1, 31),
        imported_price=Decimal("91.42"),
        imported_yield=Decimal("4.10"),
        valuation_date=date(2026, 5, 15),
        source_name="dmo",
    )


def test_workbook_has_expected_sheets() -> None:
    workbook = build_workbook([sample_row()])
    assert workbook.sheetnames == [
        "Settings",
        "Market Data",
        "Inputs",
        "Analysis",
        "Instructions",
    ]


def test_analysis_sheet_contains_formulas() -> None:
    workbook = build_workbook([sample_row()])
    analysis = workbook["Analysis"]

    assert analysis["E2"].value.startswith("=IFERROR(")
    assert analysis["F2"].value.startswith("=IFERROR(")
    assert analysis["H2"].value.startswith("=IFERROR(")
    assert "Settings!B:B" in analysis["H2"].value
    # Net gain columns (P/Q/R) no longer subtract a CGT column
    assert analysis["P2"].value == "=L2-M2"
    assert analysis["Q2"].value == "=L2-N2"
    assert analysis["R2"].value == "=L2-O2"
    # Annual Net columns (S/T/U) — blue
    assert analysis["S2"].value == '=IFERROR(P2/I2,"")'
    assert analysis["T2"].value == '=IFERROR(Q2/I2,"")'
    assert analysis["U2"].value == '=IFERROR(R2/I2,"")'


def test_settings_sheet_has_default_nominal_amount() -> None:
    workbook = build_workbook([sample_row()])
    settings = workbook["Settings"]
    assert settings["A2"].value == "DefaultNominalAmount"
    assert settings["B2"].value == 10_000


def test_settings_sheet_contains_tax_scenarios() -> None:
    workbook = build_workbook([sample_row()])
    settings = workbook["Settings"]
    # Row 3 is blank separator, row 4 is the section label, rows 5-7 are tax scenarios
    assert settings["A5"].value == "Basic rate"
    assert settings["B5"].value == 0.20
    assert settings["A6"].value == "Higher rate"
    assert settings["B6"].value == 0.40
    assert settings["A7"].value == "Additional rate"
    assert settings["B7"].value == 0.45


def test_settings_default_nominal_amount_is_configurable() -> None:
    workbook = build_workbook([sample_row()], default_nominal_amount=25000)
    assert workbook["Settings"]["B2"].value == 25000


def test_inputs_has_gilt_name_formula_in_column_b() -> None:
    workbook = build_workbook([sample_row()])
    inputs = workbook["Inputs"]
    assert inputs["B1"].value == "Gilt Name"
    assert "Market Data" in inputs["B2"].value


def test_inputs_nominal_amount_column_is_blank_by_default() -> None:
    workbook = build_workbook([sample_row()])
    assert workbook["Inputs"]["C2"].value is None


def test_instructions_sheet_exists() -> None:
    workbook = build_workbook([sample_row()])
    instructions = workbook["Instructions"]
    assert instructions["A1"].value is not None
