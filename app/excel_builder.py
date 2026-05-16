from __future__ import annotations

import datetime
from collections.abc import Iterable
from dataclasses import dataclass

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.worksheet import Worksheet

from app import formulas
from app.models import GiltMarketRow
from app.tax_scenarios import DEFAULT_TAX_SCENARIOS

DEFAULT_NOMINAL_AMOUNT = 10_000

SETTINGS_HEADERS = ("Setting", "Value", "Notes")

MARKET_DATA_HEADERS = (
    "ISIN",
    "Gilt Name",
    "Maturity",
    "Coupon %",
    "Imported Price",
    "Imported Yield %",
    "Valuation Date",
    "Source",
)

INPUT_HEADERS = ("ISIN", "Gilt Name", "Nominal Amount (£)", "Override Price", "Override Yield %")

ANALYSIS_HEADERS = (
    "ISIN",
    "Gilt Name",
    "Maturity",
    "Coupon %",
    "Effective Price",
    "Effective Yield %",
    "Approx Retail Ask Yield %",
    "Nominal Amount (£)",
    "Years to Maturity",
    "Annual Coupon Cash (£)",
    "Capital Uplift to Par (£)",
    "Approx Gross Cash Gain to Maturity (£)",
    "Coupon Tax @20% (£)",
    "Coupon Tax @40% (£)",
    "Coupon Tax @45% (£)",
    "CGT on Conventional Gilt Capital Gain (£)",
    "Approx Net Cash Gain @20% (£)",
    "Approx Net Cash Gain @40% (£)",
    "Approx Net Cash Gain @45% (£)",
)


def build_workbook(
    rows: Iterable[GiltMarketRow],
    *,
    default_nominal_amount: int = DEFAULT_NOMINAL_AMOUNT,
    retail_ask_yields: dict[str, float] | None = None,
) -> Workbook:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    settings = workbook.create_sheet("Settings")
    market_data = workbook.create_sheet("Market Data")
    inputs = workbook.create_sheet("Inputs")
    analysis = workbook.create_sheet("Analysis")
    summary_yield = workbook.create_sheet("Summary — Yield Ranking")
    summary_aftertax = workbook.create_sheet("Summary — After-tax Return")
    summary_bestval = workbook.create_sheet("Summary — Best Value")
    instructions = workbook.create_sheet("Instructions")

    _write_headers(settings, SETTINGS_HEADERS)
    _write_headers(market_data, MARKET_DATA_HEADERS)
    _write_headers(inputs, INPUT_HEADERS)
    _write_headers(analysis, ANALYSIS_HEADERS)

    materialized_rows = list(rows)
    resolved_yields = retail_ask_yields or {}
    _populate_settings(settings, default_nominal_amount)
    _populate_market_data(market_data, materialized_rows)
    _populate_inputs(inputs, materialized_rows)
    _populate_analysis(analysis, materialized_rows, resolved_yields)

    snapshots = _build_snapshots(materialized_rows, resolved_yields, default_nominal_amount)
    _populate_summary_yield(summary_yield, snapshots)
    _populate_summary_aftertax(summary_aftertax, snapshots)
    _populate_summary_bestval(summary_bestval, snapshots)

    _populate_instructions(instructions)

    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

    _format_analysis(analysis, len(materialized_rows))

    return workbook


# Columns (1-based) whose values are £ amounts — formatted to 2 decimal places.
_GBP_COLUMNS = (10, 11, 12, 13, 14, 15, 17, 18, 19)
_GBP_FORMAT = '£#,##0.00'
_ANALYSIS_FONT_SIZE = 9


def _format_analysis(sheet: Worksheet, data_row_count: int) -> None:
    # Word-wrap header row
    for cell in sheet[1]:
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Font size + £ format on data rows
    for r in range(2, data_row_count + 2):
        for cell in sheet[r]:
            cell.font = Font(size=_ANALYSIS_FONT_SIZE)
        for col in _GBP_COLUMNS:
            sheet.cell(r, col).number_format = _GBP_FORMAT

    # Print: fit all columns onto 1 page wide, portrait, landscape if needed
    sheet.page_setup.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0  # unlimited pages tall
    sheet.sheet_properties.pageSetUpPr.fitToPage = True


_RETAIL_ASK_YIELD_NOTE = (
    "Pre-computed import from DMO D10B sale dirty price.\n"
    "This value does NOT update when price or yield overrides change.\n"
    "Day count: days/365.25 (approximation — not exact Actual/Actual ICMA)."
)


def _write_headers(sheet: Worksheet, headers: tuple[str, ...]) -> None:
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    if "Approx Retail Ask Yield %" in headers:
        col_idx = headers.index("Approx Retail Ask Yield %") + 1
        col_letter = sheet.cell(1, col_idx).column_letter
        cell = sheet[f"{col_letter}1"]
        cell.comment = Comment(_RETAIL_ASK_YIELD_NOTE, "GiltAnalyzer")


def _populate_settings(sheet: Worksheet, default_nominal_amount: int) -> None:
    sheet.append(("DefaultNominalAmount", default_nominal_amount, "Default £ face value applied to all gilts unless overridden in the Inputs sheet (column B)."))
    sheet.append((None, None, None))
    sheet.append(("--- Tax Scenarios ---", None, None))
    for scenario in DEFAULT_TAX_SCENARIOS:
        sheet.append((scenario.name, float(scenario.coupon_tax_rate), scenario.notes))


def _populate_market_data(sheet: Worksheet, rows: list[GiltMarketRow]) -> None:
    for row in rows:
        sheet.append(
            (
                row.isin,
                row.gilt_name,
                row.maturity_date,
                float(row.coupon_rate),
                float(row.imported_price) if row.imported_price is not None else None,
                float(row.imported_yield) if row.imported_yield is not None else None,
                row.valuation_date,
                row.source_name,
            )
        )


def _populate_inputs(sheet: Worksheet, rows: list[GiltMarketRow]) -> None:
    for excel_row, row in enumerate(rows, start=2):
        sheet.cell(excel_row, 1, row.isin)
        # Gilt Name: read-only lookup from Market Data — do not edit
        sheet.cell(excel_row, 2, f"=IFERROR(INDEX('Market Data'!B:B,MATCH(A{excel_row},'Market Data'!A:A,0)),\"\")")
        # Cols C, D, E: user-editable (Nominal Amount, Override Price, Override Yield)


def _populate_analysis(
    sheet: Worksheet,
    rows: list[GiltMarketRow],
    retail_ask_yields: dict[str, float],
) -> None:
    for excel_row, row in enumerate(rows, start=2):
        sheet.cell(excel_row, 1, row.isin)
        sheet.cell(excel_row, 2, row.gilt_name)
        sheet.cell(excel_row, 3, row.maturity_date)
        sheet.cell(excel_row, 4, float(row.coupon_rate))
        sheet.cell(excel_row, 5, formulas.effective_price(excel_row))
        sheet.cell(excel_row, 6, formulas.effective_yield(excel_row))
        sheet.cell(excel_row, 7, retail_ask_yields.get(row.isin))
        sheet.cell(excel_row, 8, formulas.nominal_amount(excel_row))
        sheet.cell(excel_row, 9, formulas.years_to_maturity(excel_row))
        sheet.cell(excel_row, 10, formulas.annual_coupon_cash(excel_row))
        sheet.cell(excel_row, 11, formulas.capital_uplift_to_par(excel_row))
        sheet.cell(excel_row, 12, formulas.approx_gross_cash_gain_to_maturity(excel_row))
        sheet.cell(excel_row, 13, formulas.coupon_tax(excel_row, "20%"))
        sheet.cell(excel_row, 14, formulas.coupon_tax(excel_row, "40%"))
        sheet.cell(excel_row, 15, formulas.coupon_tax(excel_row, "45%"))
        sheet.cell(excel_row, 16, formulas.cgt_on_conventional_gilt_capital_gain(excel_row))
        sheet.cell(excel_row, 17, formulas.approx_net_cash_gain(excel_row, "M"))
        sheet.cell(excel_row, 18, formulas.approx_net_cash_gain(excel_row, "N"))
        sheet.cell(excel_row, 19, formulas.approx_net_cash_gain(excel_row, "O"))


@dataclass(frozen=True)
class _GiltSnapshot:
    isin: str
    gilt_name: str
    maturity: datetime.date
    coupon_pct: float
    effective_yield: float | None
    retail_ask_yield: float | None
    nominal: float
    years: float
    annual_coupon: float
    capital_uplift: float
    gross_gain: float
    net_20: float
    net_40: float
    net_45: float


def _build_snapshots(
    rows: list[GiltMarketRow],
    retail_ask_yields: dict[str, float],
    default_nominal_amount: int,
) -> list[_GiltSnapshot]:
    today = datetime.date.today()
    snapshots = []
    for row in rows:
        nominal = float(default_nominal_amount)
        price = float(row.imported_price) if row.imported_price is not None else None
        eff_yield = float(row.imported_yield) if row.imported_yield is not None else None
        coupon_pct = float(row.coupon_rate)

        days = (row.maturity_date - today).days
        years = days / 365.25 if days > 0 else 0.0

        annual_coupon = nominal * coupon_pct / 100
        capital_uplift = nominal * (100 - price) / 100 if price is not None else 0.0
        gross_gain = annual_coupon * years + capital_uplift

        net_20 = gross_gain - annual_coupon * years * 0.20
        net_40 = gross_gain - annual_coupon * years * 0.40
        net_45 = gross_gain - annual_coupon * years * 0.45

        snapshots.append(_GiltSnapshot(
            isin=row.isin,
            gilt_name=row.gilt_name,
            maturity=row.maturity_date,
            coupon_pct=coupon_pct,
            effective_yield=eff_yield,
            retail_ask_yield=retail_ask_yields.get(row.isin),
            nominal=nominal,
            years=years,
            annual_coupon=annual_coupon,
            capital_uplift=capital_uplift,
            gross_gain=gross_gain,
            net_20=net_20,
            net_40=net_40,
            net_45=net_45,
        ))
    return snapshots


def _write_summary_headers(sheet: Worksheet, headers: tuple[str, ...]) -> None:
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def _format_summary_sheet(sheet: Worksheet, data_row_count: int, gbp_cols: tuple[int, ...]) -> None:
    for r in range(2, data_row_count + 2):
        for cell in sheet[r]:
            cell.font = Font(size=_ANALYSIS_FONT_SIZE)
        for col in gbp_cols:
            sheet.cell(r, col).number_format = _GBP_FORMAT
    sheet.page_setup.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True


def _populate_summary_yield(sheet: Worksheet, snapshots: list[_GiltSnapshot]) -> None:
    headers = (
        "Gilt Name", "Maturity", "Coupon %",
        "Effective Yield %", "Retail Ask Yield %", "Nominal Amount (£)",
    )
    _write_summary_headers(sheet, headers)
    sorted_rows = sorted(
        (s for s in snapshots if s.effective_yield is not None),
        key=lambda s: s.effective_yield,  # type: ignore[arg-type]
        reverse=True,
    )
    for s in sorted_rows:
        sheet.append((
            s.gilt_name, s.maturity, s.coupon_pct,
            s.effective_yield, s.retail_ask_yield, s.nominal,
        ))
    _format_summary_sheet(sheet, len(sorted_rows), gbp_cols=(6,))


def _populate_summary_aftertax(sheet: Worksheet, snapshots: list[_GiltSnapshot]) -> None:
    headers = (
        "Gilt Name", "Maturity", "Years to Maturity", "Nominal Amount (£)",
        "Gross Cash Gain (£)",
        "Net Gain @20% (£)", "Net Gain @40% (£)", "Net Gain @45% (£)",
        "Annual Net @20% (£)", "Annual Net @40% (£)", "Annual Net @45% (£)",
    )
    _write_summary_headers(sheet, headers)
    sorted_rows = sorted(
        snapshots,
        key=lambda s: (s.net_40 / s.years) if s.years > 0 else 0.0,
        reverse=True,
    )
    for s in sorted_rows:
        ann_net_20 = s.net_20 / s.years if s.years > 0 else 0.0
        ann_net_40 = s.net_40 / s.years if s.years > 0 else 0.0
        ann_net_45 = s.net_45 / s.years if s.years > 0 else 0.0
        sheet.append((
            s.gilt_name, s.maturity, round(s.years, 2), s.nominal,
            s.gross_gain, s.net_20, s.net_40, s.net_45,
            ann_net_20, ann_net_40, ann_net_45,
        ))
    _format_summary_sheet(sheet, len(sorted_rows), gbp_cols=(4, 5, 6, 7, 8, 9, 10, 11))


def _populate_summary_bestval(sheet: Worksheet, snapshots: list[_GiltSnapshot]) -> None:
    headers = (
        "Gilt Name", "Maturity", "Years to Maturity",
        "Effective Yield %", "Retail Ask Yield %", "Coupon %",
        "Capital Uplift to Par (£)",
        "Annual Net Return @20% per £10k",
        "Annual Net Return @40% per £10k",
        "Annual Net Return @45% per £10k",
        "Total Net Return @20% per £10k",
        "Total Net Return @40% per £10k",
        "Total Net Return @45% per £10k",
    )
    _write_summary_headers(sheet, headers)
    # Unsorted — user can sort by any column in Excel
    for s in snapshots:
        scale = 10_000 / s.nominal if s.nominal else 1.0
        ann_20 = (s.net_20 / s.years * scale) if s.years > 0 else 0.0
        ann_40 = (s.net_40 / s.years * scale) if s.years > 0 else 0.0
        ann_45 = (s.net_45 / s.years * scale) if s.years > 0 else 0.0
        tot_20 = s.net_20 * scale
        tot_40 = s.net_40 * scale
        tot_45 = s.net_45 * scale
        sheet.append((
            s.gilt_name, s.maturity, round(s.years, 2),
            s.effective_yield, s.retail_ask_yield, s.coupon_pct,
            s.capital_uplift * scale,
            ann_20, ann_40, ann_45,
            tot_20, tot_40, tot_45,
        ))
    _format_summary_sheet(sheet, len(snapshots), gbp_cols=(7, 8, 9, 10, 11, 12, 13))


_INSTRUCTIONS = (
    ("GiltAnalyzer — How to use the Inputs sheet", True),
    ("", False),
    ("ISIN (column A)", True),
    ("Read-only. Identifies each gilt. Do not edit.", False),
    ("", False),
    ("Gilt Name (column B)", True),
    ("Read-only formula — populated automatically from Market Data. Do not edit.", False),
    ("", False),
    ("Nominal Amount — column C", True),
    ("The £ face value of your holding in this gilt.", False),
    ("Leave blank to use the default from Settings!B2 (DefaultNominalAmount).", False),
    ("Enter a number (e.g. 50000) to override for this gilt only.", False),
    ("All cash flow columns in Analysis (J, K, L, M, N, O, Q, R, S) scale with this value.", False),
    ("", False),
    ("Override Price — column D", True),
    ("Optional. Enter a clean price to replace the DividendData market price.", False),
    ("When set, Effective Price (Analysis col E) uses this value instead of the live price.", False),
    ("Effective Yield (Analysis col F) also updates to reflect the overridden price.", False),
    ("Leave blank to use the live market price.", False),
    ("", False),
    ("Override Yield % — column E", True),
    ("Optional. Enter a yield (in %) to replace the DividendData market yield.", False),
    ("When set, Effective Yield (Analysis col F) uses this value directly.", False),
    ("Note: Override Price takes precedence — if both are set, price wins for col E,", False),
    ("but yield override still applies to col F.", False),
    ("Leave blank to use the live market yield.", False),
    ("", False),
    ("Yield columns explained", True),
    ("Effective Yield % (Analysis col F): yield derived from the DividendData market price", False),
    ("  (or your Override Price / Override Yield if set).", False),
    ("Approx Retail Ask Yield % (Analysis col G): yield at the DMO retail ask (dirty) price,", False),
    ("  sourced from the D10B XLS at export time. Static — does not update with overrides.", False),
    ("  The DMO retail price is typically slightly higher than the market mid, so this yield", False),
    ("  will generally be slightly lower than Effective Yield. The gap reflects the DMO spread.", False),
    ("", False),
    ("Settings sheet", True),
    ("DefaultNominalAmount (row 2): change this single cell to set the default holding size", False),
    ("  for all gilts that have no override in the Inputs sheet.", False),
    ("Tax scenario rates (rows 5–7) are for reference only — the Analysis sheet uses", False),
    ("  hardcoded 20%, 40%, 45% rates in columns M, N, O.", False),
    ("", False),
    ("Summary — Yield Ranking", True),
    ("Shows all gilts sorted by Effective Yield % (highest first).", False),
    ("What yield means: if you buy today and hold to maturity, this is your annualised total", False),
    ("  return — combining coupon income AND the capital gain/loss as the price returns to £100.", False),
    ("This is NOT the cash you receive in year 1. Year-1 income is the coupon % applied to", False),
    ("  your nominal amount (e.g. 1.5% coupon on £10,000 = £150 cash in year 1).", False),
    ("Coupons do not compound — each payment is the same fixed cash amount.", False),
    ("Higher yield often means longer maturity or lower price — use with the other summary", False),
    ("  sheets to understand the full picture.", False),
    ("", False),
    ("Summary — After-tax Return", True),
    ("Shows the actual cash you keep after income tax on coupons, sorted by Annual Net @40%.", False),
    ("Gross Cash Gain = all coupon payments over the holding period + capital uplift to £100 par.", False),
    ("Net Gain @20/40/45% = Gross Gain minus income tax on coupons at that rate.", False),
    ("  Capital gains on conventional gilts are always CGT-exempt, so only coupons are taxed.", False),
    ("Annual Net = Net Gain divided by Years to Maturity — this makes short and long-dated", False),
    ("  gilts directly comparable on a per-year basis.", False),
    ("Sort by whichever Annual Net column matches your tax rate to find the best after-tax", False),
    ("  annual return for your situation.", False),
    ("", False),
    ("Summary — Best Value", True),
    ("An unsorted multi-dimension view — sort any column to answer a specific question.", False),
    ("All cash figures are normalised to per £10,000 nominal so gilts are directly comparable", False),
    ("  regardless of how much you invest.", False),
    ("Suggested sorts:", False),
    ("  Annual Net @40% per £10k  — best after-tax income per year (higher rate taxpayer)", False),
    ("  Annual Net @20% per £10k  — best after-tax income per year (basic rate taxpayer)", False),
    ("  Total Net @40% per £10k   — best total cash over the full holding period", False),
    ("  Effective Yield %          — best annualised return pre-tax", False),
    ("  Capital Uplift to Par      — most capital gain locked in (all tax-free)", False),
    ("  Years to Maturity          — filter by how long you are willing to hold", False),
    ("", False),
    ("Selling before maturity", True),
    ("You do not have to hold a gilt to maturity. If you sell early:", False),
    ("  - All coupons already paid are yours to keep.", False),
    ("  - Accrued interest since the last coupon is paid by the buyer automatically.", False),
    ("  - Your capital outcome depends on the price at the date you sell.", False),
    ("    If rates have risen since you bought, the price will be lower — potential capital loss.", False),
    ("    If rates have fallen, the price will be higher — capital gain (still CGT-exempt).", False),
    ("To model an early exit: enter your expected sale price in Override Price (Inputs col D).", False),
    ("  The Analysis and Summary sheets will recalculate using that price instead of £100 par.", False),
)


def _populate_instructions(sheet: Worksheet) -> None:
    sheet.column_dimensions["A"].width = 90
    for text, is_heading in _INSTRUCTIONS:
        sheet.append((text,))
        if text and is_heading:
            sheet.cell(sheet.max_row, 1).font = Font(bold=True, size=11)
        elif text:
            sheet.cell(sheet.max_row, 1).alignment = Alignment(wrap_text=True)
    sheet.freeze_panes = None  # no freeze on instructions
    sheet.auto_filter.ref = None
