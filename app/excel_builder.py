from __future__ import annotations

from collections.abc import Iterable

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
    "Approx Net Cash Gain @20% (£)",
    "Approx Net Cash Gain @40% (£)",
    "Approx Net Cash Gain @45% (£)",
    "Annual Net @20% (£)",
    "Annual Net @40% (£)",
    "Annual Net @45% (£)",
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

    _populate_instructions(instructions)

    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

    _format_analysis(analysis, len(materialized_rows))

    return workbook


# Columns (1-based) whose values are £ amounts — formatted to 2 decimal places.
# A=1..O=15 unchanged; P=Net@20(16), Q=Net@40(17), R=Net@45(18), S=AnnNet@20(19), T=AnnNet@40(20), U=AnnNet@45(21)
_GBP_COLUMNS = (10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21)
_GBP_FORMAT = '£#,##0.00'
_ANALYSIS_FONT_SIZE = 9


def _format_analysis(sheet: Worksheet, data_row_count: int) -> None:
    # Word-wrap header row
    for cell in sheet[1]:
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    _BLUE = "0070C0"
    _ANNUAL_NET_COLS = (19, 20, 21)  # S, T, U

    # Word-wrap + blue header for Annual Net columns
    for col in _ANNUAL_NET_COLS:
        sheet.cell(1, col).font = Font(bold=True, color=_BLUE)

    # Font size + £ format on data rows; blue for Annual Net columns
    for r in range(2, data_row_count + 2):
        for cell in sheet[r]:
            cell.font = Font(size=_ANALYSIS_FONT_SIZE)
        for col in _GBP_COLUMNS:
            sheet.cell(r, col).number_format = _GBP_FORMAT
        for col in _ANNUAL_NET_COLS:
            sheet.cell(r, col).font = Font(size=_ANALYSIS_FONT_SIZE, color=_BLUE)

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
        sheet.cell(excel_row, 16, formulas.approx_net_cash_gain(excel_row, "M"))
        sheet.cell(excel_row, 17, formulas.approx_net_cash_gain(excel_row, "N"))
        sheet.cell(excel_row, 18, formulas.approx_net_cash_gain(excel_row, "O"))
        sheet.cell(excel_row, 19, formulas.annual_net_gain(excel_row, "P"))
        sheet.cell(excel_row, 20, formulas.annual_net_gain(excel_row, "Q"))
        sheet.cell(excel_row, 21, formulas.annual_net_gain(excel_row, "R"))



_INSTRUCTIONS = (
    ("GiltAnalyzer — How to use this workbook", True),
    ("", False),
    ("--- INPUTS SHEET ---", True),
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
    ("All cash flow columns in Analysis (J–R) scale with this value.", False),
    ("", False),
    ("Override Price — column D", True),
    ("Optional. Enter a clean price to replace the DividendData market price.", False),
    ("When set, Effective Price (Analysis col E) uses this value instead of the live price.", False),
    ("Leave blank to use the live market price.", False),
    ("", False),
    ("Override Yield % — column E", True),
    ("Optional. Enter a yield (in %) to replace the DividendData market yield.", False),
    ("When set, Effective Yield (Analysis col F) uses this value directly.", False),
    ("Leave blank to use the live market yield.", False),
    ("", False),
    ("--- ANALYSIS SHEET ---", True),
    ("", False),
    ("Yield columns (E, F, G)", True),
    ("Effective Yield % (col F): annualised total return if bought today and held to maturity.", False),
    ("  Combines coupon income AND capital gain/loss as the price returns to £100 par.", False),
    ("  This is NOT the cash you receive in year 1 — year-1 income is just the coupon.", False),
    ("  Coupons do not compound — each payment is the same fixed cash amount.", False),
    ("Approx Retail Ask Yield % (col G): yield at the DMO retail ask (dirty) price.", False),
    ("  Static import from D10B at export time — does not update when overrides change.", False),
    ("  Typically slightly lower than Effective Yield because the DMO retail price is higher.", False),
    ("", False),
    ("Cash flow columns (J–R)", True),
    ("J  Annual Coupon Cash (£)         = Nominal × Coupon % / 100", False),
    ("K  Capital Uplift to Par (£)       = Nominal × (100 − Price) / 100  [CGT-exempt]", False),
    ("L  Approx Gross Cash Gain (£)      = J × Years + K", False),
    ("M  Coupon Tax @20% (£)             = J × Years × 20%", False),
    ("N  Coupon Tax @40% (£)             = J × Years × 40%", False),
    ("O  Coupon Tax @45% (£)             = J × Years × 45%", False),
    ("P  Approx Net Cash Gain @20% (£)   = L − M  [capital uplift is never taxed]", False),
    ("Q  Approx Net Cash Gain @40% (£)   = L − N", False),
    ("R  Approx Net Cash Gain @45% (£)   = L − O", False),
    ("", False),
    ("Annual Net columns (S, T, U — shown in blue)", True),
    ("S  Annual Net @20% = P ÷ Years to Maturity", False),
    ("T  Annual Net @40% = Q ÷ Years to Maturity", False),
    ("U  Annual Net @45% = R ÷ Years to Maturity", False),
    ("These are the primary comparison figures — after-tax cash per year, normalised for", False),
    ("  duration so gilts with different maturities are on a level playing field.", False),
    ("Without dividing by years, a 30-year gilt always looks better in total than a 2-year", False),
    ("  gilt, regardless of quality. Annual Net fixes this.", False),
    ("Note: Annual Net is a comparison metric — it is not the actual cash you receive each year.", False),
    ("  Actual annual cash = J (coupon only). Capital uplift arrives as a lump sum at maturity.", False),
    ("", False),
    ("How to rank gilts using the Analysis sheet", True),
    ("1. Click the autofilter arrow on column T (Annual Net @40%) and sort largest to smallest.", False),
    ("   (Use S for basic rate 20%, U for additional rate 45%.)", False),
    ("2. Use the Years to Maturity autofilter to restrict to your preferred holding window.", False),
    ("3. Among similarly-ranked gilts, prefer higher Capital Uplift (col K) — tax-free gain.", False),
    ("4. To rank by best pre-tax yield instead, sort column F (Effective Yield %) descending.", False),
    ("These columns update live when you change Nominal Amount or Override Price in Inputs.", False),
    ("", False),
    ("--- SETTINGS SHEET ---", True),
    ("", False),
    ("DefaultNominalAmount (row 2)", True),
    ("Change this single cell to set the default £ holding size for all gilts.", False),
    ("All gilts where Inputs col C is blank will use this value.", False),
    ("Tax scenario rates (rows 5–7) are reference only — Analysis uses hardcoded 20/40/45%.", False),
    ("", False),
    ("--- SELLING BEFORE MATURITY ---", True),
    ("", False),
    ("You do not have to hold a gilt to maturity. If you sell early:", False),
    ("  - All coupons already paid are yours to keep.", False),
    ("  - Accrued interest since the last coupon is paid by the buyer automatically.", False),
    ("  - Your capital outcome depends on the price at the date you sell.", False),
    ("    If rates have risen, the price will be lower — potential capital loss.", False),
    ("    If rates have fallen, the price will be higher — capital gain (still CGT-exempt).", False),
    ("To model an early exit: enter your expected sale price in Override Price (Inputs col D).", False),
    ("  The Analysis sheet will recalculate using that price instead of £100 par.", False),
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
