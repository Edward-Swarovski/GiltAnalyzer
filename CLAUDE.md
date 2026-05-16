# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```powershell
# Activate environment (required before any python command)
.\.venv\Scripts\Activate.ps1

# Run all tests
.\.venv\Scripts\python.exe -m pytest tests\ -v

# Run a single test
.\.venv\Scripts\python.exe -m pytest tests\test_d10b_collector.py::test_approximate_retail_ask_yield_is_plausible -v

# Generate workbook (standard workflow — requires files in data/)
.\.venv\Scripts\python.exe main.py export-xml-with-quotes-and-retail-ask .\data\d1a.xml ".\data\20260515 - DMO Gilt Purchase and Sale Service Prices.xls" --output .\output\gilt_analysis.xlsx

# Generate workbook without retail ask yield
.\.venv\Scripts\python.exe main.py export-xml-with-quotes .\data\d1a.xml --output .\output\gilt_analysis.xlsx
```

Expected test result: **23 passed**. The test `test_parse_d10b_xls_real_file` is skipped automatically when the real D10B file is absent.

## Architecture

### Data flow

```
DMO D1A XML (saved locally)  ->  parse_d1a_xml()       ->  list[GiltMarketRow]
                                                                    |
DividendData (live HTTP)     ->  DividendDataCollector  ->  list[GiltPriceQuote]
                                                                    |
                                                         enrich_with_quotes()
                                                                    |
DMO D10B XLS (saved locally) ->  parse_d10b_xls()      ->  dict[isin, retail_ask_yield]
                                                                    |
                                                         build_workbook()
                                                                    |
                                                         output/*.xlsx
```

### Module roles

- **`app/models.py`** — four frozen dataclasses: `GiltMarketRow` (primary canonical row), `GiltPriceQuote` (DividendData quote), `GiltRetailQuote` (DMO D10B row), `GiltUserInput` (placeholder — not yet used).
- **`app/collectors/base.py`** — two ABC bases: `BaseMarketRowCollector` (→ `list[GiltMarketRow]`) and `BasePriceQuoteCollector` (→ `list[GiltPriceQuote]`). These are not interchangeable.
- **`app/collectors/dmo.py`** — parses DMO D1A XML. Produces identity/reference data only — no prices. Handles Unicode (`½`, `⅛`) and ASCII (`1 5/8%`) coupon fractions.
- **`app/collectors/dividenddata.py`** — live HTML scrape of DividendData gilt table. `enrich_with_quotes()` merges by `(normalised_name, maturity_date_iso, coupon_rate)` key. Logs a WARNING per unmatched row.
- **`app/collectors/d10b.py`** — parses DMO Purchase and Sale Service XLS. Locates the header row dynamically (robust to column reordering). `approximate_retail_ask_yield()` uses a 120-iteration bisection over semi-annual cashflows; raises `ValueError` if result is outside −5% to +20%.
- **`app/formulas.py`** — pure functions returning Excel formula strings. All formulas use `INDEX`/`MATCH`/`IFERROR` — `XLOOKUP` and `LET` are excluded because they cause `#NAME?` on the target Excel installation.
- **`app/excel_builder.py`** — assembles the workbook. `build_workbook()` accepts `retail_ask_yields: dict[str, float] | None`; retail ask yield is written as a static value (not a formula) because the bisection cannot be expressed in Excel.
- **`app/transform.py`** — `normalize_market_row()` for dict-based raw input. `coupon_rate` must be supplied in percentage points by the caller (e.g. `"1.5"` for 1.5%).

### Key conventions

- **`coupon_rate` is always in percentage points** throughout the codebase: `Decimal("1.5")` = 1.5% gilt, `Decimal("0.125")` = 0.125% gilt. The formula `=H*D/100` in the Analysis sheet depends on this.
- **DMO D1A provides identity only** — `imported_price` and `imported_yield` are always `None` after D1A parsing. Prices come exclusively from DividendData enrichment.
- **Live DMO fetch is unreliable** — the `DmoCollector.fetch()` live path is often blocked by anti-bot protection. The standard workflow saves D1A XML from a browser and uses `export-xml-with-quotes` or `export-xml-with-quotes-and-retail-ask`.

### Workbook structure

The Analysis sheet has 19 columns (A–S). Column letters are hardcoded in `formulas.py`. If columns are added or reordered, formula strings and test assertions in `test_excel_builder.py` must both be updated.

| Key columns | Letter |
|---|---|
| Effective Price | E |
| Nominal Amount | H |
| Years to Maturity | I |
| Annual Coupon Cash | J |
| Capital Uplift to Par | K |
| CGT (always =0) | P |
| Net @20/40/45% | Q/R/S |

### Input files (not committed)

Place locally saved DMO files in `data/` before running export commands:

- `data/d1a.xml` — DMO D1A XML report (save from browser)
- `data/<date> - DMO Gilt Purchase and Sale Service Prices.xls` — DMO D10B workbook
