# Gilt Tax Analyzer — Implementation Progress

## Current State

The project is now an early working application rather than just a plan.

What works today:

- parses real DMO D1A XML exports for conventional gilts
- enriches those rows with live DividendData clean prices and yield-to-maturity values
- parses DMO D10B retail purchase/sale-price workbooks
- derives an approximate retail ask yield from DMO sale dirty price
- generates an Excel workbook with:
  - `Market Data`
  - `Inputs`
  - `Tax Scenarios`
  - `Analysis`
  - `Summary`
- supports manual price and yield overrides
- works with older Excel versions by using `INDEX` / `MATCH` / `IFERROR`
- exports a live workbook from saved DMO XML plus live quote enrichment

Validated against a real DMO XML export:

```text
70 conventional gilts
70 matched live prices
70 retail ask-yield estimates
22 automated tests passing
```

---

## Current Design

### Source Architecture

```text
DMO D1A XML
    -> canonical gilt identity/reference data

DividendData quote feed
    -> clean price and yield-to-maturity enrichment

DMO D10B XLS
    -> retail purchase/sale prices
    -> approximate retail ask-yield derivation

Merged rows
    -> Excel workbook output
```

### Workbook Architecture

| Sheet | Purpose |
|---|---|
| `Market Data` | Imported source values |
| `Inputs` | User-editable nominal amount, price override, yield override |
| `Tax Scenarios` | Editable scenario tax rates |
| `Analysis` | Formula-driven outputs |
| `Summary` | Placeholder for ranking/dashboard logic |

### Override Behavior

For both price and yield:

1. use the user override if non-blank
2. otherwise use the imported market value
3. otherwise leave the value blank

This prevents blank override rows from masking valid imported data.

### Formula Compatibility

The workbook originally used newer Excel functions such as:

- `XLOOKUP`
- `LET`

These produced `#NAME?` errors in the user's Excel installation, so formulas were redesigned around:

- `INDEX`
- `MATCH`
- `IFERROR`

This keeps the workbook broadly compatible while preserving the intended behavior.

---

## Current Assumptions

### Market Data

- DMO D1A is the canonical source for:
  - ISIN
  - gilt name
  - coupon
  - maturity date
  - instrument type
- DividendData is the current practical source for:
  - clean price
  - yield to maturity
- The workbook field `Effective Yield %` currently means:
  - override yield if entered
  - otherwise imported yield to maturity
- `Effective Yield %` is **not the same as dealer ask yield**.
  - True ask-side analytics would require a different source exposing ask quotes explicitly.
- `Approx Retail Ask Yield %` is derived from DMO D10B sale dirty price.
- `Approx Retail Ask Yield %` is a public retail-service proxy, not yet an exact official-convention redemption yield.

### Tax

- Conventional UK gilt capital gains are modeled as CGT-exempt.
- Coupon taxation is modeled using scenario rates:
  - 20%
  - 40%
  - 45%
- The workbook does **not** yet model:
  - personal savings allowance
  - starting rate for savings
  - wider taxable-income interactions
  - investor-specific circumstances

### Calculations

Current calculations are intentionally approximate.

Included:

- coupon cash
- capital uplift to par
- scenario coupon tax
- approximate net cash gain to maturity

Not yet included:

- accrued interest
- dirty price
- exact coupon schedule
- exact after-tax IRR
- exact after-tax yield to maturity

---

## Implemented Features

### Core Code

```text
app/
|-- collectors/
|   |-- base.py
|   |-- d10b.py
|   |-- dividenddata.py
|   `-- dmo.py
|-- excel_builder.py
|-- formulas.py
|-- models.py
|-- tax_scenarios.py
`-- transform.py
```

### CLI Commands

```powershell
python main.py info
python main.py export
python main.py export-xml .\path\to\d1a.xml
python main.py export-xml-with-quotes .\path\to\d1a.xml --output .\output\gilt_analysis.xlsx
python main.py export-xml-with-quotes-and-retail-ask .\path\to\d1a.xml .\path\to\d10b.xls --output .\output\gilt_analysis_with_retail_ask.xlsx
```

### Current Test Coverage

The suite currently covers:

- DMO XML parsing
- Unicode and ASCII coupon-fraction parsing
- DividendData quote parsing
- D10B workbook parsing
- malformed quote-row handling
- quote enrichment and mismatch handling
- approximate retail ask-yield plausibility
- workbook structure
- workbook formulas
- normalization behavior
- defensive currency parsing

Current result:

```text
22 passed
```

---

## Completed Milestones

### Planning

- reviewed the original implementation plan
- wrote `implementation_plan_v2.md`
- reframed the project as an after-tax scenario analyzer rather than a full tax engine

### Environment

- created local `.venv`
- installed project dependencies
- verified local PowerShell workflow

### Data Ingestion

- implemented DMO D1A XML parsing
- added local XML fallback because live DMO requests can be blocked by anti-bot protection
- implemented DividendData quote collection
- matched real-world rows successfully across both sources
- implemented DMO D10B `.xls` parsing
- added reproducible `.xls` support via `xlrd`

### Workbook / Formula Layer

- built workbook skeleton and formula pipeline
- added effective price and effective yield logic
- added approximate retail ask yield to the Analysis sheet
- made default nominal amount configurable
- added Summary-sheet placeholder
- migrated formulas away from unsupported newer Excel functions

### Robustness Improvements

- added malformed-row handling for quote parsing
- added warning logs for unmatched rows
- clarified market-row vs quote-row collector abstractions
- documented coupon-rate scale convention
- added contextual parsing errors for malformed coupon strings

---

## Known Limitations

1. **No true ask-side feed**
   - current quote enrichment gives clean price and YTM
   - D10B adds an approximate retail ask-yield proxy
   - true dealer ask price / ask yield still require a different source

2. **Live DMO fetch remains unreliable**
   - browser-saved XML is the dependable path today

3. **Summary sheet is still a placeholder**
   - no ranking or dashboard logic yet

4. **Analytics remain approximate**
   - no accrued interest, dirty price, exact coupon schedule, or exact after-tax IRR yet

---

## Next Priorities

| Priority | Action |
|---|---|
| 1 | Improve workbook presentation: widths, number formats, conditional formatting, notes |
| 2 | Build the `Summary` sheet with rankings and maturity buckets |
| 3 | Surface source transparency: valuation date, source labels, unmatched-row diagnostics |
| 4 | Add stronger validation: duplicate checks, missing-price checks, cross-source consistency checks |
| 5 | Investigate a true ask-side data source if ask yield becomes a product requirement |
| 6 | Refine `Approx Retail Ask Yield %` toward official DMO convention if higher precision becomes important |

---

## Code Review Session — 2026-05-16 (D10B Review)

### Context

This session specifically reviews the D10B implementation: the parser, the retail ask yield calculation, the model, and how the output is integrated into the workbook.

---

### What D10B Does

The DMO D10B workbook is a daily Excel file published by the UK Debt Management Office containing the retail Purchase and Sale Service prices for each conventional gilt. These are the prices at which retail investors can transact through the DMO's own service — distinct from market maker prices.

The D10B data added to this project:

- `GiltRetailQuote` model — captures ISIN, gilt name, purchase/sale clean and dirty prices, redemption date, data date
- `parse_d10b_xls` — reads the raw `.xls` file via pandas, skips the header block, parses data rows
- `approximate_retail_ask_yield` — estimates the annual redemption yield implied by the sale dirty price using a bisection solver over a semi-annual coupon cashflow model
- `Approx Retail Ask Yield %` column in the `Analysis` sheet — written as a static float per ISIN, not an Excel formula

---

### What Is Working Well

- **Model is clean and well-typed.** `GiltRetailQuote` uses `Decimal` throughout, is `frozen=True`, and carries both clean and dirty prices — the right granularity for this source.
- **`parse_d10b_xls` is robust to date types.** `_coerce_date` handles `datetime`, `date`, and pandas `Timestamp` — all three formats can appear depending on the Excel file and pandas version.
- **Bisection solver is correct in structure.** Iterating 120 rounds gives convergence to well under a basis point. Using semi-annual coupons (`coupon_rate / 2`, `**2*years`) is the right convention for UK gilts.
- **`_subtract_six_months` handles month-end correctly.** Day clamping via `_days_in_month` prevents invalid dates like 31 November.
- **Test `test_parse_d10b_xls_real_file` pins real values.** Exact ISIN, price, and date are asserted — this will catch any silent parse regression.

---

### Issues Found

#### 1. `approximate_retail_ask_yield` uses `365.25` day count — not the standard UK gilt convention

[app/collectors/d10b.py:60](app/collectors/d10b.py#L60):

```python
years = Decimal((coupon_date - valuation_date).days) / Decimal('365.25')
```

UK gilts use the **Actual/Actual (ICMA)** day count convention for yield calculations, not `365.25`. Using `365.25` introduces a small but systematic bias, particularly for gilts with coupons straddling a leap year. The progress doc correctly labels this column `Approx Retail Ask Yield %` — so the approximation is acknowledged — but the limitation should be explicitly noted in the function docstring so it is not forgotten when precision matters.

Action: add a note to the `approximate_retail_ask_yield` docstring that the `365.25` day count is a simplification and that Actual/Actual ICMA would be needed for exact DMO-convention results.

---

#### 2. Bisection solver lower bound of `-0.99` is theoretically sound but undocumented

[app/collectors/d10b.py:38-39](app/collectors/d10b.py#L38-L39):

```python
low = Decimal('-0.99')
high = Decimal('1.00')
```

The bounds `[-99%, +100%]` cover all realistic gilt yield scenarios. However, if a malformed dirty price is passed — e.g. `0` or a very large number — the solver will silently converge to `-0.99` or `1.00` and return a nonsense result with no error. There is no guard on whether the result is in a plausible range after convergence.

Action: after the bisection loop, assert or raise if the result falls outside a plausible band (e.g. `-5%` to `20%` covers all realistic UK gilt yields with margin).

---

#### 3. `parse_d10b_xls` uses positional column indices — fragile to layout changes

[app/collectors/d10b.py:24-28](app/collectors/d10b.py#L24-L28):

```python
purchase_clean_price=Decimal(str(row[2])),
purchase_dirty_price=Decimal(str(row[3])),
sale_clean_price=Decimal(str(row[4])),
sale_dirty_price=Decimal(str(row[5])),
...
redemption_date=_coerce_date(row[7]),
```

The parser assumes a fixed column layout: ISIN at 0, name at 1, purchase clean at 2, purchase dirty at 3, sale clean at 4, sale dirty at 5, (skipping index 6), redemption date at 7. If the DMO adds or reorders a column — which they have done historically — every parsed value after the change would be silently wrong.

Action: read the header row and locate columns by name rather than by index. The D10B header row is present in the file; using it would make the parser robust to column reordering.

---

#### 4. `retail_ask_yields` is passed to `build_workbook` as `dict[str, float]` but values are written as static cells

[app/excel_builder.py:136](app/excel_builder.py#L136):

```python
sheet.cell(excel_row, 7, retail_ask_yields.get(row.isin))
```

The retail ask yield is written as a plain Python float directly into the cell — not an Excel formula. This means:

- it cannot be recalculated if the user changes the override price
- it is not transparent to the user (no formula bar to inspect)
- it is not linked to the `Market Data` sheet

This is a known trade-off given the complexity of the bisection solver in Excel. But the column header `Approx Retail Ask Yield %` gives no indication that this value is a static import. A user adjusting nominal amount or price would not realize this column is frozen.

Action: add a visible note in the `Market Data` sheet or a workbook-level comment clarifying that the retail ask yield column is a pre-computed import and does not update when overrides change.

---

#### 5. Test `test_parse_d10b_xls_real_file` depends on a file in `data/` that is not committed

[tests/test_d10b_collector.py:9](tests/test_d10b_collector.py#L9):

```python
quotes = parse_d10b_xls(Path("data/20260515 - DMO Gilt Purchase and Sale Service Prices.xls"))
```

This test will fail on any machine that does not have this specific file at this path. It is a real-data integration test rather than a unit test, which is fine — but it is not marked as such (e.g. with `pytest.mark.skip` or a custom marker). Any CI run or collaborator without the file will see an unexplained failure.

The plausibility test (`test_approximate_retail_ask_yield_is_plausible`) is self-contained and does not have this problem.

Action: either add a `pytest.mark.skipif(not Path(...).exists(), reason="...")` guard, or extract the file-dependent assertion into a separate `tests/integration/` directory that is excluded from the default run.

---

#### 6. `GiltRetailQuote` has no `coupon_rate` field

[app/models.py:46-55](app/models.py#L46-L55):

```python
@dataclass(frozen=True)
class GiltRetailQuote:
    isin: str
    gilt_name: str
    purchase_clean_price: Decimal
    purchase_dirty_price: Decimal
    sale_clean_price: Decimal
    sale_dirty_price: Decimal
    redemption_date: date
    data_date: date
```

`approximate_retail_ask_yield` requires `coupon_rate` as an argument, which means the caller must look it up from `GiltMarketRow` separately and pass it in. The coupon rate is present in the D10B file (implicitly via the gilt name, and possibly as a column). Not storing it in `GiltRetailQuote` means the model is incomplete and the caller must join two data sources to do the yield calculation.

Action: either add `coupon_rate` to `GiltRetailQuote` if D10B provides it, or document explicitly why the join is the caller's responsibility.

---

### Priority Order for Next Actions

| Priority | Action |
|---|---|
| 1 | **Guard `parse_d10b_xls` column lookup by header name** — prevent silent wrong-column parsing on layout changes |
| 2 | **Add plausibility guard to bisection result** — raise if yield outside realistic band after convergence |
| 3 | **Mark `test_parse_d10b_xls_real_file` as conditional** — skip gracefully when the data file is absent |
| 4 | **Document `365.25` limitation in `approximate_retail_ask_yield`** — one-line note in the docstring |
| 5 | **Add workbook note that retail ask yield column is a static import** — prevents user confusion when overrides change |
| 6 | **Decide on `coupon_rate` in `GiltRetailQuote`** — add the field or document the join requirement |

---

## D10B Review Fixes Applied — 2026-05-16

All six D10B review issues have been fixed and verified at 23 passed.

| Fix | Change |
|---|---|
| Column indices hardcoded | Parser now locates the header row dynamically and maps columns by normalised name — robust to added/reordered columns and multi-line header cells |
| Footnote rows crash parser | Added `_is_valid_isin` guard: rows whose first cell does not match a 12-character alphanumeric ISIN pattern are silently skipped |
| No plausibility guard on bisection | `approximate_retail_ask_yield` now raises `ValueError` with full context if result falls outside −5% to +20% |
| `365.25` day count undocumented | Added explicit note to docstring that this is a simplification and Actual/Actual ICMA would be needed for exact DMO-convention precision |
| Real-file test fails without data file | Test now uses `pytest.mark.skipif(not path.exists(), ...)` — skips gracefully on any machine without the file |
| `GiltRetailQuote` missing `coupon_rate` | Field added to model with same percentage-point convention as `GiltMarketRow`; parser derives it from the gilt name using the same fraction-parsing logic as the DMO collector |
| Workbook gives no hint retail ask yield is static | `_write_headers` now attaches an Excel comment to the `Approx Retail Ask Yield %` header cell explaining it is a pre-computed import and does not update when overrides change |

New tests added:

- `test_approximate_retail_ask_yield_raises_on_implausible_result`
- `test_parse_d10b_xls_real_file` updated with `coupon_rate` assertion

Current automated test status:

```text
23 passed
```

---

## Historical Notes

### Data-source evolution

1. **DMO only**
   - authoritative for reference data
   - not enough for analysis because D1A lacks prices and yields

2. **DMO + local XML fallback**
   - introduced after live DMO requests sometimes returned anti-bot pages

3. **DMO + DividendData**
   - adopted as the first workable two-source design
   - validated on real data with full price-match coverage

4. **DMO + DividendData + D10B**
   - added official retail purchase/sale prices
   - introduced an approximate retail ask-yield lens alongside general market YTM

### Important implementation lessons

- Real DMO exports use both Unicode and ASCII fraction notation:
  - `0⅛%`
  - `0 3/8%`
- Excel compatibility matters more than elegant formulas if the user cannot open the workbook cleanly.
- A price-enrichment path must fail loudly if a quote page is malformed; silent blanks are more dangerous than visible errors.
- Public retail sale price can support a useful ask-side proxy before true institutional ask-yield data is available.

---

## Review Log

### Session 1

Resolved:

- lack of price source
- `effective_price` fallback flaws
- undocumented coupon-rate scale
- hardcoded nominal amount
- blank Summary tab
- unclear normalization responsibilities

### Session 2

Resolved:

- unsafe silent failure in `DividendDataCollector.fetch`
- lack of unmatched-row visibility
- incomplete net-gain formula assertions
- currency parsing edge case
- ambiguous collector abstraction
- stale test-count reporting

### Session 3

Resolved:

- malformed quote rows aborting the entire parse
- bare-dash currency parse failures
- missing context in coupon-parse errors
- stale formula description in project notes
- non-idiomatic exception test style
- missing coupon-mismatch regression coverage

### Session 4

Added:

- D10B `.xls` ingestion
- `GiltRetailQuote`
- `Approx Retail Ask Yield %`
- CLI export path combining:
  - saved D1A XML
  - live DividendData quotes
  - downloaded D10B workbook
- regression coverage for:
  - D10B parsing
  - retail ask-yield plausibility

---

## Current Working Command

```powershell
python main.py export-xml-with-quotes-and-retail-ask .\data\d1a.xml ".\data\20260515 - DMO Gilt Purchase and Sale Service Prices.xls" --output .\output\gilt_analysis_with_retail_ask.xlsx
```

Expected current result:

```text
Wrote 70 conventional gilts to output\gilt_analysis_with_retail_ask.xlsx with 70 retail ask-yield estimates
```

# Workbook Analysis Review — 2026-05-16

## Context

This review evaluates the generated Excel workbook, specifically the `Analysis` sheet, focusing on:

- financial correctness
- formula integrity
- tax modeling
- alignment with market conventions
- suitability for retail-investor analysis

The workbook architecture is strong overall and already suitable for an MVP retail-analysis tool. However, several calculation limitations and naming ambiguities were identified that should be clarified before positioning the workbook as a professional fixed-income analytics engine.

---

## What Is Working Well

### Formula Architecture

The workbook correctly uses:

- native Excel formulas
- dynamic recalculation
- editable workbook logic
- override-driven market inputs

This is the correct architectural direction.

---

### Tax Modeling

The current tax logic is broadly correct for:

- UK resident retail investors
- direct ownership of conventional UK gilts

Correctly modeled:

- coupon taxation
- CGT exemption on conventional gilts
- scenario-based tax rates

---

### Capital Uplift Logic

The current formula:

```excel
=Nominal*(100-Price)/100
```

correctly models:

- pull-to-par capital appreciation
- redemption at par
- simplified clean-price capital uplift

This is appropriate for the current MVP.

---

## Issues Found

### 1. Approximate Gross Gain calculation is not true YTM

Current formula:

```excel
=AnnualCoupon*YearsToMaturity + CapitalGain
```

This is only a simplified investor estimate.

It does NOT model:

- exact yield-to-maturity
- IRR
- accrued interest
- coupon schedules
- dirty price
- coupon reinvestment

Recommendation:

Rename approximate-return columns to avoid implying exact YTM precision.

---

### 2. Effective Yield naming is ambiguous

The workbook currently mixes:

- imported market YTM
- override yields
- retail ask-yield approximations

The terminology may imply institutional-grade ask-side analytics when the calculations are actually simplified.

Recommendation:

Prefer names such as:

- Market Yield to Maturity %
- Imported Market YTM %

---

### 3. Approx Retail Ask Yield is static

The retail ask yield is currently written as a static imported value rather than an Excel formula.

This is acceptable because the calculation uses iterative numerical solving.

However, users may incorrectly assume the value recalculates when override prices change.

Recommendation:

Add workbook comments or notes clearly stating that the value is static and imported.

---

### 4. YEARFRAC approximation differs from official gilt convention

The workbook currently uses simplified year-fraction approximations.

Official UK gilt calculations use:

- Actual/Actual (ICMA)

This introduces small differences around leap years and coupon boundaries.

Recommendation:

Document the simplification explicitly.

---

### 5. Purchase Cost column is missing

The workbook currently does not explicitly expose:

```excel
=Nominal*Price/100
```

Without purchase cost:

- investor cash outlay is less visible
- return percentages become harder to interpret

Recommendation:

Add:

```text
Estimated Purchase Cost (£)
```

---

### 6. Workbook mixes professional and simplified metrics

The workbook currently combines:

- imported market YTM
- simplified investor return estimates
- retail ask-yield approximations

This is acceptable for MVP retail analysis but should not yet be positioned as a professional fixed-income pricing engine.

---

## Recommended Immediate Improvements

| Priority | Action |
|---|---|
| 1 | Rename approximate return columns |
| 2 | Clarify static retail ask-yield behavior |
| 3 | Add Purchase Cost column |
| 4 | Document Actual/Actual vs YEARFRAC simplification |
| 5 | Keep market YTM and investor estimates conceptually separate |

---

## Overall Assessment

| Area | Assessment |
|---|---|
| Workbook structure | Strong |
| Formula architecture | Strong |
| Tax scenario modeling | Good |
| Retail-investor suitability | Good |
| Market-pricing precision | Approximate |
| Professional fixed-income accuracy | Not yet implemented |

---

## Final Recommendation

The workbook is already a strong MVP for:

- retail gilt analysis
- after-tax comparison
- scenario modeling
- finance/data-engineering portfolio demonstration

The next major step toward institutional-grade analytics would require:

- dirty price handling
- accrued interest
- exact coupon schedules
- Actual/Actual ICMA conventions
- duration and convexity
- QuantLib integration

---

## Workbook Analysis Review — 2026-05-16

### Context

This session reviews the workbook layer end-to-end: formula correctness, column layout, number formatting, usability, and structural issues. The generated file `output/gilt_analysis.xlsx` was the reference point alongside the current source in `app/excel_builder.py` and `app/formulas.py`.

---

### Formula Correctness Audit

Full column trace for the `Analysis` sheet (19 columns, data from row 2):

| Col | Letter | Header | Formula / Content | Verdict |
|---|---|---|---|---|
| 1 | A | ISIN | static | Correct |
| 2 | B | Gilt Name | static | Correct |
| 3 | C | Maturity | static `date` | Correct |
| 4 | D | Coupon % | `float(coupon_rate)` — percentage points (e.g. `1.5`) | Correct |
| 5 | E | Effective Price | `IFERROR(IF(override<>"", override, imported), "")` | Correct |
| 6 | F | Effective Yield % | same pattern as E, from Inputs!D | Correct |
| 7 | G | Approx Retail Ask Yield % | static float from D10B bisection | Correct — static import |
| 8 | H | Nominal Amount (£) | `IFERROR(INDEX(Inputs!B:B, MATCH(...)), "")` | Correct |
| 9 | I | Years to Maturity | `IFERROR(YEARFRAC(valuation_date, C{row}, 1), "")` | Correct |
| 10 | J | Annual Coupon Cash (£) | `=H{row}*D{row}/100` | Correct — H×D/100 = £ coupon |
| 11 | K | Capital Uplift to Par (£) | `=H{row}*(100-E{row})/100` | Correct |
| 12 | L | Approx Gross Cash Gain (£) | `=J{row}*I{row}+K{row}` | Correct — undiscounted |
| 13 | M | Coupon Tax @20% (£) | `=J{row}*I{row}*20%` | Correct |
| 14 | N | Coupon Tax @40% (£) | `=J{row}*I{row}*40%` | Correct |
| 15 | O | Coupon Tax @45% (£) | `=J{row}*I{row}*45%` | Correct |
| 16 | P | CGT (£) | `=0` | Correct — CGT-exempt |
| 17 | Q | Net @20% (£) | `=L{row}-M{row}-P{row}` | Correct |
| 18 | R | Net @40% (£) | `=L{row}-N{row}-P{row}` | Correct |
| 19 | S | Net @45% (£) | `=L{row}-O{row}-P{row}` | Correct |

**All formulas are arithmetically correct.** No column reference errors found.

---

### Issues Found

#### 1. No number formatting applied to any cell

The workbook writes all values as raw Python types. Excel will display:

- dates as integers (serial numbers) if the column is not date-formatted
- prices as plain decimals with no alignment (e.g. `93.2` instead of `93.200`)
- £ amounts as unformatted numbers (e.g. `1500` instead of `£1,500.00`)
- percentage columns (Coupon %, Effective Yield %) as raw decimals

The implementation plan (`implementation_plan_v2.md` §11) specified:

- Currency: `£#,##0.00`
- Percentage: `0.000%`
- Date: `dd-mmm-yyyy`

None of these are currently applied.

Action: add `NumberFormat` styles to `_populate_market_data`, `_populate_analysis`, and `_populate_inputs` for the relevant columns.

---

#### 2. No column widths set — all columns default to 8.43 characters

Headers like `"Approx Gross Cash Gain to Maturity (£)"` are truncated to `########` at default width. Users cannot read any column without manually resizing.

Action: call `sheet.column_dimensions[letter].width = N` after populating each sheet. Suggested widths: ISIN 16, Gilt Name 36, Maturity 14, currency columns 18, percentage columns 12.

---

#### 3. `autofilter` and `freeze_panes` applied to the `Summary` sheet

[app/excel_builder.py:82-84](app/excel_builder.py#L82-L84):

```python
for sheet in workbook.worksheets:
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
```

The Summary sheet currently contains only `A1` with a placeholder label. `sheet.dimensions` resolves to `"A1:A1"`, so the autofilter is applied to a single cell. `freeze_panes = "A2"` on a one-row sheet freezes nothing useful. Both are harmless but will produce a confusing autofilter dropdown arrow on the placeholder label.

Action: exclude the Summary sheet from the loop, or only apply autofilter when the sheet has more than one row of data.

---

#### 4. `GiltUserInput` model is defined but never used

[app/models.py:26-31](app/models.py#L26-L31):

```python
@dataclass(frozen=True)
class GiltUserInput:
    isin: str
    nominal_amount: Decimal
    override_price: Decimal | None = None
    override_yield: Decimal | None = None
```

The `Inputs` sheet is populated directly from `GiltMarketRow` rows (writing the ISIN and a default nominal). `GiltUserInput` is never instantiated or imported anywhere. It was presumably intended as the model for reading user edits back from the workbook, but that read-back path has not been built.

Action: either remove the model if the read-back feature is not planned, or document it as a placeholder for the future import-overrides-from-Excel feature.

---

#### 5. `Inputs` ISIN column has no protection

The `Inputs` sheet ISIN column (A) is the MATCH key used by every formula in the `Analysis` sheet. If a user accidentally overwrites an ISIN, all lookups for that gilt silently return `""` — the analysis row goes blank with no error.

Action: when workbook protection is applied in a future phase, lock column A of the `Inputs` sheet while leaving columns B–D editable.

---

#### 6. `coupon_tax` and `gross_cash_gain` formulas double-count the MATCH lookup

[app/formulas.py:37-42](app/formulas.py#L37-L42):

```python
def approx_gross_cash_gain_to_maturity(row: int) -> str:
    return f"=J{row}*I{row}+K{row}"

def coupon_tax(row: int, tax_rate: str) -> str:
    return f"=J{row}*I{row}*{tax_rate}"
```

Both formulas reference `J{row}` (Annual Coupon Cash) and `I{row}` (Years to Maturity), which are themselves formula cells that each perform an `INDEX/MATCH` lookup against two sheets. Excel recalculates these each time any dependent cell changes, so with 70 rows and 5 formulas referencing J and I, that is 350 redundant lookups per recalculation cycle.

This is not a correctness issue, but it will cause slow recalculation on large workbooks. The clean solution is to use intermediate named ranges or helper columns, but for 70 rows the impact is negligible today.

Action: no action needed now — note for Phase 3 if the row count grows significantly.

---

### What Is Working Well

- All 19 formula cells are arithmetically correct and reference the right columns.
- The `IFERROR` wrapper on all lookup formulas means missing data produces `""` rather than `#N/A` error propagation.
- The override logic (check override first, fall back to imported) is correctly implemented for both price and yield.
- The `Approx Retail Ask Yield %` column correctly carries a header comment flagging it as a static import.
- The workbook structure (five sheets, correct sheet order) matches the design specification.

---

### Priority Order for Next Actions

| Priority | Action |
|---|---|
| 1 | **Apply number formatting** — `£#,##0.00` for currency, `0.000%` for percentages, `dd-mmm-yyyy` for dates |
| 2 | **Set column widths** — prevent `########` truncation on all sheets |
| 3 | **Exclude Summary from autofilter/freeze loop** — remove confusing dropdown on placeholder label |
| 4 | **Remove or document `GiltUserInput`** — resolve the dead model ambiguity |
| 5 | **Note Inputs ISIN lock** — add to the workbook protection task list for a future phase |
