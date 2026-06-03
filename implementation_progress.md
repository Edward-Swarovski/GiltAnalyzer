# Gilt Tax Analyzer — Implementation Progress

## Current State (as of 2026-05-16)

A fully working retail gilt analysis tool. Generates a multi-sheet Excel workbook from two daily DMO files plus a live price feed.

Current automated test status: **28 passed**

Current working command:

```powershell
.\.venv\Scripts\Activate.ps1
python main.py export-xml-with-quotes-and-retail-ask .\data\d1a.xml --output .\output\gilt_analysis.xlsx
```

Expected output:

```text
Auto-detected D10B file: data\YYYYMMDD - DMO Gilt Purchase and Sale Service Prices.xls
Wrote 70 conventional gilts to output\gilt_analysis.xlsx with 70 retail ask-yield estimates
```

---

## Data Flow

```text
DMO D1A XML (saved manually from browser — bot-blocked)
    -> canonical gilt identity: ISIN, coupon, maturity, instrument type

DividendData (live HTTP fetch)
    -> clean price + yield-to-maturity enrichment

DMO D10B XLS (saved manually from https://www.dmo.gov.uk/data/pdfdatareport?reportCode=D10B)
    -> retail purchase/sale dirty prices
    -> approximate retail ask yield (bisection over semi-annual cashflows)

Merged rows
    -> Excel workbook (output/gilt_analysis.xlsx)
```

---

## Workbook Structure (5 sheets)

| Sheet | Purpose | Editable |
|---|---|---|
| Settings | DefaultNominalAmount control cell + tax scenario reference | Yes |
| Market Data | Raw imported values (ISIN, price, yield, maturity, source) | No |
| Inputs | Per-gilt overrides: nominal amount, price, yield | Yes |
| Analysis | 23-column formula-driven cash flow outputs — sort/filter here | No |
| Instructions | How to use the workbook | No |

---

## Analysis Sheet Columns (A–W)

| Col | Header | Type |
|---|---|---|
| A | ISIN | Static |
| B | Gilt Name | Static |
| C | Maturity | Static |
| D | Coupon % | Static |
| E | Effective Price | Formula (override → market mid price) |
| F | Effective Yield % | Formula (override → market yield) |
| G | Approx Retail Ask Yield % | Static from D10B dirty price — does not update with overrides |
| H | DMO Retail Sale Clean Price | Static from D10B — N/A if no D10B loaded |
| I | DMO Retail Sale Dirty Price | Static from D10B — actual settlement price; N/A if no D10B loaded |
| J | Nominal Amount (£) | Formula (Inputs col C → Settings DefaultNominalAmount) |
| K | Years to Maturity | Formula: YEARFRAC(valuation_date, C, 1) |
| L | Annual Coupon Cash (£) | Formula: J × D / 100 |
| M | Capital Uplift to Par (£) | Formula: J × (100 − E) / 100 — CGT-exempt |
| N | Approx Gross Cash Gain to Maturity (£) | Formula: L × K + M |
| O | Coupon Tax @20% (£) | Formula: L × K × 20% |
| P | Coupon Tax @40% (£) | Formula: L × K × 40% |
| Q | Coupon Tax @45% (£) | Formula: L × K × 45% |
| R | Approx Net Cash Gain @20% (£) | Formula: N − O |
| S | Approx Net Cash Gain @40% (£) | Formula: N − P |
| T | Approx Net Cash Gain @45% (£) | Formula: N − Q |
| U | Annual Net @20% (£) | Formula: R / K — shown in blue |
| V | Annual Net @40% (£) | Formula: S / K — shown in blue |
| W | Annual Net @45% (£) | Formula: T / K — shown in blue |

All formulas use `INDEX`/`MATCH`/`IFERROR` — `XLOOKUP` and `LET` excluded for Excel compatibility.

---

## Summary — Best Value Columns

Two sets of per-£10k figures to correctly account for the fact that gilts trade below par:

| Column | What it measures |
|---|---|
| Cash Invested per £10k nominal | Actual spend: Price × 100 (e.g. £9,320 at price 93.20) |
| Annual Net @20/40/45% per £10k nominal | Return per £10k face value |
| Annual Net @20/40/45% per £10k cash invested | Return per £10k actually spent — correct comparison metric |
| Total Net @20/40/45% per £10k nominal | Total cash over full life per £10k face value |
| Capital Uplift to Par per £10k nominal | Tax-free capital gain |

---

## Key Design Decisions

### YEARFRAC matching Excel exactly
Python `_yearfrac()` implements the exact Excel `YEARFRAC(basis=1)` algorithm: average the days-per-year across all calendar years the period touches (start year through end year inclusive), then divide actual days by that average. This ensures Summary snapshot figures match Analysis formula outputs exactly.

### Nominal vs cash-invested normalisation
Summary — Best Value exposes both "per £10k nominal" and "per £10k cash invested" columns. The cash-invested figure is the correct return-on-capital metric because discount gilts (price < 100) cost less than their nominal value to purchase.

### Annual Net as the primary comparison metric
Columns T/U/V divide the total net gain by years to maturity, making gilts with different durations directly comparable. Without duration normalisation, longer-dated gilts always appear to return more in total regardless of quality.

### D10B auto-detection
`_find_latest_d10b()` globs `data/` for the most recently modified DMO filename pattern — no need to type the exact dated filename.

### Settings sheet as control panel
`DefaultNominalAmount` in Settings!B2 applies to all gilts where Inputs col C is blank. Change one cell to model a different holding size across the entire workbook.

---

## Implemented Features

### Data ingestion
- DMO D1A XML parsing (conventional gilts only, Unicode + ASCII coupon fractions)
- DividendData live HTML scrape with `INDEX/MATCH`-style enrichment by `(name, maturity, coupon)` key
- DMO D10B XLS parsing with dynamic header detection (robust to column reordering)
- Approximate retail ask yield via 120-iteration bisection over semi-annual cashflows
- D10B auto-detection from `data/` directory

### Workbook / formula layer
- 22-column Analysis sheet with live Excel formulas
- Settings sheet DefaultNominalAmount control cell with per-gilt override in Inputs
- Annual Net columns (T/U/V) in blue — duration-normalised after-tax return
- Summary — Yield Ranking sorted by Effective Yield % desc
- Summary — Best Value with per-£10k nominal and per-£10k cash invested columns
- Instructions sheet explaining all sheets, columns, and workflow
- Print setup: fit to 1 page wide, font size 9, word-wrapped headers
- £ columns formatted `£#,##0.00`
- Gilt Name in Inputs sheet (formula from Market Data, read-only)

### CLI
```powershell
python main.py info
python main.py export-xml <xml_path>
python main.py export-xml-with-quotes <xml_path>
python main.py export-xml-with-quotes-and-retail-ask <xml_path> [d10b_path]
```

### Reference documentation
- `gilt_knowledge.md` — UK gilt concepts, tax treatment, pricing mechanics, tool limitations
- `README.md` — installation (Windows + Linux), quick start, CLI reference, workbook guide

---

## Known Limitations

| Limitation | Detail |
|---|---|
| D1A cannot be automated | DMO uses ShieldSquare bot-detection — always save manually from browser |
| Approximate cash calculations | No accrued interest, dirty price, exact coupon schedule, or exact IRR |
| Retail ask yield is static | Derived from D10B at export time; does not update if overrides change |
| Day count in bisection | Retail ask yield uses days/365.25, not exact Actual/Actual ICMA |
| Summary sheets are static snapshots | Computed using default nominal amount; per-gilt Inputs overrides not reflected |
| No true ask-side feed | DividendData provides mid/close yield, not live dealer ask yield |

---

## Tax Modeling

Correctly modeled:
- Coupon income taxed at 20%, 40%, 45% scenario rates
- Conventional gilt capital gains CGT-exempt (column P always =0)

Not modeled:
- Personal savings allowance (£500 higher rate, £1,000 basic rate)
- Starting rate for savings (0% on up to £5,000)
- ISA / SIPP wrapper effects
- Wider taxable income interactions

Results are scenarios for comparison, not personal tax advice.

---

## Resolved Issues (historical)

| Issue | Resolution |
|---|---|
| D10B parser used positional column indices | Dynamic header detection by normalised column name |
| Bisection had no plausibility guard | Raises `ValueError` if result outside −5% to +20% |
| Real-file test failed without data file | `pytest.mark.skipif(not path.exists(), ...)` guard added |
| `365.25` day count undocumented | Docstring note added |
| Retail ask yield column gave no hint it was static | Excel comment added to header cell |
| `GiltRetailQuote` missing `coupon_rate` | Field added, derived from gilt name using same fraction-parsing logic |
| Summary snapshots used `datetime.date.today()` | Fixed to use `row.valuation_date` (D1A close-of-business date) |
| Python yearfrac didn't match Excel | Implemented exact Excel `YEARFRAC(basis=1)` algorithm |
| Summary — After-tax Return was redundant | Removed — Annual Net columns in Analysis serve the same purpose |
| Summary — Yield Ranking was redundant | Removed — users sort Analysis col F directly |
| Summary — Best Value was confusing | Removed — Analysis sheet with sort/filter is cleaner |
| CGT column P was always =0 | Removed — conventional gilt gains are always CGT-exempt |
| Added DMO retail sale prices | Cols H (clean) and I (dirty) added from D10B; N/A when D10B not loaded |
