# Gilt Tax Analyzer — Operation Manual

Python tooling that fetches UK conventional gilt data from official sources, enriches it with live market prices, and produces an Excel workbook for after-tax scenario analysis.

---

## Contents

1. [What it does](#what-it-does)
2. [Installation](#installation)
3. [Quick start](#quick-start)
4. [Data sources](#data-sources)
5. [CLI commands](#cli-commands)
6. [Workbook guide](#workbook-guide)
7. [Tax scenario assumptions](#tax-scenario-assumptions)
8. [Known limitations](#known-limitations)
9. [Running tests](#running-tests)

---

## What it does

The tool produces an Excel workbook containing:

- a clean copy of imported market data (prices, yields, maturities)
- a user-editable inputs sheet (nominal amount, price and yield overrides)
- three after-tax scenario columns (20%, 40%, 45% income tax on coupon)
- approximate net cash gain to maturity under each scenario
- an approximate retail ask yield derived from DMO retail sale dirty prices

Conventional UK gilt capital gains are modeled as **CGT-exempt**. Coupon income is taxed at the scenario rate. The workbook does **not** model personal allowances, the starting rate for savings, or any investor-specific tax position.

---

## Installation

Requires Python 3.12 or later.

```powershell
# Create and activate a virtual environment
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt
```

Verify:

```powershell
python main.py info
```

Expected output:

```
Gilt Tax Analyzer
Configured tax scenarios:
- Basic rate: 20%
- Higher rate: 40%
- Additional rate: 45%
```

---

## Quick start

The standard workflow requires two files downloaded from the DMO website each day.

### Step 1 — Download the DMO D1A XML

Go to the DMO gilt data page and save the D1A report as a `.xml` file:

```
https://www.dmo.gov.uk/data/XmlDataReport?reportCode=D1A
```

Save it as:

```
data\d1a.xml
```

> The DMO endpoint blocks automated scripts with an anti-bot page. Always save this file manually from a browser.

### Step 2 — Download the DMO D10B XLS

Go to the DMO Purchase and Sale Service prices page:

```
https://www.dmo.gov.uk/data/gilt-market/purchase-and-sale-service/
```

Download the daily XLS file (labelled with today's date) and save it to the `data\` folder, keeping the original filename, for example:

```
data\20260516 - DMO Gilt Purchase and Sale Service Prices.xls
```

> If today's D10B file shows "No information to display" (e.g. on a bank holiday), use the most recent available file instead.

### Step 3 — Generate the workbook

```powershell
python main.py export-xml-with-quotes-and-retail-ask .\data\d1a.xml ".\data\20260516 - DMO Gilt Purchase and Sale Service Prices.xls" --output .\output\gilt_analysis.xlsx
```

Replace the D10B filename with the actual file you downloaded. Open `output\gilt_analysis.xlsx` in Excel.

---

## Data sources

The tool combines three sources:

| Source | What it provides | URL | How to obtain |
|---|---|---|---|
| DMO D1A XML | ISIN, gilt name, coupon, maturity, instrument type | `https://www.dmo.gov.uk/data/XmlDataReport?reportCode=D1A` | Save from browser — automated fetch is blocked |
| DividendData | Clean price, yield to maturity | `https://www.dividenddata.co.uk/uk-gilts-prices-yields.py` | Fetched live automatically |
| DMO D10B XLS | Retail purchase/sale dirty prices | `https://www.dmo.gov.uk/data/gilt-market/purchase-and-sale-service/` | Download daily file manually |

### Source architecture

```
DMO D1A XML (saved locally)
    -> canonical gilt identity (ISIN, coupon, maturity)

DividendData (live fetch)
    -> clean price + yield-to-maturity enrichment

DMO D10B XLS (saved locally)
    -> retail sale dirty price
    -> approximate retail ask yield derivation

Merged rows
    -> Excel workbook output
```

### Note on the D10B file

The D10B file is date-stamped. If the DMO did not publish prices for a given day (bank holidays or non-business days), the file will contain "No information to display" and the retail ask yield column will be blank. In that case use the most recently published D10B file.

---

## CLI commands

### `info`

Prints the configured tax scenarios. Useful to confirm the environment is working.

```powershell
python main.py info
```

---

### `export`

Attempts a **live** fetch of the DMO D1A XML, then generates a workbook with no price data (DMO D1A does not include prices). Unreliable due to anti-bot blocking — prefer `export-xml-with-quotes`.

```powershell
python main.py export [--output OUTPUT] [--nominal-amount AMOUNT]
```

| Option | Default | Description |
|---|---|---|
| `--output` | `output/gilt_analysis.xlsx` | Path for the output workbook |
| `--nominal-amount` | `10000` | Default holding size in £ written to the Inputs sheet |

---

### `export-xml`

Parses a **locally saved** D1A XML file. Produces a workbook with gilt identity data but no prices or yields. Useful for testing the workbook structure.

```powershell
python main.py export-xml .\data\d1a.xml [--output OUTPUT] [--nominal-amount AMOUNT]
```

| Argument | Description |
|---|---|
| `XML_PATH` | Path to the locally saved D1A XML file |

| Option | Default | Description |
|---|---|---|
| `--output` | `output/gilt_analysis.xlsx` | Path for the output workbook |
| `--nominal-amount` | `10000` | Default holding size in £ |

---

### `export-xml-with-quotes`

Parses a locally saved D1A XML file, fetches live prices and yields from DividendData, and produces a priced workbook. The `Approx Retail Ask Yield %` column will be blank. Use `export-xml-with-quotes-and-retail-ask` for the full workflow.

```powershell
python main.py export-xml-with-quotes .\data\d1a.xml [--output OUTPUT] [--nominal-amount AMOUNT]
```

| Argument | Description |
|---|---|
| `XML_PATH` | Path to the locally saved D1A XML file |

| Option | Default | Description |
|---|---|---|
| `--output` | `output/gilt_analysis.xlsx` | Path for the output workbook |
| `--nominal-amount` | `10000` | Default holding size in £ |

---

### `export-xml-with-quotes-and-retail-ask`

**Standard workflow.** Parses a locally saved D1A XML file, fetches live prices and yields from DividendData, parses a locally saved D10B XLS file, and produces a fully populated workbook including the **Approx Retail Ask Yield %** column.

```powershell
python main.py export-xml-with-quotes-and-retail-ask .\data\d1a.xml ".\data\YYYYMMDD - DMO Gilt Purchase and Sale Service Prices.xls" [--output OUTPUT] [--nominal-amount AMOUNT]
```

| Argument | Description |
|---|---|
| `XML_PATH` | Path to the locally saved D1A XML file |
| `D10B_PATH` | Path to the locally saved D10B XLS file (use the actual dated filename) |

| Option | Default | Description |
|---|---|---|
| `--output` | `output/gilt_analysis.xlsx` | Path for the output workbook |
| `--nominal-amount` | `10000` | Default holding size in £ |

Example:

```powershell
python main.py export-xml-with-quotes-and-retail-ask .\data\d1a.xml ".\data\20260516 - DMO Gilt Purchase and Sale Service Prices.xls" --output .\output\gilt_analysis.xlsx
```

---

## Workbook guide

### Sheet: Market Data

Read-only. Contains all values imported from external sources.

| Column | Content | Source |
|---|---|---|
| ISIN | Security identifier | DMO D1A |
| Gilt Name | Instrument name | DMO D1A |
| Maturity | Redemption date | DMO D1A |
| Coupon % | Annual coupon in percentage points | DMO D1A |
| Imported Price | Clean price | DividendData |
| Imported Yield % | Yield to maturity | DividendData |
| Valuation Date | Date of the D1A export | DMO D1A |
| Source | Source identifier string | Internal |

Do not edit this sheet. Re-run the CLI to refresh it.

---

### Sheet: Inputs

**User-editable.** Adjust values here to customise the analysis.

| Column | Content | How to use |
|---|---|---|
| ISIN | Security identifier | Do not edit — used as the lookup key |
| Nominal Amount (£) | Holding size in £ | Enter your intended holding size per gilt |
| Override Price | Manual price | Leave blank to use the imported price; enter a value to override |
| Override Yield % | Manual yield | Leave blank to use the imported yield; enter a value to override |

The Analysis sheet uses the override value when non-blank, otherwise falls back to the imported value.

---

### Sheet: Tax Scenarios

**User-editable.** Contains the three coupon tax rates used in analysis.

| Column | Content |
|---|---|
| Scenario | Scenario label (Basic rate / Higher rate / Additional rate) |
| Coupon Tax Rate | Rate as a decimal (e.g. 0.20) |
| Notes | Disclaimer text |

You can add rows for custom rates, but the Analysis sheet formulas are currently hard-coded to columns K, L, M (20%, 40%, 45%). Additional custom rates would require formula adjustments.

---

### Sheet: Analysis

Formula-driven outputs. **Do not edit formula cells.** The only data entered here by Python is the ISIN, Gilt Name, Maturity, Coupon %, and (if available) the Approx Retail Ask Yield %.

| Column | Content | Notes |
|---|---|---|
| A | ISIN | Static |
| B | Gilt Name | Static |
| C | Maturity | Static |
| D | Coupon % | Static |
| E | Effective Price | Formula: override if present, else imported |
| F | Effective Yield % | Formula: override if present, else imported |
| G | Approx Retail Ask Yield % | Static import from D10B — does **not** update when overrides change |
| H | Nominal Amount (£) | Formula: looked up from Inputs sheet |
| I | Years to Maturity | Formula: YEARFRAC from valuation date |
| J | Annual Coupon Cash (£) | Formula: Nominal × Coupon / 100 |
| K | Capital Uplift to Par (£) | Formula: Nominal × (100 − Price) / 100 |
| L | Approx Gross Cash Gain to Maturity (£) | Formula: (Annual coupon cash × Years) + Capital uplift |
| M | Coupon Tax @20% (£) | Formula: annual coupon cash × years × 20% |
| N | Coupon Tax @40% (£) | Formula: annual coupon cash × years × 40% |
| O | Coupon Tax @45% (£) | Formula: annual coupon cash × years × 45% |
| P | CGT on Capital Gain (£) | Always zero — conventional gilt gains are CGT-exempt |
| Q | Approx Net Cash Gain @20% (£) | L − M − P |
| R | Approx Net Cash Gain @40% (£) | L − N − P |
| S | Approx Net Cash Gain @45% (£) | L − O − P |

All cash calculations are **approximate**. They do not model accrued interest, dirty price, exact coupon schedules, or reinvestment. They are useful for comparing gilts relative to each other, not for exact return projection.

---

### Sheet: Summary

Currently a placeholder. A ranking dashboard is planned for a future phase.

---

## Tax scenario assumptions

- Conventional UK gilt capital gains are **modeled as CGT-exempt** in all scenarios. Column P is always zero.
- Coupon income is taxed at the scenario rate applied to the total projected coupon cash over the holding period.
- The model does **not** account for:
  - the personal savings allowance (£500 for higher-rate taxpayers, £1,000 for basic-rate)
  - the starting rate for savings (0% on up to £5,000 for low earners)
  - ISA or SIPP wrapper effects
  - any other element of the investor's wider tax position
- Results are **scenarios for comparison**, not personal tax advice.

---

## Known limitations

| Limitation | Detail |
|---|---|
| Approximate cash calculations | No accrued interest, dirty price, exact coupon schedule, or exact IRR |
| Approx Retail Ask Yield % is static | Derived from D10B at export time; does not update if you change price overrides in Excel |
| Day count approximation | Retail ask yield uses days/365.25, not the exact Actual/Actual ICMA convention |
| Live DMO fetch unreliable | Anti-bot protection blocks automated requests — always use the local XML workflow |
| Summary sheet empty | Rankings and maturity-bucket views are not yet implemented |
| No true ask-side feed | DividendData provides a mid/close yield, not a live dealer ask yield |

---

## Running tests

```powershell
.\.venv\Scripts\python.exe -m pytest tests\ -v
```

Expected: **23 passed**

The test `test_parse_d10b_xls_real_file` is automatically skipped if the real D10B file is not present at `data/20260515 - DMO Gilt Purchase and Sale Service Prices.xls`. All other tests are self-contained and run without any external files.
