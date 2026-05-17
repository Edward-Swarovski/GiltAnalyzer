# Gilt Tax Analyzer — Operation Manual

Python tooling that fetches UK conventional gilt data from official sources, enriches it with live market prices, and produces an Excel workbook for after-tax scenario analysis.

**Repository:** `https://github.com/Edward-Swarovski/GiltAnalyzer.git`

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
- a user-editable Inputs sheet (nominal amount per gilt, price and yield overrides)
- three after-tax scenario columns (20%, 40%, 45% income tax on coupons)
- approximate net cash gain to maturity under each scenario
- an approximate retail ask yield derived from DMO retail sale dirty prices
- Annual Net columns (S/T/U) showing after-tax annual return normalised for duration, shown in blue
- an Instructions sheet and a separate `gilt_knowledge.md` reference file

Conventional UK gilt capital gains are modelled as **CGT-exempt**. Coupon income is taxed at the scenario rate. The workbook does **not** model personal allowances, the starting rate for savings, or any investor-specific tax position.

---

## Installation

Requires Python 3.12 or later and Git.

### Step 1 — Clone the repository

```bash
git clone https://github.com/Edward-Swarovski/GiltAnalyzer.git
cd GiltAnalyzer
```

### Step 2 — Create environment and install dependencies

**Windows (PowerShell):**
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

**Linux / macOS:**
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Step 3 — Verify

```bash
python main.py info
```

Expected output (both platforms):

```
Gilt Tax Analyzer
Configured tax scenarios:
- Basic rate: 20%
- Higher rate: 40%
- Additional rate: 45%
```

> **Linux note:** The output workbook is standard `.xlsx` format and opens in LibreOffice Calc as well as Excel. All formulas use `INDEX`/`MATCH`/`IFERROR` — no Excel-only functions are used.

---

## Quick start

The standard workflow requires two files downloaded from the DMO website each day.

### Step 1 — Download the DMO D1A XML

Go to the DMO gilt data page and save the D1A report as a `.xml` file:

```
https://www.dmo.gov.uk/data/XmlDataReport?reportCode=D1A
```

Save it as `data/d1a.xml`.

> The DMO uses bot-detection (ShieldSquare) that blocks automated script downloads. Always save this file manually from a browser.

### Step 2 — Download the DMO D10B XLS

Go to the DMO Purchase and Sale Service prices page:

```
https://www.dmo.gov.uk/data/pdfdatareport?reportCode=D10B
```

1. The page shows a date picker — today's date should be pre-selected
2. Click the **Excel** button to download the XLS file
3. Save it to the `data/` folder, keeping the original filename, for example:

```
data/20260516 - DMO Gilt Purchase and Sale Service Prices.xls
```

> If today's file shows "No information to display" (e.g. on a bank holiday), pick the most recent business day from the date picker instead.

### Step 3 — Generate the workbook

Activate the virtual environment first, then run the export command.

**Windows (PowerShell):**
```powershell
.\.venv\Scripts\Activate.ps1
python main.py export-xml-with-quotes-and-retail-ask ./data/d1a.xml --output ./output/gilt_analysis.xlsx
```

**Linux / macOS:**
```bash
source .venv/bin/activate
python main.py export-xml-with-quotes-and-retail-ask ./data/d1a.xml --output ./output/gilt_analysis.xlsx
```

The D10B file is auto-detected from `data/` — no need to type the filename. If multiple D10B files are present, the most recently modified one is used. You can pass the path explicitly if needed:

```powershell
python main.py export-xml-with-quotes-and-retail-ask ./data/d1a.xml "./data/20260516 - DMO Gilt Purchase and Sale Service Prices.xls" --output ./output/gilt_analysis.xlsx
```

Open `output/gilt_analysis.xlsx` in Excel or LibreOffice Calc.

---

## Data sources

The tool combines three sources:

| Source | What it provides | URL | How to obtain |
|---|---|---|---|
| DMO D1A XML | ISIN, gilt name, coupon, maturity, instrument type | `https://www.dmo.gov.uk/data/XmlDataReport?reportCode=D1A` | Save from browser — automated fetch is blocked |
| DividendData | Clean price, yield to maturity | `https://www.dividenddata.co.uk/uk-gilts-prices-yields.py` | Fetched live automatically |
| DMO D10B XLS | Retail purchase/sale dirty prices | `https://www.dmo.gov.uk/data/pdfdatareport?reportCode=D10B` | Select date, click Excel, save to data/ |

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

```bash
python main.py info
```

---

### `export`

Attempts a **live** fetch of the DMO D1A XML, then generates a workbook with no price data (DMO D1A does not include prices). Unreliable due to anti-bot blocking — prefer `export-xml-with-quotes`.

```bash
python main.py export [--output OUTPUT] [--nominal-amount AMOUNT]
```

| Option | Default | Description |
|---|---|---|
| `--output` | `output/gilt_analysis.xlsx` | Path for the output workbook |
| `--nominal-amount` | `10000` | Default holding size in £ written to the Inputs sheet |

---

### `export-xml`

Parses a **locally saved** D1A XML file. Produces a workbook with gilt identity data but no prices or yields. Useful for testing the workbook structure.

```bash
python main.py export-xml ./data/d1a.xml [--output OUTPUT] [--nominal-amount AMOUNT]
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

```bash
python main.py export-xml-with-quotes ./data/d1a.xml [--output OUTPUT] [--nominal-amount AMOUNT]
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

```bash
python main.py export-xml-with-quotes-and-retail-ask ./data/d1a.xml [D10B_PATH] [--output OUTPUT] [--nominal-amount AMOUNT]
```

| Argument | Description |
|---|---|
| `XML_PATH` | Path to the locally saved D1A XML file |
| `D10B_PATH` | Optional. Path to the D10B XLS file. If omitted, the most recently modified matching file in `data/` is used automatically. |

| Option | Default | Description |
|---|---|---|
| `--output` | `output/gilt_analysis.xlsx` | Path for the output workbook |
| `--nominal-amount` | `10000` | Default holding size in £ |

Examples:

```bash
# Auto-detect the D10B file (recommended)
python main.py export-xml-with-quotes-and-retail-ask ./data/d1a.xml --output ./output/gilt_analysis.xlsx

# Explicit D10B path
python main.py export-xml-with-quotes-and-retail-ask ./data/d1a.xml "./data/20260516 - DMO Gilt Purchase and Sale Service Prices.xls" --output ./output/gilt_analysis.xlsx
```

---

## Workbook guide

The workbook contains 5 sheets:

| Sheet | Purpose | Editable? |
|---|---|---|
| Settings | Default nominal amount and tax scenario reference | Yes — change DefaultNominalAmount |
| Market Data | Raw imported data | No |
| Inputs | Per-gilt overrides | Yes |
| Analysis | Formula-driven cash flow outputs — sort and filter here | No |
| Instructions | How to use the workbook | No |

---

### Sheet: Settings

Contains global configuration and tax scenario reference data.

| Row | Setting | How to use |
|---|---|---|
| 2 | DefaultNominalAmount | Change this to set the default £ holding size for all gilts |
| 5–7 | Tax scenario rates | Reference only — Analysis sheet uses hardcoded 20/40/45% |

---

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
| A | ISIN | Do not edit — lookup key |
| B | Gilt Name | Do not edit — formula from Market Data |
| C | Nominal Amount (£) | Leave blank to use Settings DefaultNominalAmount; enter a number to override for this gilt |
| D | Override Price | Leave blank to use the imported price; enter a clean price to override |
| E | Override Yield % | Leave blank to use the imported yield; enter a yield % to override |

The Analysis sheet uses the override value when non-blank, otherwise falls back to the imported value.

---

### Sheet: Analysis

Formula-driven outputs. **Do not edit formula cells.**

| Column | Content | Notes |
|---|---|---|
| A | ISIN | Static |
| B | Gilt Name | Static |
| C | Maturity | Static |
| D | Coupon % | Static |
| E | Effective Price | Formula: override if present, else imported |
| F | Effective Yield % | Formula: override if present, else imported |
| G | Approx Retail Ask Yield % | Static — from D10B at export time, does not update with overrides |
| H | Nominal Amount (£) | Formula: Inputs col C, or Settings DefaultNominalAmount if blank |
| I | Years to Maturity | Formula: YEARFRAC from valuation date |
| J | Annual Coupon Cash (£) | Formula: Nominal × Coupon / 100 |
| K | Capital Uplift to Par (£) | Formula: Nominal × (100 − Price) / 100 |
| L | Approx Gross Cash Gain to Maturity (£) | Formula: (J × I) + K |
| M | Coupon Tax @20% (£) | Formula: J × I × 20% |
| N | Coupon Tax @40% (£) | Formula: J × I × 40% |
| O | Coupon Tax @45% (£) | Formula: J × I × 45% |
| P | CGT on Capital Gain (£) | Always zero — conventional gilt gains are CGT-exempt |
| P | Approx Net Cash Gain @20% (£) | L − M |
| Q | Approx Net Cash Gain @40% (£) | L − N |
| R | Approx Net Cash Gain @45% (£) | L − O |
| S | Annual Net @20% (£) | P ÷ I — net gain per year at 20% tax; shown in **blue** |
| T | Annual Net @40% (£) | Q ÷ I — net gain per year at 40% tax; shown in **blue** |
| U | Annual Net @45% (£) | R ÷ I — net gain per year at 45% tax; shown in **blue** |

> CGT column removed — conventional gilt capital gains are always exempt, so it added no information.

**Annual Net columns (S/T/U)** are the primary comparison metric. Dividing by Years to Maturity normalises for duration — a 2-year gilt at £450/year ranks correctly above a 30-year gilt at £400/year. **Sort column T descending for a higher-rate taxpayer ranking.** Use S (20%) or U (45%) for other rates. These columns update live when you change Nominal Amount or Override Price in Inputs.
| T | Annual Net @20% (£) | Q ÷ I — net gain per year at 20% tax; shown in **blue** |
| U | Annual Net @40% (£) | R ÷ I — net gain per year at 40% tax; shown in **blue** |
| V | Annual Net @45% (£) | S ÷ I — net gain per year at 45% tax; shown in **blue** |

---

---

## Tax scenario assumptions

- Conventional UK gilt capital gains are **modelled as CGT-exempt** in all scenarios. Column P is always zero.
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
| No true ask-side feed | DividendData provides a mid/close yield, not a live dealer ask yield |
| Summary sheets are static snapshots | Computed in Python at export time using the default nominal amount; per-gilt overrides in Inputs are not reflected in Summary sheets |

---

## Getting updates

When new changes are published to GitHub:

```bash
cd GiltAnalyzer
git pull
pip install -r requirements.txt
```

The `pip install` step is only needed if `requirements.txt` changed. It is safe to run every time.

---

## Running tests

**Windows:**
```powershell
.\.venv\Scripts\python.exe -m pytest tests\ -v
```

**Linux / macOS:**
```bash
python -m pytest tests/ -v
```

Expected: **28 passed**

The test `test_parse_d10b_xls_real_file` is automatically skipped if the real D10B file is not present at `data/20260515 - DMO Gilt Purchase and Sale Service Prices.xls`. All other tests are self-contained and run without any external files.
