# Gilt Tax Analyzer — Implementation Plan v2

## 1. Project Overview

Build a Python-based UK gilt analysis tool that:

- Retrieves and normalizes UK gilt market data
- Produces professional Excel workbooks
- Keeps Excel as the editable calculation surface
- Supports manual market-data overrides without destroying imported values
- Models after-tax gilt outcomes across user-defined tax scenarios
- Makes the distinctive tax treatment of conventional UK gilts clear to the user

The product should be framed as an **after-tax scenario analyzer**, not as a full personal tax calculator.

---

## 2. Product Principles

1. **Financial correctness before feature breadth**
   - Do not label a rough cash calculation as "total return" if it is not time-adjusted.
   - Distinguish clearly between coupon cash flow, capital uplift to par, gross yield, and after-tax yield.

2. **Excel remains the calculation engine**
   - Python imports, validates, and lays out the workbook.
   - Excel formulas stay visible and editable for auditability and what-if work.

3. **Imported data and user edits must not fight each other**
   - Keep source data, manual overrides, and calculated outputs separate.

4. **Tax outputs are scenarios**
   - Conventional gilt capital gains are generally CGT-exempt.
   - Coupon interest is taxable as savings income, but actual liability depends on the investor's wider tax position and available allowances.

---

## 3. MVP Scope

### Include

- Gilt data import from a defined initial source
- Data normalization and validation
- Excel workbook generation
- Manual override workflow
- Scenario-based tax analysis for:
  - 20%
  - 40%
  - 45%
- Gross and after-tax outcome calculations
- Workbook formatting and summary views

### Exclude Initially

- MongoDB or other persistence layer
- Historical analytics
- Scheduling / automation
- QuantLib
- Full personal tax return logic
- Index-linked gilts
- Strips

---

## 4. Recommended Technology Stack

| Component | Technology |
|---|---|
| Language | Python 3.12 |
| HTTP/API | requests |
| HTML Parsing | BeautifulSoup4, only if the chosen source requires it |
| Data Processing | pandas |
| Excel Generation | openpyxl |
| Configuration | python-dotenv |
| CLI | typer |
| Testing | pytest |

Optional later:

- MongoDB or SQLite for snapshots / history
- QuantLib for more advanced bond math if needed

---

## 5. Suggested Project Structure

```text
gilt-tax-analyzer/
|
|-- app/
|   |-- collectors/
|   |   |-- base.py
|   |   |-- dmo.py
|   |   `-- tradeweb.py
|   |
|   |-- transform.py
|   |-- excel_builder.py
|   |-- formulas.py
|   |-- tax_scenarios.py
|   |-- config.py
|   `-- models.py
|
|-- templates/
|   `-- gilt_template.xlsx
|
|-- output/
|-- tests/
|   |-- test_transform.py
|   |-- test_formulas.py
|   `-- test_excel_builder.py
|
|-- .env.example
|-- requirements.txt
|-- README.md
`-- main.py
```

---

## 6. Data Source Strategy

### Source Requirements

The first source should provide, at minimum:

- ISIN
- Gilt name
- Coupon
- Maturity date
- Clean price or reference price
- Yield, if available
- Data timestamp / valuation date

### Recommended Approach

1. **Primary MVP source:** choose a source that is actually accessible and reproducible in the implementation environment.
2. **Likely baseline candidate:** UK Debt Management Office data, because it is official and suitable for reference datasets.
3. **Possible later enhancement:** Tradeweb / Tradeweb FTSE gilt closing prices, if access and licensing are acceptable.

### Important Constraint

Do not assume a public marketing page is a machine-consumable data feed. Source selection should be finalized only after a short feasibility spike confirms:

- availability
- licensing
- update frequency
- field coverage
- stable retrieval method

---

## 7. Data Model

### Core Gilt Fields

| Field | Description |
|---|---|
| isin | Security identifier |
| gilt_name | Display name |
| coupon_rate | Annual coupon percentage |
| maturity_date | Redemption date |
| imported_price | Imported clean/reference price |
| imported_yield | Imported yield to maturity, if provided |
| valuation_date | Date of imported market data |
| source_name | Origin of the data |

### User Override Fields

| Field | Description |
|---|---|
| override_price | Optional manually entered price |
| override_yield | Optional manually entered yield |
| nominal_amount | User-selected holding size |

### Derived Fields

| Field | Description |
|---|---|
| effective_price | Override price if present, otherwise imported price |
| effective_yield | Override yield if present, otherwise imported yield |
| years_to_maturity | Time from valuation date to maturity |
| annual_coupon_cash | Coupon cash amount per year |
| capital_uplift_to_par | Redemption gain from purchase price to par |
| coupon_tax | Scenario-dependent tax on coupon income |
| net_cash_gain_to_maturity | Capital uplift plus coupon cash less coupon tax |
| approximate_after_tax_yield | Scenario-based approximation for ranking / comparison |

---

## 8. Excel Workbook Design

### Worksheet 1: `Market Data`

Raw imported values only.

Suggested columns:

| Column | Header |
|---|---|
| A | ISIN |
| B | Gilt Name |
| C | Maturity |
| D | Coupon % |
| E | Imported Price |
| F | Imported Yield % |
| G | Valuation Date |
| H | Source |

### Worksheet 2: `Inputs`

User-editable assumptions and overrides.

Suggested columns:

| Column | Header |
|---|---|
| A | ISIN |
| B | Nominal Amount (£) |
| C | Override Price |
| D | Override Yield % |

### Worksheet 3: `Tax Scenarios`

Editable scenario assumptions.

| Column | Header |
|---|---|
| A | Scenario |
| B | Coupon Tax Rate |
| C | Notes |

Default rows:

- Basic rate — 20%
- Higher rate — 40%
- Additional rate — 45%

Note in workbook:

> Conventional UK gilt capital gains are treated as CGT-exempt in this model. Coupon-tax outputs are scenarios and do not attempt to model the investor's full personal allowances or wider tax position.

### Worksheet 4: `Analysis`

Calculated outputs only.

Suggested columns:

| Column | Header |
|---|---|
| A | ISIN |
| B | Gilt Name |
| C | Maturity |
| D | Coupon % |
| E | Effective Price |
| F | Nominal Amount (£) |
| G | Years to Maturity |
| H | Annual Coupon Cash (£) |
| I | Capital Uplift to Par (£) |
| J | Approx Gross Cash Gain to Maturity (£) |
| K | Coupon Tax @20% (£) |
| L | Coupon Tax @40% (£) |
| M | Coupon Tax @45% (£) |
| N | CGT on Conventional Gilt Capital Gain (£) |
| O | Approx Net Cash Gain @20% (£) |
| P | Approx Net Cash Gain @40% (£) |
| Q | Approx Net Cash Gain @45% (£) |

### Worksheet 5: `Summary`

Ranked, presentation-grade view for decision-making.

Possible sections:

- Best low-coupon gilts by approximate after-tax outcome
- Highest capital-uplift opportunities
- Short / medium / long maturity buckets
- Filters by maturity date, coupon, and effective yield

---

## 9. Excel Formula Design

Assume row 2 of `Analysis`.

### Effective Price

```excel
=IFERROR(XLOOKUP(A2,Inputs!A:A,Inputs!C:C),XLOOKUP(A2,'Market Data'!A:A,'Market Data'!E:E))
```

If compatibility with older Excel versions is required, use `INDEX` / `MATCH` instead of `XLOOKUP`.

### Nominal Amount

```excel
=XLOOKUP(A2,Inputs!A:A,Inputs!B:B)
```

### Years to Maturity

```excel
=YEARFRAC(XLOOKUP(A2,'Market Data'!A:A,'Market Data'!G:G),C2,1)
```

### Annual Coupon Cash

```excel
=F2*D2/100
```

### Capital Uplift to Par

```excel
=F2*(100-E2)/100
```

### Approx Gross Cash Gain to Maturity

```excel
=H2*G2+I2
```

### Coupon Tax @20%

```excel
=H2*G2*20%
```

### Coupon Tax @40%

```excel
=H2*G2*40%
```

### Coupon Tax @45%

```excel
=H2*G2*45%
```

### CGT on Conventional Gilt Capital Gain

```excel
=0
```

### Approx Net Cash Gain @20%

```excel
=J2-K2-N2
```

### Approx Net Cash Gain @40%

```excel
=J2-L2-N2
```

### Approx Net Cash Gain @45%

```excel
=J2-M2-N2
```

---

## 10. Calculation Notes and Naming Discipline

### Use "Approx" Where Appropriate

The workbook should use labels such as:

- `Approx Gross Cash Gain to Maturity`
- `Approx Net Cash Gain`

These calculations are useful for comparing gilts, but they are not the same as:

- full total return
- exact yield to maturity
- exact after-tax IRR

### Why

An exact bond return model may require:

- clean vs dirty price handling
- accrued interest
- actual coupon schedule
- exact settlement date
- compounding convention
- day-count convention

Those can be added later, but should not be silently implied in v1.

---

## 11. Formatting and Usability Standards

- Freeze panes
- Auto filter on data tables
- Bold, finance-style headers
- Currency format: `£#,##0.00`
- Percentage format: `0.000%`
- Date format: `dd-mmm-yyyy`
- Conditional formatting for:
  - low coupon gilts
  - positive capital uplift
  - highest approximate net gain by scenario
- Workbook note / disclaimer sheet or banner
- Protect formula cells if practical, while leaving input cells editable

---

## 12. Testing Strategy

### Unit Tests

- Parse and normalize source rows
- Handle missing / malformed prices and yields
- Validate maturity-date parsing
- Generate expected formulas
- Confirm override logic

### Workbook Tests

- Required sheets exist
- Required headers exist
- Formula cells are populated correctly
- Input cells remain editable
- Formatting rules are applied

### Financial Logic Tests

- Capital gain on conventional gilts resolves to zero CGT in modeled scenarios
- Coupon tax differs correctly by scenario
- Approx cash gain scales correctly with:
  - price discount / premium
  - coupon rate
  - nominal amount
  - years to maturity

---

## 13. Delivery Phases

### Phase 0 — Feasibility Spike

- Confirm usable data source
- Capture sample rows
- Finalize schema
- Decide whether imported yield is trustworthy enough for display in v1

### Phase 1 — Core Workbook MVP

- Data import
- Normalization
- Workbook generation
- Manual override support
- Basic scenario calculations
- Formatting

### Phase 2 — Better Analytics

- Approximate after-tax yield ranking
- Summary dashboard
- More robust validation
- Export metadata and timestamps

### Phase 3 — Exact Bond Math

- Dirty-price support
- Accrued interest
- Exact coupon schedule
- Exact YTM / IRR
- Potential QuantLib adoption

---

## 14. Recommended First Build Order

1. Finalize the data source
2. Define the normalized gilt schema
3. Build workbook skeleton with all sheets
4. Implement formulas and override behavior
5. Add formatting
6. Add validation and tests
7. Add summary ranking

---

## 15. Final Recommendation

Keep the v1 product narrow, legible, and honest:

- Python should fetch, normalize, validate, and assemble.
- Excel should calculate, expose assumptions, and invite what-if analysis.
- The product should emphasize the core gilt insight:
  - coupon income is taxable
  - conventional gilt capital gains are modeled as CGT-exempt
  - low-coupon discounted gilts can therefore behave very differently after tax from higher-coupon alternatives

That is the vein worth mining first.
