# Gilt Tax Analyzer — Implementation Plan

## Project Overview

The project will build a Python-based UK Gilt analysis tool that:

- Retrieves UK Gilt market data
- Generates Excel reports
- Uses Excel formulas instead of pre-calculated values
- Supports manual market data updates
- Performs UK tax analysis
- Calculates after-tax returns for multiple tax bands

---

# 1. Project Objectives

## Functional Requirements

The application shall:

1. Fetch UK Gilt market data
2. Normalize raw market data
3. Generate Excel workbook output
4. Use Excel formulas for calculations
5. Allow manual editing of market data
6. Support tax analysis for:
   - 20% taxpayer
   - 40% taxpayer
   - 45% taxpayer
7. Export professional finance-grade Excel reports

---

# 2. Recommended Technology Stack

| Component | Technology |
|---|---|
| Language | Python 3.12 |
| HTTP/API | requests |
| HTML Parsing | BeautifulSoup4 |
| Data Processing | pandas |
| Excel Generation | openpyxl |
| Configuration | python-dotenv |
| Optional Database | MongoDB |
| Optional CLI | typer |

---

# 3. Suggested Project Structure

```text
gilt-tax-analyzer/
│
├── app/
│   ├── collector.py
│   ├── transform.py
│   ├── excel_builder.py
│   ├── formulas.py
│   ├── tax_engine.py
│   ├── config.py
│   └── models.py
│
├── templates/
│   └── gilt_template.xlsx
│
├── output/
│
├── tests/
│
├── requirements.txt
├── README.md
└── main.py
```

---

# 4. Data Source Options

## Primary Recommendation

### Tradeweb

https://www.tradeweb.com/our-markets/government-bonds/uk-gilts/

Data fields:

- Bond name
- Coupon
- Yield
- Ask price
- Bid price
- Maturity
- ISIN

---

# 5. Excel Workbook Design

## Worksheet: Gilts

| Column | Header |
|---|---|
| A | Bond |
| B | Maturity |
| C | Coupon % |
| D | Ask Yield % |
| E | Ask Price |
| F | Assumed Nominal (£) |
| G | Annual Coupon (£) |
| H | Capital Gain at Maturity (£) |
| I | Approx Total Return to Maturity (£) |
| J | Income Tax @20% (£) |
| K | Income Tax @40% (£) |
| L | Income Tax @45% (£) |
| M | CGT on Capital Gain |
| N | Net Return @20% Tax (£) |
| O | Net Return @40% Tax (£) |
| P | Net Return @45% Tax (£) |

---

# 6. Excel Formula Design

## Annual Coupon

```excel
=F2*C2/100
```

## Capital Gain

```excel
=F2*(100-E2)/100
```

## Approx Total Return

```excel
=G2+H2
```

## Income Tax @20%

```excel
=G2*20%
```

## Income Tax @40%

```excel
=G2*40%
```

## Income Tax @45%

```excel
=G2*45%
```

## Net Return @20%

```excel
=I2-J2
```

## Net Return @40%

```excel
=I2-K2
```

## Net Return @45%

```excel
=I2-L2
```

---

# 7. Recommended Formatting

- Freeze top row
- Auto filter
- Bold headers
- Currency format: £#,##0.00
- Percentage format: 0.000%

---

# 8. Phase 1 MVP Scope

Include:

- Gilt data collection
- Formula-based Excel workbook
- Tax analysis
- Workbook formatting

Exclude initially:

- MongoDB
- QuantLib
- Historical analytics
- Scheduling

---

# 9. Final Recommendation

Keep Excel as the calculation engine.

Python should:
- fetch data
- normalize data
- build workbook

Excel should:
- calculate formulas
- allow manual edits
- support what-if analysis
