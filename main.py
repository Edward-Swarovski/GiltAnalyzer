from pathlib import Path
from typing import Optional

import typer

from app.collectors.dmo import DmoCollector, parse_d1a_xml
from app.collectors.dividenddata import DividendDataCollector, enrich_with_quotes
from app.collectors.d10b import approximate_retail_ask_yield, parse_d10b_xls
from app.excel_builder import build_workbook
from app.tax_scenarios import DEFAULT_TAX_SCENARIOS

cli = typer.Typer()

_D10B_GLOB = "*DMO Gilt Purchase and Sale Service Prices*.xls*"


def _find_latest_d10b(data_dir: Path = Path("data")) -> Path:
    """Return the most recently modified D10B XLS file in data_dir, or raise."""
    candidates = sorted(data_dir.glob(_D10B_GLOB), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        raise FileNotFoundError(
            f"No D10B file found in {data_dir}/. "
            "Download the daily XLS from https://www.dmo.gov.uk/data/pdfdatareport?reportCode=D10B "
            "and save it to the data/ folder."
        )
    return candidates[0]


@cli.command()
def info() -> None:
    print("Gilt Tax Analyzer")
    print("Configured tax scenarios:")
    for scenario in DEFAULT_TAX_SCENARIOS:
        print(f"- {scenario.name}: {scenario.coupon_tax_rate:.0%}")


@cli.command()
def export(
    output: Path = Path("output/gilt_analysis.xlsx"),
    nominal_amount: int = 10_000,
) -> None:
    rows = DmoCollector().fetch()
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = build_workbook(rows, default_nominal_amount=nominal_amount)
    workbook.save(output)
    print(f"Wrote {len(rows)} conventional gilts to {output}")


@cli.command("export-xml")
def export_xml(
    xml_path: Path,
    output: Path = Path("output/gilt_analysis.xlsx"),
    nominal_amount: int = 10_000,
) -> None:
    rows = parse_d1a_xml(xml_path.read_text(encoding="utf-8"))
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = build_workbook(rows, default_nominal_amount=nominal_amount)
    workbook.save(output)
    print(f"Wrote {len(rows)} conventional gilts from {xml_path} to {output}")


@cli.command("export-xml-with-quotes")
def export_xml_with_quotes(
    xml_path: Path,
    output: Path = Path("output/gilt_analysis.xlsx"),
    nominal_amount: int = 10_000,
) -> None:
    rows = parse_d1a_xml(xml_path.read_text(encoding="utf-8"))
    quotes = DividendDataCollector().fetch()
    enriched_rows = enrich_with_quotes(rows, quotes)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = build_workbook(
        enriched_rows,
        default_nominal_amount=nominal_amount,
    )
    workbook.save(output)
    priced_count = sum(row.imported_price is not None for row in enriched_rows)
    print(
        f"Wrote {len(enriched_rows)} conventional gilts to {output} "
        f"with {priced_count} matched prices"
    )


@cli.command("export-xml-with-quotes-and-retail-ask")
def export_xml_with_quotes_and_retail_ask(
    xml_path: Path,
    d10b_path: Optional[Path] = typer.Argument(
        default=None,
        help="Path to D10B XLS file. If omitted, the most recent file in data/ is used automatically.",
    ),
    output: Path = Path("output/gilt_analysis.xlsx"),
    nominal_amount: int = 10_000,
) -> None:
    resolved_d10b = d10b_path or _find_latest_d10b()
    if d10b_path is None:
        print(f"Auto-detected D10B file: {resolved_d10b}")
    rows = parse_d1a_xml(xml_path.read_text(encoding="utf-8"))
    enriched_rows = enrich_with_quotes(rows, DividendDataCollector().fetch())
    retail_quotes = {quote.isin: quote for quote in parse_d10b_xls(resolved_d10b)}
    retail_ask_yields = {
        row.isin: float(
            approximate_retail_ask_yield(
                coupon_rate=row.coupon_rate,
                sale_dirty_price=retail_quotes[row.isin].sale_dirty_price,
                redemption_date=row.maturity_date,
                valuation_date=retail_quotes[row.isin].data_date,
            )
        )
        for row in enriched_rows
        if row.isin in retail_quotes
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = build_workbook(
        enriched_rows,
        default_nominal_amount=nominal_amount,
        retail_ask_yields=retail_ask_yields,
    )
    workbook.save(output)
    print(
        f"Wrote {len(enriched_rows)} conventional gilts to {output} "
        f"with {len(retail_ask_yields)} retail ask-yield estimates"
    )


def main() -> None:
    cli()


if __name__ == "__main__":
    main()
