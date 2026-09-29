"""CLI for matching yacht invoices and exporting a CSV."""

from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

from pydantic import ValidationError
from rich.console import Console
from rich.markup import escape
from rich.table import Table

from config import Settings
from services.csv_export import write_csv
from services.matcher import (
    discover_inputs,
    load_payments,
    match_payments,
    resolve_archive_target,
)
from services.pdf_extractor import PdfExtractor, to_extracted_invoice

_SETTING_ENV_NAMES = {
    "gemini_api_key": "GEMINI_API_KEY",
    "gemini_model": "GEMINI_MODEL",
}


def _settings_message(exc: ValidationError) -> str:
    names = []
    for error in exc.errors():
        field = str(error["loc"][0]) if error["loc"] else "settings"
        names.append(_SETTING_ENV_NAMES.get(field, field))
    joined = ", ".join(dict.fromkeys(names))
    return f"Missing or blank in .env: {joined}."


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Match yacht invoices to outstanding payments and export a CSV.",
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=Path("input_files"),
        help="Drop folder with one .xlsx and any number of .pdf invoices",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("output.csv"),
        help="CSV file to write",
    )
    parser.add_argument(
        "--boat",
        help="Boat name used when the workbook filename has no boat-date suffix",
    )
    args = parser.parse_args(argv)
    console = Console()
    try:
        return _run(args.input, args.output, args.boat, console)
    except ValidationError as exc:
        console.print(f"[red]{escape(_settings_message(exc))}[/red]")
        return 1
    except ValueError as exc:
        console.print(f"[red]{escape(str(exc))}[/red]")
        return 1
    except Exception as exc:
        console.print(f"[red]{escape(str(exc))}[/red]")
        console.print_exception()
        return 1


def _run(input_dir: Path, output: Path, boat_override: str | None, console: Console) -> int:
    workbook, pdfs = discover_inputs(input_dir)
    boat_name, _archive_date = resolve_archive_target(workbook.name, boat_override, date.today())
    payments = load_payments(workbook)
    settings = Settings()
    console.print(
        f"Found [bold]{escape(workbook.name)}[/bold] ({len(payments)} payments) and {len(pdfs)} PDF(s)."
    )

    extractor = PdfExtractor.from_settings(settings)
    extracted = []
    for index, path in enumerate(pdfs, start=1):
        console.print(f"Extracting {index}/{len(pdfs)} {escape(path.name)}")
        extraction = extractor.extract(path)
        extracted.append(to_extracted_invoice(path.name, extraction))

    matched = match_payments(payments, extracted)
    write_csv(output, boat_name, matched)
    _print_summary(console, matched, extracted, output)
    return 0


def _print_summary(console: Console, matched, extracted, output: Path) -> None:
    table = Table(title="Payment matches")
    table.add_column("Row", justify="right")
    table.add_column("Supplier")
    table.add_column("Amount", justify="right")
    table.add_column("Status")
    table.add_column("Invoices")
    for item in matched:
        names = ", ".join(invoice.source_name for invoice in item.invoices) or "-"
        table.add_row(
            str(item.payment.row_number),
            escape(item.payment.supplier_name),
            f"{item.payment.amount_original:.2f}",
            "MATCHED" if item.matched else "UNMATCHED",
            escape(names),
        )
    console.print(table)
    used = {invoice.source_name for item in matched for invoice in item.invoices}
    unused = [invoice.source_name for invoice in extracted if invoice.source_name not in used]
    unmatched = sum(1 for item in matched if not item.matched)
    console.print(f"Wrote {len(matched)} row(s) to {escape(str(output))}, {unmatched} unmatched.")
    if unused:
        console.print("Unused PDFs: " + escape(", ".join(unused)))


if __name__ == "__main__":
    raise SystemExit(main())
