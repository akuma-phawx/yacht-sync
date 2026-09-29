"""Write matched payments to a CSV using the master task columns."""

from __future__ import annotations

import csv
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from services.matcher import ExtractedInvoice, MatchedPayment, format_date, format_money

CSV_HEADERS = (
    "Tasks",
    "Supplier",
    "Start Date",
    "Delivery Date",
    "Owner's Approval",
    "Status",
    "Estimated Cost",
    "Net Price",
    "VAT 24%",
    "Final Cost",
    "Notes",
    "Invoice",
)


def to_csv_fields(boat_name: str, matched: MatchedPayment) -> dict[str, str]:
    payment = matched.payment
    notes = (
        f"{boat_name} | {payment.boat_category} / {payment.boat_subcategory} | "
        f"{payment.invoice_number}"
    )
    if not matched.matched:
        notes = f"{notes} | UNMATCHED"
        net_price = ""
        vat = ""
        final_cost = format_money(payment.amount_original)
        invoice_cell = ""
    else:
        net_price = _sum_text(invoice.net_amount for invoice in matched.invoices)
        vat = _sum_text(_resolved_vat(invoice) for invoice in matched.invoices)
        final_cost = format_money(
            sum((invoice.total_amount for invoice in matched.invoices), Decimal("0"))
        )
        invoice_cell = "; ".join(invoice.source_name for invoice in matched.invoices)
    return {
        "Tasks": payment.description,
        "Supplier": payment.supplier_name,
        "Start Date": format_date(payment.create_date),
        "Delivery Date": format_date(payment.invoice_date),
        "Owner's Approval": "",
        "Status": payment.work_order_status,
        "Estimated Cost": format_money(payment.amount_original),
        "Net Price": net_price,
        "VAT 24%": vat,
        "Final Cost": final_cost,
        "Notes": notes,
        "Invoice": invoice_cell,
    }


def write_csv(path: Path, boat_name: str, matches: list[MatchedPayment]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(CSV_HEADERS))
        writer.writeheader()
        for matched in matches:
            writer.writerow(to_csv_fields(boat_name, matched))


def _resolved_vat(invoice: ExtractedInvoice) -> Decimal | None:
    if invoice.vat_amount is not None:
        return invoice.vat_amount
    if invoice.net_amount is not None:
        difference = invoice.total_amount - invoice.net_amount
        return difference.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return None


def _sum_text(amounts: object) -> str:
    present: list[Decimal] = []
    for amount in amounts:
        if amount is None:
            return ""
        present.append(amount)
    if not present:
        return ""
    return format_money(sum(present, Decimal("0")))
