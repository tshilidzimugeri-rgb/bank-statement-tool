"""Generates synthetic bank-statement-shaped PDFs for smoke-testing the
extraction pipeline before real client statements are available.

These are NOT real statements - just plausible layouts (FNB-style split
debit/credit columns, Capitec-style signed amount column, and a
low-text/near-blank page to exercise the OCR fallback path) so the
text-extraction and normalization logic can be verified end to end.
Run: .venv/Scripts/python.exe tests/fixtures/generate_fixtures.py
"""
from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

OUT_DIR = Path(__file__).parent
styles = getSampleStyleSheet()


def build_fnb_statement() -> None:
    path = OUT_DIR / "sample_fnb_statement.pdf"
    doc = SimpleDocTemplate(str(path), pagesize=A4)
    elements = [
        Paragraph("First National Bank", styles["Title"]),
        Paragraph("Account Holder: Acme Trading CC", styles["Normal"]),
        Paragraph("Account Number: 6205 1234 567", styles["Normal"]),
        Paragraph("Statement Period: 01 Jan 2024 to 31 Jan 2024", styles["Normal"]),
        Spacer(1, 10 * mm),
    ]

    data = [["Date", "Description", "Debit", "Credit", "Balance"]]
    rows = [
        ("01/01/2024", "Opening Balance", "", "", "10,000.00"),
        ("03/01/2024", "POS Purchase Woolworths", "450.75", "", "9,549.25"),
        ("05/01/2024", "Salary Payment", "", "25,000.00", "34,549.25"),
        ("08/01/2024", "Debit Order Insurance", "1,200.00", "", "33,349.25"),
        ("12/01/2024", "ATM Withdrawal", "2,000.00", "", "31,349.25"),
        ("15/01/2024", "EFT Received - Client Invoice 1042", "", "8,750.00", "40,099.25"),
        ("20/01/2024", "Bank Charges", "89.50", "", "40,009.75"),
        ("25/01/2024", "Transfer to Savings", "5,000.00", "", "35,009.75"),
        ("31/01/2024", "Closing Balance", "", "", "35,009.75"),
    ]
    data.extend(list(r) for r in rows)

    table = Table(data, colWidths=[22 * mm, 65 * mm, 25 * mm, 25 * mm, 28 * mm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003366")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
            ]
        )
    )
    elements.append(table)
    doc.build(elements)
    print(f"wrote {path}")


def build_capitec_statement() -> None:
    path = OUT_DIR / "sample_capitec_statement.pdf"
    doc = SimpleDocTemplate(str(path), pagesize=A4)
    elements = [
        Paragraph("Capitec Bank", styles["Title"]),
        Paragraph("Client: Jane Dlamini", styles["Normal"]),
        Paragraph("Date Range: 01 Feb 2024 to 29 Feb 2024", styles["Normal"]),
        Spacer(1, 10 * mm),
    ]

    data = [["Date", "Description", "Amount", "Balance"]]
    rows = [
        ("01/02/2024", "Opening Balance", "", "5,000.00"),
        ("02/02/2024", "Grocery Store Purchase", "-620.40", "4,379.60"),
        ("04/02/2024", "Salary", "18,500.00", "22,879.60"),
        ("10/02/2024", "Airtime Purchase", "-150.00", "22,729.60"),
        ("18/02/2024", "Client Payment Received", "6,200.00", "28,929.60"),
        ("28/02/2024", "Monthly Fee", "-99.00", "28,830.60"),
    ]
    data.extend(list(r) for r in rows)

    table = Table(data, colWidths=[24 * mm, 75 * mm, 28 * mm, 28 * mm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E4002B")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
            ]
        )
    )
    elements.append(table)
    doc.build(elements)
    print(f"wrote {path}")


def build_unmapped_client_statement() -> None:
    """Standard Bank layout with no matching entry in clients.yaml, to
    exercise the interactive/UNMAPPED_CLIENT path."""
    path = OUT_DIR / "sample_standardbank_statement_unmapped.pdf"
    doc = SimpleDocTemplate(str(path), pagesize=A4)
    elements = [
        Paragraph("Standard Bank", styles["Title"]),
        Paragraph("Account Holder: Zolani Manufacturing (Pty) Ltd", styles["Normal"]),
        Paragraph("Statement Period: 01 Mar 2024 to 31 Mar 2024", styles["Normal"]),
        Spacer(1, 10 * mm),
    ]
    data = [["Date", "Description", "Debit", "Credit", "Balance"]]
    rows = [
        ("01/03/2024", "Opening Balance", "", "", "12,300.00"),
        ("06/03/2024", "Supplier Payment", "3,400.00", "", "8,900.00"),
        ("14/03/2024", "Customer EFT", "", "9,600.00", "18,500.00"),
    ]
    data.extend(list(r) for r in rows)
    table = Table(data, colWidths=[22 * mm, 65 * mm, 25 * mm, 25 * mm, 28 * mm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0033A0")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
            ]
        )
    )
    elements.append(table)
    doc.build(elements)
    print(f"wrote {path}")


if __name__ == "__main__":
    build_fnb_statement()
    build_capitec_statement()
    build_unmapped_client_statement()
