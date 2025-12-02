"""
Travel invoice cross-charge extractor.

Reads PDF invoices from the input folder, extracts key fields using regexes,
and writes a consolidated Excel file.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from datetime import datetime, date
from pathlib import Path
from typing import Iterable, List, Optional

import pandas as pd
import pdfplumber


INPUT_DIR_PRIMARY = Path("02-inputs/Cross charge list")
INPUT_DIR_FALLBACK = Path("02-inputs/invoices")
OUTPUT_PATH = Path("03-outputs/cross charge list/travel_cross_charge.xlsx")


@dataclass
class InvoiceRecord:
    """Structured representation of an extracted invoice."""

    passenger_name: Optional[str]
    invoice_number: Optional[str]
    invoice_date: Optional[date]
    invoice_amount_gross: Optional[float]
    gst_amount: Optional[float]
    net_amount: Optional[float]
    source_file: str


def setup_logging() -> None:
    """Configure basic logging."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )


def clean_amount(value: str) -> Optional[float]:
    """Convert a numeric string with optional commas to float."""
    if value is None:
        return None
    cleaned = value.replace(",", "").strip()
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        logging.warning("Could not parse amount from %s", value)
        return None


def extract_invoice_number(text: str) -> Optional[str]:
    """Extract invoice number from 'Tax Invoice - <number>'."""
    match = re.search(r"Tax\s+Invoice\s*-\s*([A-Za-z0-9.\-]+)", text, flags=re.IGNORECASE)
    return match.group(1).strip() if match else None


def extract_invoice_date(text: str) -> Optional[date]:
    """Extract issue date and return as date object."""
    match = re.search(r"Issue\s+Date\s+(\d{1,2}/\d{1,2}/\d{4})", text, flags=re.IGNORECASE)
    if not match:
        return None
    try:
        return datetime.strptime(match.group(1), "%d/%m/%Y").date()
    except ValueError:
        logging.warning("Could not parse invoice date from %s", match.group(1))
        return None


def extract_passenger_name(text: str) -> Optional[str]:
    """Extract passenger name from the line starting with 'Passenger' or 'Passengers'."""
    for line in text.splitlines():
        match = re.match(r"\s*Passengers?:\s*(.+)", line, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return None


def _last_number_in_line(line: str) -> Optional[float]:
    """Return the last numeric value in a line, if any."""
    numbers = re.findall(r"([-+]?\d[\d,]*\.?\d*)", line)
    if not numbers:
        return None
    return clean_amount(numbers[-1])


def extract_invoice_total(text: str) -> Optional[float]:
    """Extract gross invoice total from the 'Invoice Total' line."""
    for line in text.splitlines():
        if re.match(r"\s*Invoice\s+Total", line, flags=re.IGNORECASE):
            amount = _last_number_in_line(line)
            if amount is not None:
                return amount
    return None


def extract_gst_amount(text: str) -> Optional[float]:
    """Extract GST amount from a line that contains GST and a single amount."""
    for line in text.splitlines():
        match = re.match(r"\s*GST\b[^\d\-]*([-+]?\d[\d,]*\.?\d*)\s*$", line, flags=re.IGNORECASE)
        if match:
            return clean_amount(match.group(1))
    for line in text.splitlines():
        if "GST" in line:
            amount = _last_number_in_line(line)
            if amount is not None:
                return amount
    return None


def extract_fields(text: str, source_file: str) -> InvoiceRecord:
    """Extract all required fields from invoice text."""
    invoice_number = extract_invoice_number(text)
    invoice_date = extract_invoice_date(text)
    passenger_name = extract_passenger_name(text)
    gross = extract_invoice_total(text)
    gst = extract_gst_amount(text)
    net = gross - gst if gross is not None and gst is not None else None

    return InvoiceRecord(
        passenger_name=passenger_name,
        invoice_number=invoice_number,
        invoice_date=invoice_date,
        invoice_amount_gross=gross,
        gst_amount=gst,
        net_amount=net,
        source_file=source_file,
    )


def load_text_from_pdf(pdf_path: Path) -> Optional[str]:
    """Extract text from the first page of a PDF."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                logging.warning("No pages found in %s", pdf_path.name)
                return None
            page = pdf.pages[0]
            text_raw = page.extract_text()
            if text_raw is None:
                logging.warning("No text extracted from %s", pdf_path.name)
                return None
            return "\n".join(text_raw.splitlines())
    except Exception as exc:
        logging.error("Failed to read %s: %s", pdf_path.name, exc)
        return None


def find_input_files() -> List[Path]:
    """Return list of PDF files from primary or fallback directory."""
    if INPUT_DIR_PRIMARY.exists():
        files = sorted(INPUT_DIR_PRIMARY.glob("*.pdf"))
        if files:
            logging.info("Using primary input dir: %s", INPUT_DIR_PRIMARY)
            return files
    if INPUT_DIR_FALLBACK.exists():
        files = sorted(INPUT_DIR_FALLBACK.glob("*.pdf"))
        if files:
            logging.info("Using fallback input dir: %s", INPUT_DIR_FALLBACK)
            return files
    logging.warning("No PDF files found in %s or %s", INPUT_DIR_PRIMARY, INPUT_DIR_FALLBACK)
    return []


def records_to_dataframe(records: Iterable[InvoiceRecord]) -> pd.DataFrame:
    """Convert records to DataFrame with correct dtypes."""
    data = [
        {
            "Passenger_Name": r.passenger_name,
            "Invoice_Number": r.invoice_number,
            "Invoice_Date": r.invoice_date,
            "Invoice_Amount_Gross": r.invoice_amount_gross,
            "GST_Amount": r.gst_amount,
            "Net_Amount": r.net_amount,
            "Source_File": r.source_file,
        }
        for r in records
    ]
    df = pd.DataFrame(data)
    if not df.empty:
        df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], errors="coerce")
        float_cols = ["Invoice_Amount_Gross", "GST_Amount", "Net_Amount"]
        for col in float_cols:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def main() -> None:
    """Orchestrate PDF extraction and Excel export."""
    setup_logging()
    pdf_files = find_input_files()
    if not pdf_files:
        logging.warning("No input files to process. Exiting.")
        return

    records: List[InvoiceRecord] = []
    for pdf_path in pdf_files:
        text = load_text_from_pdf(pdf_path)
        if not text:
            logging.warning("Skipping %s due to missing text", pdf_path.name)
            continue
        record = extract_fields(text, pdf_path.name)

        if not all(
            [
                record.invoice_number,
                record.invoice_date,
                record.passenger_name,
                record.invoice_amount_gross is not None,
                record.gst_amount is not None,
                record.net_amount is not None,
            ]
        ):
            logging.warning("Missing fields in %s -> %s", pdf_path.name, record)

        records.append(record)
        logging.info("Processed %s", pdf_path.name)

    df = records_to_dataframe(records)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(OUTPUT_PATH, index=False, sheet_name="Invoices", engine="openpyxl")
    logging.info("Wrote %d records to %s", len(df), OUTPUT_PATH)


if __name__ == "__main__":
    main()
