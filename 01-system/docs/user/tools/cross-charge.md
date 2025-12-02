# Travel Cross-Charge Extractor
**Category**: ops  
**Version**: v0.1 (Released: 2025-12-02)

## What it does
- Reads travel invoice PDFs and extracts passenger, invoice number/date, gross, GST, and net amounts into a single Excel file.

## Inputs
- `input_folder`: PDF invoices under `02-inputs/Cross charge list/` (automatically falls back to `02-inputs/invoices/`).

## Steps (routine)
1. Place all travel invoice PDFs into the input folder.
2. Run `python 01-system/tools/ops/cross-charge/cross_charge.py` from the repo root.
3. Open the Excel output and spot-check missing fields or totals.

## Outputs
- **Primary**: `03-outputs/cross charge list/travel_cross_charge.xlsx` (sheet `Invoices`)

## Inputs / Downloads
- Source: `02-inputs/Cross charge list/` (fallback `02-inputs/invoices/`)
- Outputs: `03-outputs/cross charge list/`

## Notes
- Logs INFO per file and WARNING when fields are missing; processing continues.
- Amounts are parsed from the first page; GST is derived from the GST line when present.

## Troubleshooting
- If no files are processed, confirm PDFs exist in the input folder and are not encrypted.
- If fields are missing, verify the invoice layout still matches the expected headings (Tax Invoice, Issue Date, Passengers, Invoice Total, GST).

## Change Log
- v0.1 (2025-12-02): initial version
