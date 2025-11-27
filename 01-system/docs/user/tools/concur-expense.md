# Concur Expense Converter
**Category**: ops
**Version**: v0.8 (Released: 2025-11-27)

## What it does
- Reads Concur "Synchronized Accounting Extract" files and aggregates per-report/account totals.
- Merges DR GST lines into matching CR expense lines (multi-key fallback) and allocates GST proportionally when multiple expenses share a key.
- Derives gross from Journal Amount, GST from Total Tax Posted/Tax Posted Amount, and net as gross minus GST after merge; tax code L1/L0 is set from GST (NZ displays Q2/Q0).
- Runs post-merge GST validation (AU 10%, NZ 15%, or zero within tolerance) and flags GST_Check rows with OK/CHECK.
- Generates SAP-ready I-N columns with a REPORT TOTAL row, GST_Check, and Raw_Input; NZ cost centers starting with 80 are converted to 81.

## Inputs / Parameters
- Raw extract: .xlsx or .csv placed under 02-inputs/Concur/<REGION>/ (files containing EXAMPLE or starting with ~$ are skipped; leave the NAME ID mapping files intact).
- Vendor list: 02-inputs/Payment run raw/<REGION> Vendor list.xlsx (supplier name to supplier ID lookup).
- Concur/SAP mapping: 02-inputs/Concur/<REGION> NAME ID.xlsx (Employee ID to SAP Supplier ID).

## Steps
1. Place the latest Concur extract plus the mapping and vendor files in their folders (AU or NZ).
2. From the repo root run python 01-system/tools/ops/concur-expense/convert_expenses.py.
3. Review 03-outputs/concur-expense/<REGION>/SAP_<REGION>_<source>.xlsx:
   - Summary: employee, account, tax code with Gross / Net / GST totals.
   - SAP_Paste: Concur ID, SAP Supplier ID, Report ID, Submit Date, columns I-N, plus a REPORT TOTAL row for validation.
   - GST_Check: Gross / Net / GST by report; Difference should be 0.
   - Raw_Input: original extract for reference.

## Outputs
- 03-outputs/concur-expense/<REGION>/SAP_<REGION>_<source>.xlsx with Summary, SAP_Paste, GST_Check, and Raw_Input sheets.

## Notes
- Processes company-paid lines (Journal Payer Payment Type Name = Company, Report Entry Payment Code Name = Cash) that include a journal account.
- REPORT TOTAL rows are for reconciliation only; do not paste them into SAP.
- Ensure the mapping files stay closed to avoid file locks when running the script.
- AU/NZ mixed items are detected automatically: if GST is materially below the full rate on gross but non-zero, the tool derives taxable vs non-taxable portions and splits into two SAP_Paste lines (L1/L0; NZ displays Q2/Q0) with GST only on the taxable portion; GST_Check shows the derived split and does not auto-correct.

## Troubleshooting
- If GST_Check Difference is not 0, inspect the source tax codes/amounts and fix before rerunning.
- If SAP Supplier IDs are blank, update <REGION> NAME ID.xlsx and the vendor list to include the missing mapping.

## Change Log
- v0.8 (2025-11-27): Auto-detect mixed GST lines (AU/NZ) and split into L1/L0/Q2/Q0 lines based on gross vs GST without user flags.
- v0.7 (2025-11-27): Auto-detect mixed AU GST lines (GST <10% of gross) and split into L1/L0 SAP lines without user flags.
- v0.6 (2025-11-27): Merge DR GST lines into CR expenses with proportional allocation, post-merge GST validation, and GST_Check status flagging.
- v0.5 (2025-11-20): Switched to W/AQ/AR fields for Gross/GST/Net.
- v0.4 (2025-11-20): Used W/AS/AT fields and added AU/NZ GST 10%/15% validation.
- v0.3 (2025-11-18): Added SAP_Paste formatting plus report totals and GST_Check reconciliation.
- v0.2 (2025-11-18): Added Concur/SAP mapping and report total row.
- v0.1 (2025-11-18): Initial AU support with GST split.
