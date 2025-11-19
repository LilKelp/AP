# PLAYBOOKS
Playbooks map short natural-language phrases to a repeatable series of steps and the expected output paths under 03-outputs/<tool>/....

## How to contribute
1. Confirm the supporting tools/prompts exist and are registered.
2. Describe the triggering phrase, intent, required inputs, and artifact target.
3. List the concise steps and cite every prompt/tool ID and version you rely on.

## Payment list refresh (AU/NZ)
- **Trigger phrases**: "generate AU payment list", "refresh NZ payments", "run regional payment routine"
- **Intent**: Use the payment-list tool to convert AU & NZ raw SAP exports into pivot-ready workbooks with DD filters.
- **Required inputs**: Latest `02-inputs/downloads/<REGION>/<date>.xlsx` files and maintained vendor list files (`AU Vendor list.xlsx`, `NZ Vendor list.xlsx`).
- **Tool**: `payment-list` (ops) — entrypoint `python 01-system/tools/ops/payment-list/payment_routine.py`
- **Steps**:
  1. Ensure the new raw extracts and vendor lists are saved in `02-inputs/downloads/<REGION>/` and `02-inputs/downloads/` respectively.
  2. Run the payment-list tool from the repo root.
  3. Verify that AU/NZ workbooks were regenerated under `03-outputs/payment-list/<REGION>/`.
  4. Share the refreshed output paths and any notable totals with the requester.
- **Outputs**: `03-outputs/payment-list/<REGION>/PMT_<REGION>_<date>.xlsx`

## Concur expense SAP paste (AU)
- **Trigger phrases**: "translate Concur expense file", "build SAP I-N columns", "refresh Concur AU expenses"
- **Intent**: Convert the Concur Synchronized Accounting extract into SAP-ready I–N columns with GST-aware L0/L1 tax codes.
- **Required inputs**: Latest extract `.xlsx` placed under `02-inputs/Concur/AU/` (files containing `Synchronized_Accounting_Extract`) plus any employee/vendor mapping already embedded in Concur.
- **Tool**: `concur-expense` (ops) - entrypoint `python 01-system/tools/ops/concur-expense/convert_expenses.py`
- **Steps**:
  1. Drop the newest Concur extract into `02-inputs/Concur/AU/` (leave `EXAMPLE` files untouched for reference).
  2. Run the concur-expense tool from the repo root.
  3. Review `Summary` sheet totals and confirm GST splits before sharing.
  4. Provide the `SAP_Paste` sheet path (columns I–N) back to the requester for SAP entry.
- **Outputs**: `03-outputs/concur-expense/AU/SAP_AU_<source>.xlsx`
