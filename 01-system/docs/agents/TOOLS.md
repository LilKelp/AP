# TOOLS INDEX
This document mirrors 01-system/configs/tools/registry.yaml for quick human scanning.

| name | category | summary | outputs |
| --- | --- | --- | --- |
| payment-list | ops | Generate AU/NZ payment workbooks from SAP .xls/.xlsx exports (including text-list local-file saves) with supplier names and DD-filterable pivot | 03-outputs/payment-list/ |
| concur-expense | ops | Convert Concur expense extracts into SAP I-N columns using W/AQ/AR for gross/GST/net with AU/NZ GST validation and auto mixed GST splitting | 03-outputs/concur-expense/ |
| cross-charge | ops | Extract travel invoice fields from PDFs into a consolidated Excel cross-charge list | 03-outputs/cross charge list/ |
| sap-fbl1n | ops | Export FBL1N vendor open items via SAP GUI scripting (VBScript local-file default with spreadsheet fallback) | 02-inputs/Payment run raw/, 02-inputs/downloads/ |
| sap-login | ops | Ensure a SAP GUI session is logged in (SAP GUI scripting) for downstream automated pipelines | 03-outputs/sap-login/ |
