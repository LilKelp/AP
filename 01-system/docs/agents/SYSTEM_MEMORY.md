# SYSTEM MEMORY (Lean Logflow)
2025-11-17 - Bootstrap scaffolding :: repo initialized | baseline workspace ready | 03-outputs/bootstrap-installer/
2025-11-18 - AU payment list :: generated PMT workbook from AU source data | payment summary ready for review | 03-outputs/payment-list/PMT_18.11.25.xlsx
2025-11-18 - AU payment pivot :: rebuilt payment workbook with pivot table summary + removed Sheet3 | sheet2 now filterable by DD | 03-outputs/payment-list/PMT_18.11.25.xlsx
2025-11-18 - AU pivot totals/DD :: re-generated pivot so DD is a row field and vendor totals appear per supplier | workbook ready for overdue screening | 03-outputs/payment-list/PMT_18.11.25.xlsx
2025-11-18 - AU&NZ payment routine :: scripted generator + produced AU/NZ pivot workbooks for 18.11.25 | reusable python routine ready | 03-outputs/payment-list/payment_routine.py
2025-11-18 - payment-list tool :: moved generator into registered ops tool + docs/playbook updates | routine now under tools + outputs tracked | 01-system/tools/ops/payment-list/payment_routine.py
2025-11-18 - concur-expense tool :: built script to split GST + produce SAP I-N columns from Concur AU extract | output + docs/playbook updated | 03-outputs/concur-expense/AU/
2025-11-18 - concur-expense id-map :: added Concur↔SAP map + report totals | SAP_Paste now shows IDs + totals | 03-outputs/concur-expense/AU/SAP_AU_AUD_Synchronized_Accounting_Extract_TEST_20251029.xlsx
2025-11-18 - concur-expense gst split :: revamped SAP_Paste format + taxable vs non-taxable grouping + Concur/SAP mapping | see tool + docs | 03-outputs/concur-expense/AU/SAP_AU_AUD_Synchronized_Accounting_Extract_TEST_20251029.xlsx
