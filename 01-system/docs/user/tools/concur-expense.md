# Concur 費用轉檔工具  
**類別**: ops  
**版本**: v0.3 (發布: 2025-11-18)  

## 功能  
- 讀取 Concur “Synchronized Accounting Extract” 檔案，彙整各報告 / 科目的金額。  
- 自動依據 Concur 稅碼 (含 Q2/Q0 等) 推論 L1/L0，並將同一科目的應稅/免稅項目分行顯示。  
- 產出 SAP 可貼上的 I–N 欄位，並保留 `REPORT TOTAL` 行與 GST 對帳表以快速驗證。  

## 參數 / 輸入  
- `原始檔`：`02-inputs/Concur/<REGION>/` 中的 `.xlsx`（或 `.csv`）；勿覆蓋 NAME ID 參考檔。  
- `Vendor list`：`02-inputs/Payment run raw/<REGION> Vendor list.xlsx`。  
- `Concur/SAP 對照表`：`02-inputs/Concur/<REGION> NAME ID.xlsx`（Employee ID ↔ SAP ID）。  

## 操作步驟  
1. 將最新 Concur extract 及對照表 / Vendor list 放入對應資料夾（AU、NZ 均可）。  
2. 於 repo 根目錄執行 `python 01-system/tools/ops/concur-expense/convert_expenses.py`。  
3. 檢查 `03-outputs/concur-expense/<REGION>/SAP_<REGION>_<來源檔>.xlsx`：  
   - `Summary`：員工、科目、稅碼、Gross / Net / GST 金額。  
   - `SAP_Paste`：列出 Concur ID、SAP ID、Report ID、Submit Date；後方為帳號 I–N 欄，並附 `REPORT TOTAL` 行。  
   - `GST_Check`：比較每份報告的 Gross / Net / GST，確保 Difference = 0（與 Concur 原檔一致）。  

## 輸出  
- `03-outputs/concur-expense/<REGION>/SAP_<REGION>_<來源檔>.xlsx`（含 `Summary`、`SAP_Paste`、`GST_Check` 工作表）。  

## 注意事項  
- 僅處理 `Report Entry Payment Code Name = CASH` 且 `Journal Payer Payment Type Name = Company` 的項目。  
- 若新增員工，需同步更新 `<REGION> NAME ID.xlsx` 以避免 SAP Supplier ID 為空。  
- `REPORT TOTAL` 行僅供對帳，不需貼入 SAP。  

## 疑難排解  
- 若 GST_Check 的 Difference ≠ 0，表示資料與 Concur 原檔不符，請檢查稅碼或金額是否異常。  
- 若找不到對照表或 Vendor list，請確認檔案名稱與位置是否維持預設格式。  

## 版本紀錄  
- v0.3 (2025-11-18): SAP_Paste 格式與報表總額、GST_Check 對帳功能。  
- v0.2 (2025-11-18): 新增 Concur/SAP 對照表與報表總額行。  
- v0.1 (2025-11-18): 初版，支援 AU 付款與 GST 拆分。  
