# 支付清單例行工具
**類別**: ops
**版本**: v0.1 (發布:2025-11-18)

## 功能
- 自動將 `02-inputs/downloads/AU` 與 `02-inputs/downloads/NZ` 的原始 XLSX 匯出轉換為付款清單活頁簿。
- 根據 `AU Vendor list.xlsx`、`NZ Vendor list.xlsx` 填入供應商名稱，並在 Worksheet2 建立可依 DD 篩選的樞紐分析。

## 參數
- `原始檔案`: 置於 `02-inputs/downloads/<REGION>/` 的 SAP 匯出檔 (自動逐一處理)。
- `Vendor list`: `02-inputs/downloads/<REGION> Vendor list.xlsx`，提供 SUPPLIER ID 對應名稱。

## 操作步驟 (例行)
1. 將最新 AU/NZ 原始檔案與 Vendor list 複製到指定資料夾。
2. 在 repo 根目錄執行：`python 01-system/tools/ops/payment-list/payment_routine.py`。
3. 程式會分別輸出 AU、NZ 活頁簿至 `03-outputs/payment-list/<REGION>/`。

## 輸出
- **活頁簿**: `03-outputs/payment-list/<REGION>/PMT_<REGION>_<日期>.xlsx`。
- **樞紐分析**: Worksheet2 `PaymentPivot`，可直接篩選 DD 判斷逾期項目。

## 輸入 / 下載
- 原始資料: `02-inputs/downloads/<REGION>/...`
- Vendor 名單: `02-inputs/downloads/<REGION> Vendor list.xlsx`

## 注意事項
- 確保 Excel COM 可用（Windows + 安裝 Excel）。
- 先關閉輸出活頁簿再重跑，以避免檔案鎖定。

## 疑難排解
- 如遇 Excel COM 無法啟動，可參考 `01-system/docs/agents/TROUBLESHOOTING.md` 或重新啟動 Excel/電腦。

## 版本紀錄
- v0.1 (2025-11-18): 初版，支援 AU/NZ 付款清單例行產生。
