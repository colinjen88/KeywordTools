# GSC 關鍵字批次報表工具

這個小工具會讀入一份關鍵字清單（CSV），然後從 Google Search Console 擷取指定期間內每個關鍵字的 Clicks、Impressions、Position，並輸出成 CSV 報表。

主要特性
- 先嘗試一次 bulk 查詢（取得前 N 筆 query），可以快速覆蓋大多數關鍵字
- 對 bulk 未命中的關鍵字以精確查詢補上
- 支援 Service Account 或 OAuth2 認證

準備工作
1. 在 Google Cloud Console 建立 Service Account 並產生 JSON 金鑰，或建立 OAuth client 憑證（Desktop）。
2. 若使用 Service Account，請確保該 service account 被授予 Search Console 的權限（通常需要把對應的 email 加到 Search Console 的使用者列表），或使用 G Suite domain-wide delegation 並設定 `--delegated-user`。

安裝
```powershell
python -m pip install -r requirements.txt
```

使用範例
```powershell
python gsc_keyword_report.py --property "https://example.com" --keywords allKeyWord.csv --start-date 2025-10-01 --end-date 2025-10-31

# 若使用非預設 service account 檔
python gsc_keyword_report.py --property "https://example.com" --keywords allKeyWord.csv --start-date 2025-10-01 --end-date 2025-10-31 --service-account my-sa.json --delegated-user admin@example.com

# 若使用 OAuth client 檔 (會彈出授權頁面)
python gsc_keyword_report.py --property "https://example.com" --keywords allKeyWord.csv --start-date 2025-10-01 --end-date 2025-10-31 --oauth-client client_secret.json
```

測試（mock）模式（不需 GSC 認證，會為每個關鍵字產生隨機樣本數據，方便測試整個流程）：
```powershell
python gsc_keyword_report.py --keywords allKeyWord_normalized.csv --start-date 2025-10-01 --end-date 2025-10-31 --property "https://example.com" --output gsc_keyword_report_sample.csv --mock
```

輸入檔案格式
- `allKeyWord.csv`：每一列為一個關鍵字，第一欄為關鍵字字串（不需 header）。

輸出
- 預設會產生 `gsc_keyword_report.csv`，欄位：`keyword, clicks, impressions, position, found_by`

注意事項
- Search Console API 有 rowLimit 與配額限制。bulk 查詢使用 `--row-limit`（預設 25000）來拿最多前 N 筆 query；若網站自然字詞超過此數，某些關鍵字可能沒被抓到，工具會再對未命中的關鍵字逐一呼叫精確查詢，但會比較慢。
- 使用 service account 時，若出現授權錯誤，請確認該帳號在 Search Console 裡有足夠權限或使用 `--delegated-user`。

如果你要我：
- 幫你把 `allKeyWord.csv` 放到 repo，或
- 幫你執行一次（請提供 property URL 與希望的日期範圍），
請回覆我下一步。祝順利！

新增 GUI 使用方式

你可以用內建的簡易 GUI 來選擇 `property`、`start-date` 與 `end-date`，並選擇輸出為 CSV 或 Excel（或兩者）：

啟動 GUI：
```powershell
python run_gui.py
```

GUI 功能說明：
- `Property`：輸入 Search Console 的完整 property URL（例如：`https://pm.shiny.com.tw/`）。
- `Start date` / `End date`：查詢期間（格式 YYYY-MM-DD）。
- `Keywords file`：預設使用 `allKeyWord_normalized.csv`（若尚未產生請先執行 `normalize_keywords.py`）。
- `Use mock data`：勾選會用模擬資料（不需 GSC 認證），方便先測試整套流程。
- `CSV` / `Excel (.xlsx)`：選擇輸出格式；若要輸出 Excel，請先安裝 `pandas` 與 `openpyxl`。

補充說明：
- 自動載入 CSV：在 GUI 右下的按鈕列中有一個勾選框 `自動載入 CSV（偵測目錄中新產生的 CSV 並自動載入）`。若勾選，GUI 會監控當前工作目錄（repo 根目錄）中最新的 `.csv` 檔案，當檔案新產生或修改時會自動載入表格顯示（適合在外部執行 CLI 並讓 GUI 自動顯示結果）。可取消勾選以避免自動載入。
- 認證與安全（變更）：為了防止不小心使用錯誤憑證或將 Service Account 金鑰一起 Commit，GUI/CLI 現在**強制**需要使用者明確選擇一個有效的 credential：
	- GUI：請在 `Service account JSON（選填）` 欄位中選擇一個 JSON 檔案 (或使用 OAuth client)，若未提供會拒絕執行。
	- CLI：請在 `--service-account` 或 `--oauth-client` 參數中指定。
	- 注意：我們不再把 `GSC_SERVICE_ACCOUNT` 或 `GOOGLE_APPLICATION_CREDENTIALS` 當作自動 fallback（以避免意外在其他環境中使用不安全憑證）；你可把環境變數用於自動化 pipeline，但 GUI 會要求明確檔案路徑。
	- 提示：若你選擇的 JSON 檔位於 repo 目錄（或 subfolder），GUI 會跳出警告，提醒你不要把金鑰加入版本控制。
	- Note (CSV encoding): PowerShell's redirection operator `>` often writes files as UTF-16 LE by default. The GUI now tries multiple encodings (utf-8-sig, utf-8, utf-16, cp950, cp936, latin1) to detect file encoding and will log the detected encoding when auto-loading. To avoid ambiguity, prefer creating CSVs as UTF-8, for example:
		```powershell
		"query,clicks,impressions,position" | Out-File -FilePath .\latest.csv -Encoding utf8
		```
- 輸出格式位置：`輸出格式`下拉已移到按鈕列左側，用來選擇 Save 時匯出的格式（CSV 或 Excel）。
- 快速區間按鈕：GUI 提供 `近7天`、`近30天`、`近1季`、`近1年` 與 `上個月` 等快捷按鈕；若使用快捷按鈕查詢，狀態欄會顯示預設名稱（例如 `查詢完成_近7天` 或 `查詢完成_上個月`）。
- 結果表格說明：結果表格包含欄位 `關鍵字`、`排名`、`點擊`、`曝光` 與 `點擊率`（CTR），數值欄位會以右對齊並有額外右側 padding。表格支援點擊標題欄做雙向排序（點一下升冪、再點一下降冪），並在標題顯示箭頭 ▲/▼。排序後表格會重新套用交替列底色以維持清晰性。
- 匯出檔案命名：匯出時會自動為檔名加入當日日期與查詢區間，例如 `gsc_keyword_report_20251118查詢(20251101-20251130).csv`。

若要我直接在本機執行真實查詢，請回覆以下資訊：
- `property`（Search Console property URL）
- `start-date` 與 `end-date`
- 是否提供 service account JSON（例如 `gsc-key-new.json`）或上傳 OAuth `client_secret.json`。

