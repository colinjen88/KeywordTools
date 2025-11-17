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
