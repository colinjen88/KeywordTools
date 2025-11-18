#!/usr/bin/env python3
"""
gsc_keyword_report.py

從 Google Search Console 擷取關鍵字報表（Position / Clicks / Impressions），
輸入一份關鍵字清單 (CSV)，輸出包含每個關鍵字最新或指定期間的數據。

支援兩種認證方式：
- Service Account (使用 `gsc-keyword-query-*.json`，可選擇委派 `subject`)
- OAuth InstalledAppFlow (若未提供 service account 檔)

用法（簡單）:
  python gsc_keyword_report.py --property "https://example.com" --keywords allKeyWord.csv --start-date 2025-10-01 --end-date 2025-10-31

請先參考 README.md 進行 API 設定。
"""
import argparse
import csv
import os
import sys
import time
import random
from collections import defaultdict

has_google = True
try:
    from google.oauth2 import service_account
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
except Exception:
    # 不要在 import 錯誤時立即退出 — 允許使用 --mock 模式在沒有套件時測試
    has_google = False
    # 顯示簡短提示，後續若嘗試呼叫 API 會再提示錯誤
    print("警告: 找不到 Google API 套件；若要呼叫 GSC API，請先安裝 requirements.txt 裡的套件。")

SCOPES = ["https://www.googleapis.com/auth/webmasters.readonly"]


def authenticate(service_account_file=None, delegated_user=None, oauth_client_file=None):
    if not has_google:
        print("缺少 Google API 套件，無法進行認證。若要測試請使用 --mock 模式。")
        sys.exit(1)
    if service_account_file and os.path.exists(service_account_file):
        creds = service_account.Credentials.from_service_account_file(
            service_account_file, scopes=SCOPES
        )
        if delegated_user:
            creds = creds.with_subject(delegated_user)
        return creds

    if oauth_client_file and os.path.exists(oauth_client_file):
        flow = InstalledAppFlow.from_client_secrets_file(oauth_client_file, SCOPES)
        creds = flow.run_local_server(port=0)
        return creds

    # For security: Do not auto-fallback to environment variable or ADC; require explicit selection.
    print("找不到有效的 service account 或 OAuth client 檔（未提供 / 無效）。請在呼叫時明確指定 --service-account 或 --oauth-client，或使用 --mock。")
    sys.exit(1)


def fetch_bulk_queries(service, site_url, start_date, end_date, row_limit=25000):
    body = {
        "startDate": start_date,
        "endDate": end_date,
        "dimensions": ["query"],
        "rowLimit": row_limit,
    }
    resp = service.searchanalytics().query(siteUrl=site_url, body=body).execute()
    rows = resp.get("rows", [])
    result = {}
    for r in rows:
        key = r.get("keys", [])[0]
        result[key.lower()] = {
            "query": key,
            "clicks": r.get("clicks", 0),
            "impressions": r.get("impressions", 0),
            "position": r.get("position", 0.0),
        }
    return result


def fetch_exact_query(service, site_url, start_date, end_date, keyword):
    body = {
        "startDate": start_date,
        "endDate": end_date,
        "dimensions": ["query"],
        "dimensionFilterGroups": [
            {
                "groupType": "and",
                "filters": [
                    {"dimension": "query", "operator": "equals", "expression": keyword}
                ],
            }
        ],
        "rowLimit": 1,
    }
    resp = service.searchanalytics().query(siteUrl=site_url, body=body).execute()
    rows = resp.get("rows", [])
    if not rows:
        return None
    r = rows[0]
    return {
        "query": r.get("keys", [])[0],
        "clicks": r.get("clicks", 0),
        "impressions": r.get("impressions", 0),
        "position": r.get("position", 0.0),
    }


def load_keywords(path):
    # 支援兩種常見格式：
    # 1) 每列一個關鍵字 (no header)
    # 2) 單列、逗號分隔的一長串關鍵字（例如你上傳的檔案）
    kws = []
    with open(path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.reader(fh)
        rows = list(reader)
        if not rows:
            return []
        # 若只有一列且該列第一欄包含逗號，則以逗號切分
        if len(rows) == 1 and "," in rows[0][0]:
            parts = [p.strip() for p in rows[0][0].split(",") if p.strip()]
            return parts
        # 否則採用每列第一欄為關鍵字
        for row in rows:
            if not row:
                continue
            kws.append(row[0].strip())
    return kws


def write_output(output_path, rows):
    fieldnames = ["keyword", "clicks", "impressions", "position", "found_by"]
    # 若輸出為 .xlsx，使用 pandas 寫入 Excel，否則寫 CSV
    if output_path.lower().endswith(('.xlsx', '.xls')):
        try:
            import pandas as pd

            df = pd.DataFrame(rows, columns=fieldnames)
            df.to_excel(output_path, index=False)
            return
        except Exception:
            # 若 pandas 不可用，回退到 CSV
            pass

    with open(output_path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for r in rows:
            writer.writerow(r)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--property", required=True, help="Search Console property URL, e.g. https://example.com")
    parser.add_argument("--keywords", required=True, help="CSV 檔，第一欄為關鍵字 (no header required)")
    parser.add_argument("--start-date", required=True, help="YYYY-MM-DD")
    parser.add_argument("--end-date", required=True, help="YYYY-MM-DD")
    parser.add_argument("--service-account", default=None, help="service account JSON 路徑 (可選)。若未提供，可透過環境變數 GSC_SERVICE_ACCOUNT 指定路徑")
    parser.add_argument("--delegated-user", default=None, help="若使用 service account 並需委派，填入被委派的帳號 email")
    parser.add_argument("--oauth-client", default=None, help="OAuth client_secret.json，若不使用 service account 可提供此檔進行 OAuth Flow")
    parser.add_argument("--row-limit", type=int, default=25000, help="bulk 查詢的 rowLimit (預設 25000)")
    parser.add_argument("--output", default="gsc_keyword_report.csv", help="輸出 CSV 檔名")
    parser.add_argument("--mock", action="store_true", help="不呼叫 GSC API，使用隨機數據產生樣本報表（方便測試）")
    args = parser.parse_args()

    service = None
    creds = None
    if not args.mock:
        creds = authenticate(args.service_account, args.delegated_user, args.oauth_client)
        service = build("searchconsole", "v1", credentials=creds)

    print("載入關鍵字清單...")
    keywords = load_keywords(args.keywords)
    print(f"載入 {len(keywords)} 個關鍵字")
    out_rows = []
    if args.mock:
        print("使用 mock 模式產生範例數據（不呼叫 GSC API）...")
        random.seed(42)
        for kw in keywords:
            clicks = random.randint(0, 200)
            impressions = clicks * random.randint(1, 50)
            position = round(random.uniform(1, 50), 2) if impressions > 0 else ""
            out_rows.append({
                "keyword": kw,
                "clicks": clicks,
                "impressions": impressions,
                "position": position,
                "found_by": "mock",
            })
    else:
        print("嘗試以 bulk 查詢擷取最多前 rows 的 query 資料（可快速覆蓋大部分關鍵字）...")
        bulk = fetch_bulk_queries(service, args.property, args.start_date, args.end_date, args.row_limit)
        print(f"bulk 查詢取得 {len(bulk)} 筆 query 資料")

        missing = []
        for kw in keywords:
            key = kw.lower()
            if key in bulk:
                d = bulk[key]
                out_rows.append({
                    "keyword": kw,
                    "clicks": d["clicks"],
                    "impressions": d["impressions"],
                    "position": d["position"],
                    "found_by": "bulk",
                })
            else:
                missing.append(kw)

        print(f"{len(missing)} 個關鍵字未在 bulk 結果中發現，將逐一以精確查詢補上（速度較慢）")
        for i, kw in enumerate(missing, 1):
            # 緩慢速率限制保護
            if i % 50 == 0:
                time.sleep(1)
            d = fetch_exact_query(service, args.property, args.start_date, args.end_date, kw)
            if d:
                out_rows.append({
                    "keyword": kw,
                    "clicks": d["clicks"],
                    "impressions": d["impressions"],
                    "position": d["position"],
                    "found_by": "exact",
                })
            else:
                out_rows.append({"keyword": kw, "clicks": 0, "impressions": 0, "position": "", "found_by": "none"})

    print(f"寫出結果到 {args.output} ...")
    write_output(args.output, out_rows)
    print("完成。可用 Excel 或 pandas 開啟 CSV。")


if __name__ == "__main__":
    main()
