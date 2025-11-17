#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Google Search Console API 查詢程式
功能：批次查詢 GSC 中每個關鍵字的排名、點擊、曝光數
輸出：Excel 檔案
"""

import pandas as pd
import openpyxl
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import time
import sys

# ========== 配置 ==========
SERVICE_ACCOUNT_FILE = 'gsc-keyword-query-a73ecc893169.json'  # 你下載的 JSON 檔案
GSC_SITE_URL = 'https://pm.shiny.com.tw/'  # 你的網站
KEYWORDS_FILE = 'keywords.csv'  # 你的關鍵字清單（CSV 或 Excel）
OUTPUT_FILE = f'gsc_results_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'  # 輸出檔案名稱
QUERY_DAYS = 90  # 查詢過去 90 天的數據


def authenticate_gsc():
    """使用 Service Account 認證連接 GSC API"""
    scopes = ['https://www.googleapis.com/auth/webmasters.readonly']
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=scopes
    )
    service = build('webmasters', 'v3', credentials=credentials)
    return service


def load_keywords(filename):
    """從 CSV 或 Excel 檔案讀取關鍵字清單"""
    try:
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            df = pd.read_excel(filename)
        else:
            df = pd.read_csv(filename)
        
        # 嘗試尋找關鍵字欄位（常見列名）
        keyword_col = None
        for col in ['keyword', 'keywords', '關鍵字', '關鍵詞', 'query']:
            if col in df.columns:
                keyword_col = col
                break
        
        if keyword_col is None:
            # 如果找不到，使用第一欄
            keyword_col = df.columns[0]
            print(f"警告：未找到標準的關鍵字欄位，使用第一欄 '{keyword_col}'")
        
        keywords = df[keyword_col].dropna().unique().tolist()
        print(f"成功讀取 {len(keywords)} 個關鍵字")
        return keywords
    except Exception as e:
        print(f"讀取關鍵字檔案失敗: {e}")
        sys.exit(1)


def query_gsc_performance(service, keywords, days=90):
    """
    批次查詢 GSC 效能數據
    
    返回：
    [{
        'keyword': '...',
        'clicks': 0,
        'impressions': 0,
        'position': 0.0,
        'ctr': 0.0
    }, ...]
    """
    results = []
    end_date = datetime.now().date()
    start_date = end_date - timedelta(days=days)
    
    print(f"\n開始查詢 {len(keywords)} 個關鍵字...")
    print(f"查詢期間：{start_date} 至 {end_date}\n")
    
    for idx, keyword in enumerate(keywords, 1):
        try:
            # 建立查詢請求
            request_body = {
                'startDate': str(start_date),
                'endDate': str(end_date),
                'dimensions': ['query'],
                'dimensionFilterGroups': [
                    {
                        'filters': [
                            {
                                'dimension': 'query',
                                'operator': 'equals',
                                'value': keyword
                            }
                        ]
                    }
                ],
                'rowLimit': 1
            }
            
            # 發送查詢
            response = service.searchanalytics().query(
                siteUrl=GSC_SITE_URL,
                body=request_body
            ).execute()
            
            # 提取數據
            if 'rows' in response and len(response['rows']) > 0:
                row = response['rows'][0]
                clicks = row.get('clicks', 0)
                impressions = row.get('impressions', 0)
                position = row.get('position', 0)
                ctr = row.get('ctr', 0)
            else:
                clicks = 0
                impressions = 0
                position = 0
                ctr = 0
            
            result = {
                'keyword': keyword,
                'clicks': int(clicks) if clicks else 0,
                'impressions': int(impressions) if impressions else 0,
                'position': round(position, 2) if position else 0,
                'ctr': round(ctr * 100, 2) if ctr else 0  # 轉換為百分比
            }
            results.append(result)
            
            # 進度顯示
            if idx % 50 == 0 or idx == len(keywords):
                print(f"已查詢 {idx}/{len(keywords)} 個關鍵字")
            
            # 尊重 API 限額 - 每秒最多 10 個請求
            time.sleep(0.1)
            
        except Exception as e:
            print(f"查詢失敗 (關鍵字: {keyword}): {str(e)}")
            results.append({
                'keyword': keyword,
                'clicks': 'ERROR',
                'impressions': 'ERROR',
                'position': 'ERROR',
                'ctr': 'ERROR'
            })
    
    return results


def save_to_excel(results, filename):
    """將結果保存為 Excel 檔案"""
    df = pd.DataFrame(results)
    
    # 排序：先按點擊數（降序），再按曝光數（降序）
    df_sorted = df.sort_values(
        by=['clicks', 'impressions'],
        ascending=[False, False],
        na_position='last'
    ).reset_index(drop=True)
    
    # 新增排名欄位
    df_sorted.insert(0, '排名', range(1, len(df_sorted) + 1))
    
    # 寫入 Excel
    df_sorted.to_excel(filename, index=False, sheet_name='GSC 數據')
    
    # 格式化 Excel（可選）
    try:
        workbook = openpyxl.load_workbook(filename)
        worksheet = workbook.active
        
        # 自動調整列寬
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        workbook.save(filename)
    except:
        pass
    
    print(f"\n✅ 結果已保存至：{filename}")
    print(f"共 {len(df_sorted)} 個關鍵字")


def main():
    """主程式"""
    print("=" * 50)
    print("Google Search Console 關鍵字查詢程式")
    print("=" * 50)
    
    try:
        # 1. 認證
        print("\n[步驟 1] 連接 Google Search Console API...")
        service = authenticate_gsc()
        print("✅ 認證成功")
        
        # 2. 讀取關鍵字
        print(f"\n[步驟 2] 讀取關鍵字清單... ({KEYWORDS_FILE})")
        keywords = load_keywords(KEYWORDS_FILE)
        
        # 3. 查詢 GSC 數據
        print(f"\n[步驟 3] 查詢 GSC 數據...")
        results = query_gsc_performance(service, keywords, days=QUERY_DAYS)
        
        # 4. 保存為 Excel
        print(f"\n[步驟 4] 保存結果為 Excel...")
        save_to_excel(results, OUTPUT_FILE)
        
        print("\n" + "=" * 50)
        print("✅ 程式執行完成！")
        print("=" * 50)
        
    except Exception as e:
        print(f"\n❌ 程式執行錯誤: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
