# KeywordTools V1.1

> Google Search Console 關鍵字分析工具

[![Version](https://img.shields.io/badge/version-1.1-blue.svg)](https://github.com/colinjen88/KeywordTools)
[![Python](https://img.shields.io/badge/python-3.9+-green.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-orange.svg)](LICENSE)

---

## 📋 簡介

KeywordTools 是一個專業的 Google Search Console (GSC) 關鍵字分析工具，提供友好的圖形界面，幫助 SEO 專業人員和網站管理員輕鬆查詢、分析和導出關鍵字數據。

### ✨ 核心功能

- 🔍 **GSC 數據查詢**: 批量查詢關鍵字的點擊、曝光、排名等數據
- 📊 **數據分析**: 自動計算統計數據（平均排名、總點擊等）
- 🎯 **智能篩選**: 支持多種篩選條件（關鍵字、排名、點擊等）
- ⭐ **收藏管理**: 標記和管理重要關鍵字
- 📤 **多格式導出**: 支持 CSV 和 Excel 格式
- 🎨 **現代化 UI**: 使用 ttkbootstrap 提供美觀的深色主題
- 📅 **日期選擇器**: 內建日曆選擇器，支持快速選擇日期範圍

---

## 🚀 快速開始

### 前置需求

- Python 3.9 或更高版本
- Google Search Console API 訪問權限
- Service Account JSON 文件

### 安裝

```bash
# 克隆倉庫
git clone https://github.com/colinjen88/KeywordTools.git
cd KeywordTools

# 安裝依賴
pip install -r requirements.txt

# 運行程式
python run_gui.py
```

### 配置

1. 準備 Google Search Console Service Account JSON 文件
2. 準備關鍵字列表 CSV 文件（每行一個關鍵字）
3. 在 GUI 中選擇文件並設置日期範圍
4. 點擊「執行報表」

---

## 📖 使用說明

### 基本流程

1. **設置屬性 URL**: 輸入 Search Console 屬性 URL
2. **選擇日期範圍**: 使用日曆選擇器或快速選擇按鈕
3. **選擇關鍵字文件**: 瀏覽並選擇包含關鍵字的 CSV 文件
4. **選擇 Service Account**: 選擇 GSC API 認證文件
5. **執行報表**: 點擊「執行報表」按鈕
6. **查看結果**: 在表格中查看數據
7. **篩選和排序**: 使用篩選功能精確查找數據
8. **導出數據**: 選擇格式並導出結果

### 快速選擇日期

- **日期區間**: 自定義日期範圍
- **近7天**: 最近 7 天
- **近30天**: 最近 30 天
- **近1季**: 最近 90 天
- **近1年**: 最近 365 天
- **上個月**: 上一個完整月份

### 篩選功能

- **關鍵字篩選**: 包含特定文字的關鍵字
- **數值篩選**: 支持 `>`, `=`, `<` 運算子
  - 排名 > 10
  - 點擊 > 100
  - 曝光 > 1000
- **已標記篩選**: 只顯示收藏的關鍵字

---

## 📁 專案結構

```
KeywordTools/
├── src/                    # 源代碼（規劃中）
├── data/                   # 數據目錄
│   ├── keywords/          # 關鍵字文件
│   ├── reports/           # 生成的報表
│   └── samples/           # 示例數據
├── docs/                   # 文檔
│   ├── ARCHITECTURE.md    # 架構設計
│   ├── STYLE_GUIDE.md     # 程式風格指南
│   └── REFACTORING_PLAN.md # 重構計畫
├── config/                 # 配置文件
├── scripts/                # 腳本
├── tests/                  # 測試（規劃中）
├── run_gui.py             # 主程式
├── gsc_keyword_report.py  # GSC 查詢腳本
├── requirements.txt        # 依賴列表
├── CHANGELOG.md           # 更新日誌
└── README.md              # 本文件
```

---

## 🛠️ 技術棧

- **語言**: Python 3.9+
- **GUI**: Tkinter / ttkbootstrap
- **API**: Google Search Console API
- **數據處理**: Pandas
- **打包**: PyInstaller

---

## 📝 開發

### 開發環境設置

```bash
# 創建虛擬環境
python -m venv .venv

# 激活虛擬環境 (Windows)
.venv\Scripts\activate

# 安裝開發依賴
pip install -r requirements-dev.txt

# 運行測試
pytest
```

### 編碼規範

請參閱 [程式風格指南](docs/STYLE_GUIDE.md)

### 架構設計

請參閱 [架構文檔](docs/ARCHITECTURE.md)

---

## 🗺️ 路線圖

### V1.0 (當前版本)
- ✅ 基本 GSC 數據查詢
- ✅ GUI 界面
- ✅ 數據篩選和排序
- ✅ 收藏功能
- ✅ CSV/Excel 導出

### V1.1 (規劃中)
- [ ] 模組化重構
- [ ] 單元測試
- [ ] 配置文件支持
- [ ] 性能優化

### V2.0 (未來)
- [ ] 數據可視化
- [ ] 多數據源支持
- [ ] 自動化報表
- [ ] Web 版本

---

## 🤝 貢獻

歡迎貢獻！請遵循以下步驟：

1. Fork 本倉庫
2. 創建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交變更 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 開啟 Pull Request

請確保：
- 遵循 [程式風格指南](docs/STYLE_GUIDE.md)
- 添加適當的測試
- 更新相關文檔

---

## 📄 授權

本專案採用 MIT 授權 - 詳見 [LICENSE](LICENSE) 文件

---

## 👤 作者

**Colinjen**

- Email: colinjen88@gmail.com
- GitHub: [@colinjen88](https://github.com/colinjen88)

---

## 🙏 致謝

- Google Search Console API
- ttkbootstrap 團隊
- 所有貢獻者

---

## 📞 支持

如有問題或建議，請：

- 開啟 [Issue](https://github.com/colinjen88/KeywordTools/issues)
- 發送郵件至 colinjen88@gmail.com

---

**KeywordTools V1.0** - Product by [Colinjen](mailto:colinjen88@gmail.com)
