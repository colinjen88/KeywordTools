# KeywordTools 專案重構計畫

## 目標
遵循 MVP（Minimum Viable Product）精神和軟體工程最佳實踐，重構專案結構以利後續擴充。

## 新的目錄結構

```
KeywordTools/
├── src/                      # 源代碼目錄
│   ├── __init__.py
│   ├── core/                 # 核心業務邏輯
│   │   ├── __init__.py
│   │   ├── gsc_client.py    # GSC API 客戶端
│   │   ├── keyword_processor.py  # 關鍵字處理邏輯
│   │   └── data_exporter.py # 數據導出邏輯
│   ├── gui/                  # GUI 相關
│   │   ├── __init__.py
│   │   ├── main_window.py   # 主視窗
│   │   ├── components/      # UI 組件
│   │   │   ├── __init__.py
│   │   │   ├── date_selector.py
│   │   │   ├── filter_panel.py
│   │   │   └── data_table.py
│   │   └── styles.py        # UI 樣式定義
│   ├── utils/                # 工具函數
│   │   ├── __init__.py
│   │   ├── file_utils.py
│   │   ├── validators.py
│   │   └── config.py
│   └── models/               # 數據模型
│       ├── __init__.py
│       └── keyword_data.py
├── data/                     # 數據文件目錄
│   ├── keywords/            # 關鍵字文件
│   │   ├── allKeyWord.csv
│   │   └── allKeyWord_normalized.csv
│   ├── reports/             # 生成的報表
│   └── samples/             # 示例數據
├── tests/                    # 測試目錄
│   ├── __init__.py
│   ├── test_core/
│   ├── test_gui/
│   └── test_utils/
├── docs/                     # 文檔目錄
│   ├── API.md
│   ├── ARCHITECTURE.md
│   ├── STYLE_GUIDE.md
│   └── USER_GUIDE.md
├── scripts/                  # 腳本目錄
│   ├── build_exe.ps1
│   └── setup_dev.ps1
├── config/                   # 配置文件目錄
│   ├── default_config.json
│   └── .env.example
├── .github/                  # GitHub 配置
│   └── workflows/
├── requirements.txt          # 依賴列表
├── requirements-dev.txt      # 開發依賴
├── setup.py                  # 安裝配置
├── README.md
├── CHANGELOG.md
├── LICENSE
└── .gitignore
```

## 要刪除/移動的文件

### 刪除（臨時/測試文件）
- `debug_kws.py`
- `debug_tkcalendar.py`
- `test_export.csv`
- `test_export.xlsx`
- `test_mock.csv`
- `test_row_export.csv`
- `gsc_keyword_report.csv`
- `tracked_json_files.txt`
- `KeywordsTool.py` (舊版本)

### 移動
- `allKeyWord*.csv` → `data/keywords/`
- `gsc_keyword_report_*.csv` → `data/reports/`
- `gsc_keyword_report_sample.csv` → `data/samples/`
- `favorites.json` → `config/`
- `build_exe.ps1` → `scripts/`
- `tools/` 內容 → `scripts/` 或 `tests/`
- `*.spec` → `scripts/`

## 重構階段

### Phase 1: 目錄結構調整（本次）
1. 創建新目錄結構
2. 移動現有文件
3. 更新 .gitignore
4. 更新 README.md

### Phase 2: 代碼重構（後續）
1. 拆分 run_gui.py 為模組化組件
2. 提取業務邏輯到 core/
3. 創建數據模型
4. 實現依賴注入

### Phase 3: 測試和文檔（後續）
1. 添加單元測試
2. 完善 API 文檔
3. 創建用戶指南

## 實施步驟

1. ✅ 創建重構計畫文檔
2. ✅ 創建新目錄結構
3. ✅ 移動和清理文件
4. ✅ 更新配置文件
5. ✅ 創建程式風格指導原則
6. ✅ 更新 README
7. ⏳ Commit 變更
