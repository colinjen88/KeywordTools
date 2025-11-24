# KeywordTools 架構文檔

## 版本
Version: 1.0  
Last Updated: 2025-11-24

---

## 1. 系統概述

KeywordTools 是一個 Google Search Console (GSC) 關鍵字分析工具，提供友好的 GUI 界面，幫助用戶查詢、分析和導出關鍵字數據。

### 1.1 核心功能
- GSC API 數據查詢
- 關鍵字數據分析
- 多種格式導出（CSV, Excel）
- 數據篩選和排序
- 關鍵字收藏管理

### 1.2 技術棧
- **語言**: Python 3.9+
- **GUI 框架**: Tkinter / ttkbootstrap
- **API**: Google Search Console API
- **數據處理**: Pandas
- **打包**: PyInstaller

---

## 2. 架構設計

### 2.1 分層架構

```
┌─────────────────────────────────────────┐
│         Presentation Layer (GUI)         │
│  - Main Window                           │
│  - UI Components (Table, Filter, etc.)  │
└─────────────────┬───────────────────────┘
                  │
┌─────────────────▼───────────────────────┐
│       Business Logic Layer (Core)        │
│  - Keyword Processor                     │
│  - Data Analyzer                         │
│  - Report Generator                      │
└─────────────────┬───────────────────────┘
                  │
┌─────────────────▼───────────────────────┐
│      Data Access Layer (API/File)        │
│  - GSC API Client                        │
│  - File I/O (CSV, Excel)                 │
│  - Configuration Manager                 │
└──────────────────────────────────────────┘
```

### 2.2 模組結構

```
src/
├── core/                    # 核心業務邏輯
│   ├── gsc_client.py       # GSC API 客戶端
│   ├── keyword_processor.py # 關鍵字處理
│   └── data_exporter.py    # 數據導出
│
├── gui/                     # GUI 層
│   ├── main_window.py      # 主視窗
│   ├── components/         # UI 組件
│   │   ├── date_selector.py
│   │   ├── filter_panel.py
│   │   └── data_table.py
│   └── styles.py           # 樣式定義
│
├── utils/                   # 工具模組
│   ├── file_utils.py       # 文件操作
│   ├── validators.py       # 數據驗證
│   └── config.py           # 配置管理
│
└── models/                  # 數據模型
    └── keyword_data.py     # 關鍵字數據模型
```

---

## 3. 核心組件設計

### 3.1 GSC Client

**職責**: 與 Google Search Console API 交互

```python
class GSCClient:
    """GSC API 客戶端"""
    
    def __init__(self, service_account_file: str):
        """初始化客戶端"""
        self.service = self._authenticate(service_account_file)
    
    def fetch_bulk_queries(
        self,
        property_url: str,
        start_date: str,
        end_date: str,
        row_limit: int = 25000
    ) -> List[Dict]:
        """批量查詢關鍵字數據"""
        pass
    
    def fetch_exact_query(
        self,
        property_url: str,
        keyword: str,
        start_date: str,
        end_date: str
    ) -> Optional[Dict]:
        """精確查詢單個關鍵字"""
        pass
```

### 3.2 Keyword Processor

**職責**: 處理關鍵字數據

```python
class KeywordProcessor:
    """關鍵字處理器"""
    
    def normalize_keywords(self, keywords: List[str]) -> List[str]:
        """標準化關鍵字列表"""
        pass
    
    def merge_results(
        self,
        bulk_results: List[Dict],
        exact_results: List[Dict]
    ) -> List[Dict]:
        """合併查詢結果"""
        pass
    
    def calculate_statistics(self, data: List[Dict]) -> Dict:
        """計算統計數據"""
        pass
```

### 3.3 Data Exporter

**職責**: 導出數據到不同格式

```python
class DataExporter:
    """數據導出器"""
    
    def export_to_csv(
        self,
        data: List[Dict],
        filepath: str,
        encoding: str = 'utf-8-sig'
    ) -> bool:
        """導出為 CSV"""
        pass
    
    def export_to_excel(
        self,
        data: List[Dict],
        filepath: str
    ) -> bool:
        """導出為 Excel"""
        pass
```

### 3.4 Main Window

**職責**: 主視窗管理和協調

```python
class MainWindow(tk.Tk):
    """主視窗"""
    
    def __init__(self, config: Config):
        super().__init__()
        self.config = config
        self.gsc_client = None
        self.processor = KeywordProcessor()
        self.exporter = DataExporter()
        self._setup_ui()
    
    def _setup_ui(self):
        """設置 UI"""
        pass
    
    def on_run_report(self):
        """執行報表生成"""
        pass
```

---

## 4. 數據流

### 4.1 報表生成流程

```
User Input (GUI)
    ↓
Validate Input (Utils)
    ↓
Load Keywords (File Utils)
    ↓
Fetch Data (GSC Client)
    ├─ Bulk Query
    └─ Exact Queries (for missing)
    ↓
Process Data (Keyword Processor)
    ├─ Merge Results
    ├─ Calculate Stats
    └─ Apply Filters
    ↓
Display Results (GUI)
    ↓
Export Data (Data Exporter)
```

### 4.2 篩選流程

```
User Selects Filter (Filter Panel)
    ↓
Get Filter Criteria
    ↓
Apply Filter Logic (Keyword Processor)
    ↓
Update Table Display (Data Table)
    ↓
Update Statistics (Stats Panel)
```

---

## 5. 配置管理

### 5.1 配置文件結構

```json
{
  "gsc": {
    "api_timeout": 30,
    "max_retries": 3,
    "row_limit": 25000
  },
  "export": {
    "default_format": "csv",
    "encoding": "utf-8-sig",
    "decimal_places": 1
  },
  "gui": {
    "theme": "superhero",
    "window_size": "780x960",
    "default_date_range": 30
  },
  "paths": {
    "keywords_dir": "data/keywords",
    "reports_dir": "data/reports",
    "favorites_file": "config/favorites.json"
  }
}
```

### 5.2 環境變數

```
GSC_SERVICE_ACCOUNT_FILE=path/to/service_account.json
GSC_PROPERTY_URL=https://example.com/
LOG_LEVEL=INFO
```

---

## 6. 錯誤處理策略

### 6.1 異常層次

```
KeywordToolsError (基礎異常)
├── GSCAPIError (API 相關)
│   ├── AuthenticationError
│   ├── QuotaExceededError
│   └── NetworkError
├── DataProcessingError (數據處理)
│   ├── InvalidDataFormatError
│   └── MissingDataError
└── FileOperationError (文件操作)
    ├── FileNotFoundError
    └── PermissionError
```

### 6.2 錯誤處理原則

1. **捕獲具體異常**: 避免使用 `except Exception`
2. **記錄詳細日誌**: 包含上下文信息
3. **用戶友好提示**: GUI 顯示可理解的錯誤訊息
4. **優雅降級**: 部分失敗不影響整體功能

---

## 7. 性能考量

### 7.1 優化策略

1. **異步處理**: 使用線程處理 API 請求，避免 GUI 凍結
2. **批量查詢**: 優先使用 bulk query 減少 API 調用
3. **緩存機制**: 緩存配置和常用數據
4. **懶加載**: 大數據集分頁加載
5. **生成器**: 處理大文件時使用生成器

### 7.2 資源限制

- API 請求：遵循 GSC API 配額限制
- 內存：單次最多處理 10,000 條記錄
- 文件大小：CSV 最大 50MB

---

## 8. 安全考量

### 8.1 敏感數據保護

1. **Service Account**: 不提交到版本控制
2. **API Keys**: 使用環境變數
3. **用戶數據**: 本地存儲，不上傳

### 8.2 輸入驗證

- 日期格式驗證
- 文件路徑驗證
- URL 格式驗證
- SQL 注入防護（如使用數據庫）

---

## 9. 擴展性設計

### 9.1 插件架構（未來）

```python
class DataSourcePlugin:
    """數據源插件接口"""
    
    def fetch_data(self, **kwargs) -> List[Dict]:
        raise NotImplementedError

class GSCDataSource(DataSourcePlugin):
    """GSC 數據源實現"""
    pass

class AnalyticsDataSource(DataSourcePlugin):
    """Google Analytics 數據源實現"""
    pass
```

### 9.2 導出格式擴展

```python
class ExportStrategy:
    """導出策略接口"""
    
    def export(self, data: List[Dict], filepath: str) -> bool:
        raise NotImplementedError

class CSVExportStrategy(ExportStrategy):
    pass

class ExcelExportStrategy(ExportStrategy):
    pass

class JSONExportStrategy(ExportStrategy):
    pass
```

---

## 10. 測試策略

### 10.1 測試層次

1. **單元測試**: 測試單個函數/方法
2. **集成測試**: 測試模組間交互
3. **UI 測試**: 測試 GUI 組件
4. **端到端測試**: 測試完整流程

### 10.2 Mock 策略

- Mock GSC API 響應
- Mock 文件系統操作
- Mock 用戶輸入

---

## 11. 部署

### 11.1 打包流程

```
1. 運行測試
2. 更新版本號
3. 生成 requirements.txt
4. PyInstaller 打包
5. 測試可執行文件
6. 創建發布包
```

### 11.2 發布清單

- [ ] 可執行文件 (.exe)
- [ ] README.md
- [ ] 示例配置文件
- [ ] 示例數據
- [ ] 用戶手冊

---

## 12. 未來規劃

### 12.1 短期（1-3 個月）
- [ ] 完成模組化重構
- [ ] 添加單元測試
- [ ] 實現配置管理
- [ ] 優化性能

### 12.2 中期（3-6 個月）
- [ ] 支持多數據源
- [ ] 添加數據可視化
- [ ] 實現自動化報表
- [ ] 多語言支持

### 12.3 長期（6-12 個月）
- [ ] Web 版本
- [ ] 雲端同步
- [ ] 協作功能
- [ ] AI 輔助分析

---

**維護者**: Colinjen (colinjen88@gmail.com)  
**最後更新**: 2025-11-24
