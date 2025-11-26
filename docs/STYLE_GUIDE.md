# KeywordTools 程式風格指導原則

## 版本
Version: 1.2  
Last Updated: 2025-11-25

## 目標
本文檔定義 KeywordTools 專案的編碼標準和最佳實踐，確保代碼的一致性、可維護性和可擴展性。

---

## 1. 通用原則

### 1.1 設計原則
- **SOLID 原則**：遵循單一職責、開放封閉、里氏替換、接口隔離、依賴倒置原則
- **DRY (Don't Repeat Yourself)**：避免代碼重複
- **KISS (Keep It Simple, Stupid)**：保持簡單
- **YAGNI (You Aren't Gonna Need It)**：不要過度設計

### 1.2 MVP 精神
- 優先實現核心功能
- 快速迭代，持續改進
- 保持代碼模組化，便於擴展
- 避免過早優化

---

## 2. Python 編碼規範

### 2.1 基本規範
遵循 **PEP 8** Python 編碼風格指南

```python
# 正確示例
def fetch_keyword_data(keyword: str, start_date: str, end_date: str) -> dict:
    """
    獲取關鍵字數據
    
    Args:
        keyword: 關鍵字字符串
        start_date: 開始日期 (YYYY-MM-DD)
        end_date: 結束日期 (YYYY-MM-DD)
    
    Returns:
        包含關鍵字數據的字典
    
    Raises:
        ValueError: 當日期格式不正確時
    """
    if not _validate_date_format(start_date):
        raise ValueError(f"Invalid date format: {start_date}")
    
    return _query_gsc_api(keyword, start_date, end_date)
```

### 2.2 命名規範

#### 變數和函數
- 使用 `snake_case`
- 名稱應具描述性
- 布林值使用 `is_`, `has_`, `can_` 前綴

```python
# 正確
user_count = 10
is_valid = True
has_permission = False

def calculate_average_position(positions: list) -> float:
    pass

# 錯誤
userCount = 10  # 應使用 snake_case
x = True  # 不具描述性
```

#### 類別
- 使用 `PascalCase`
- 名稱應為名詞

```python
# 正確
class KeywordProcessor:
    pass

class GSCClient:
    pass

# 錯誤
class keyword_processor:  # 應使用 PascalCase
    pass
```

#### 常數
- 使用 `UPPER_SNAKE_CASE`
- 定義在模組頂部或專門的 config 文件中

```python
# 正確
MAX_KEYWORDS = 1000
DEFAULT_DATE_FORMAT = '%Y-%m-%d'
API_TIMEOUT_SECONDS = 30
```

#### 私有成員
- 使用單下劃線 `_` 前綴表示內部使用
- 使用雙下劃線 `__` 前綴表示強私有（name mangling）

```python
class DataProcessor:
    def __init__(self):
        self._cache = {}  # 內部使用
        self.__secret_key = "xxx"  # 強私有
    
    def _internal_method(self):  # 內部方法
        pass
    
    def public_method(self):  # 公開方法
        pass
```

### 2.3 類型提示
- 使用類型提示提高代碼可讀性
- Python 3.9+ 使用內建類型（list, dict 等）
- 複雜類型使用 typing 模組

```python
from typing import Optional, Union, List, Dict, Callable

def process_keywords(
    keywords: List[str],
    config: Dict[str, any],
    callback: Optional[Callable] = None
) -> List[Dict[str, Union[str, int, float]]]:
    """處理關鍵字列表"""
    results = []
    for keyword in keywords:
        data = _fetch_data(keyword)
        results.append(data)
        if callback:
            callback(keyword)
    return results
```

### 2.4 文檔字符串
使用 **Google Style** docstrings

```python
def export_to_csv(data: List[dict], filepath: str, encoding: str = 'utf-8-sig') -> bool:
    """
    將數據導出為 CSV 文件
    
    Args:
        data: 要導出的數據列表，每個元素為字典
        filepath: 輸出文件路徑
        encoding: 文件編碼，默認為 'utf-8-sig'
    
    Returns:
        成功返回 True，失敗返回 False
    
    Raises:
        IOError: 當文件無法寫入時
        ValueError: 當數據格式不正確時
    
    Example:
        >>> data = [{'keyword': '測試', 'clicks': 100}]
        >>> export_to_csv(data, 'output.csv')
        True
    """
    try:
        with open(filepath, 'w', newline='', encoding=encoding) as f:
            # 實現邏輯
            pass
        return True
    except IOError as e:
        logger.error(f"Failed to write file: {e}")
        return False
```

### 2.5 錯誤處理

#### 使用具體的異常類型
```python
# 正確
try:
    data = fetch_data()
except ConnectionError as e:
    logger.error(f"Connection failed: {e}")
    raise
except ValueError as e:
    logger.warning(f"Invalid data: {e}")
    return None

# 錯誤
try:
    data = fetch_data()
except Exception:  # 太寬泛
    pass  # 吞掉異常
```

#### 自定義異常
```python
class KeywordToolsError(Exception):
    """基礎異常類"""
    pass

class GSCAPIError(KeywordToolsError):
    """GSC API 相關錯誤"""
    pass

class InvalidDateRangeError(KeywordToolsError):
    """無效的日期範圍"""
    pass
```

### 2.6 導入順序
遵循 PEP 8 導入順序：

```python
# 1. 標準庫
import os
import sys
from datetime import date, timedelta

# 2. 第三方庫
import pandas as pd
from google.oauth2 import service_account

# 3. 本地模組
from src.core.gsc_client import GSCClient
from src.utils.validators import validate_date_range
```

---

## 3. 架構設計原則

### 3.1 模組化
- 每個模組應有單一職責
- 模組間低耦合，高內聚
- 使用依賴注入而非硬編碼依賴

```python
# 正確：依賴注入
class KeywordReportGenerator:
    def __init__(self, gsc_client: GSCClient, exporter: DataExporter):
        self.gsc_client = gsc_client
        self.exporter = exporter
    
    def generate_report(self, keywords: List[str]) -> str:
        data = self.gsc_client.fetch_data(keywords)
        return self.exporter.export(data)

# 錯誤：硬編碼依賴
class KeywordReportGenerator:
    def __init__(self):
        self.gsc_client = GSCClient()  # 硬編碼
        self.exporter = CSVExporter()  # 硬編碼
```

### 3.2 分層架構
```
Presentation Layer (GUI)
    ↓
Business Logic Layer (Core)
    ↓
Data Access Layer (API/File)
```

- GUI 層只負責顯示和用戶交互
- 業務邏輯層處理核心功能
- 數據訪問層負責與外部系統交互

### 3.3 配置管理
- 使用配置文件而非硬編碼
- 支持環境變數覆蓋
- 提供合理的默認值

```python
# config/default_config.json
{
    "gsc": {
        "api_timeout": 30,
        "max_retries": 3
    },
    "export": {
        "default_format": "csv",
        "encoding": "utf-8-sig"
    }
}

# src/utils/config.py
import json
import os

class Config:
    def __init__(self, config_path: str = 'config/default_config.json'):
        with open(config_path) as f:
            self._config = json.load(f)
    
    def get(self, key: str, default=None):
        """支持點號分隔的鍵，如 'gsc.api_timeout'"""
        keys = key.split('.')
        value = self._config
        for k in keys:
            value = value.get(k)
            if value is None:
                return default
        # 環境變數覆蓋
        env_key = '_'.join(keys).upper()
        return os.getenv(env_key, value)
```

---

## 4. GUI 開發規範

### 4.1 組件化
- 將 UI 拆分為可重用組件
- 每個組件負責單一功能
- 使用事件驅動架構

```python
# src/gui/components/filter_panel.py
class FilterPanel(ttk.Frame):
    """篩選面板組件"""
    
    def __init__(self, parent, on_filter_change: Callable = None):
        super().__init__(parent)
        self.on_filter_change = on_filter_change
        self._setup_ui()
    
    def _setup_ui(self):
        """設置 UI 元素"""
        # 創建 UI 組件
        pass
    
    def get_filter_criteria(self) -> dict:
        """獲取當前篩選條件"""
        return {
            'column': self.column_var.get(),
            'operator': self.operator_var.get(),
            'value': self.value_var.get()
        }
```

### 4.2 樣式管理
- 集中管理樣式定義
- 使用常數而非魔法數字
- 支持主題切換

```python
# src/gui/styles.py
class AppStyles:
    # 顏色
    PRIMARY_COLOR = '#1565c0'
    SECONDARY_COLOR = '#808080'
    BACKGROUND_COLOR = '#ffffff'
    
    # 字體
    DEFAULT_FONT = ('Segoe UI', 10)
    HEADING_FONT = ('Segoe UI', 12, 'bold')
    SMALL_FONT = ('Segoe UI', 8)
    
    # 間距
    PADDING_SMALL = 4
    PADDING_MEDIUM = 8
    PADDING_LARGE = 12
    
    @staticmethod
    def configure_styles(style: ttk.Style):
        """配置 ttk 樣式"""
        style.configure('Primary.TButton',
                       background=AppStyles.PRIMARY_COLOR,
                       font=AppStyles.DEFAULT_FONT)
```

---

## 5. 測試規範

### 5.1 測試結構
```
tests/
├── test_core/
│   ├── test_gsc_client.py
│   └── test_keyword_processor.py
├── test_gui/
│   └── test_components.py
└── test_utils/
    └── test_validators.py
```

### 5.2 測試命名
- 測試文件：`test_<module_name>.py`
- 測試類：`Test<ClassName>`
- 測試方法：`test_<function_name>_<scenario>`

```python
# tests/test_core/test_keyword_processor.py
import unittest
from src.core.keyword_processor import KeywordProcessor

class TestKeywordProcessor(unittest.TestCase):
    def setUp(self):
        self.processor = KeywordProcessor()
    
    def test_normalize_keyword_removes_whitespace(self):
        result = self.processor.normalize("  測試  ")
        self.assertEqual(result, "測試")
    
    def test_normalize_keyword_converts_to_lowercase(self):
        result = self.processor.normalize("TEST")
        self.assertEqual(result, "test")
    
    def test_process_keywords_raises_error_on_empty_list(self):
        with self.assertRaises(ValueError):
            self.processor.process([])
```

### 5.3 測試覆蓋率
- 目標：核心業務邏輯 >80%
- 使用 `pytest-cov` 測量覆蓋率
- CI/CD 中自動檢查

---

## 6. 版本控制規範

### 6.1 Commit 訊息格式
使用 **Conventional Commits** 格式：

```
<type>(<scope>): <subject>

<body>

<footer>
```

**Type:**
- `feat`: 新功能
- `fix`: 修復 bug
- `docs`: 文檔更新
- `style`: 代碼格式（不影響功能）
- `refactor`: 重構
- `test`: 測試相關
- `chore`: 構建/工具相關

**示例:**
```
feat(gui): 加入日期選擇器組件

- 使用 tkcalendar 實現日期選擇
- 支持快速選擇預設日期範圍
- 加入日期格式驗證

Closes #123
```

### 6.2 分支策略
- `main`: 穩定版本
- `develop`: 開發分支
- `feature/<name>`: 功能分支
- `fix/<name>`: 修復分支
- `release/<version>`: 發布分支

---

## 7. 文檔規範

### 7.1 代碼文檔
- 所有公開 API 必須有 docstring
- 複雜邏輯加入註釋說明
- 使用 Markdown 格式編寫文檔

### 7.2 README 結構
```markdown
# 專案名稱

## 簡介
## 功能特性
## 安裝
## 使用方法
## 配置
## 開發
## 貢獻
## 授權
```

---

## 8. 性能和安全

### 8.1 性能最佳實踐
- 避免過早優化
- 使用生成器處理大數據
- 適當使用緩存
- 異步處理耗時操作

```python
# 使用生成器
def read_large_csv(filepath: str):
    """逐行讀取大型 CSV 文件"""
    with open(filepath, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        for row in reader:
            yield row

# 使用緩存
from functools import lru_cache

@lru_cache(maxsize=128)
def get_config_value(key: str) -> str:
    """緩存配置值"""
    return config.get(key)
```

### 8.2 安全最佳實踐
- 不在代碼中硬編碼敏感信息
- 使用環境變數或加密配置文件
- 驗證所有用戶輸入
- 使用參數化查詢防止注入

```python
# 正確：使用環境變數
import os
API_KEY = os.getenv('GSC_API_KEY')

# 錯誤：硬編碼
API_KEY = "AIzaSyXXXXXXXXXX"  # 不要這樣做！
```

---

## 9. 日誌規範

### 9.1 日誌級別
- `DEBUG`: 詳細的調試信息
- `INFO`: 一般信息
- `WARNING`: 警告信息
- `ERROR`: 錯誤信息
- `CRITICAL`: 嚴重錯誤

### 9.2 日誌格式
```python
import logging

# 配置日誌
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/app.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# 使用日誌
logger.info("Starting keyword processing")
logger.warning(f"Keyword '{keyword}' not found in results")
logger.error(f"Failed to connect to GSC API: {error}", exc_info=True)
```

---

## 10. 持續改進

### 10.1 代碼審查
- 所有代碼變更需經過審查
- 使用 Pull Request 流程
- 檢查清單：
  - [ ] 符合編碼規範
  - [ ] 有適當的測試
  - [ ] 文檔已更新
  - [ ] 無安全隱患

### 10.2 重構原則
- 小步重構，頻繁提交
- 保持測試通過
- 一次只重構一個問題
- 重構後驗證功能

---

## 附錄

### A. 推薦工具
- **代碼格式化**: `black`, `autopep8`
- **代碼檢查**: `pylint`, `flake8`
- **類型檢查**: `mypy`
- **測試**: `pytest`
- **文檔生成**: `sphinx`

### B. 參考資源
- [PEP 8 -- Style Guide for Python Code](https://www.python.org/dev/peps/pep-0008/)
- [Google Python Style Guide](https://google.github.io/styleguide/pyguide.html)
- [The Zen of Python (PEP 20)](https://www.python.org/dev/peps/pep-0020/)
- [Clean Code by Robert C. Martin](https://www.amazon.com/Clean-Code-Handbook-Software-Craftsmanship/dp/0132350882)

---

**最後更新**: 2025-11-24  
**維護者**: Colinjen (colinjen88@gmail.com)
