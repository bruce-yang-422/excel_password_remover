# Excel 密碼移除工具

## 專案簡介

本專案是一個自動化的 Excel 檔案密碼移除工具，專門用於批次處理受密碼保護的 Excel 檔案。透過 `msoffcrypto-tool` 函式庫，能夠安全且快速地移除 Excel 檔案的開啟密碼，並將解密後的檔案儲存至指定資料夾。

## 主要功能

- 🔓 **批次密碼移除**：自動處理 `input/` 資料夾內的所有 Excel 檔案
- 👥 **多帳號支援**：透過 YAML 設定檔管理多組帳號密碼對應
- 📝 **詳細記錄**：自動產生執行日誌，記錄處理結果和錯誤訊息
- 🛡️ **安全處理**：使用專業的加密函式庫確保資料安全
- 📊 **目錄樹生成**：內建專案結構視覺化工具

## 專案結構

```
excel_password_remover/
├── main.py              # 主程式入口
├── passwords.yaml       # 密碼設定檔
├── requirements.txt     # Python 依賴套件
├── input/              # 待處理的 Excel 檔案
├── output/             # 解密後的檔案輸出
├── log/                # 執行日誌存放
└── scripts/
    ├── remover.py      # 密碼移除核心功能
    ├── utils.py        # 工具函式
    └── tree.py         # 目錄樹生成工具
```

## 安裝需求

### 系統需求
- Python 3.7 或以上版本
- Windows/Linux/macOS 作業系統

### 依賴套件
```
cffi==1.17.1
cryptography==45.0.4
et_xmlfile==2.0.0
msoffcrypto-tool==5.4.2
olefile==0.47
openpyxl==3.1.5
pycparser==2.22
PyYAML==6.0.2
```

### 安裝步驟
```bash
# 1. 複製專案
git clone [專案網址]
cd excel_password_remover

# 2. 建立虛擬環境（建議）
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
# 或
.venv\Scripts\activate     # Windows

# 3. 安裝依賴套件
pip install -r requirements.txt
```

## 使用方式

### 1. 準備檔案
將需要移除密碼的 Excel 檔案放入 `input/` 資料夾。

### 2. 設定密碼對應
編輯 `passwords.yaml` 檔案，設定帳號與密碼的對應關係：

```yaml
excel_files:
  - name: 測試店家A
    account: testaccountA
    password: "123456"
  - name: 測試店家B
    account: testaccountB
    password: "654321"
  - name: 測試店家C
    account: testaccountC
    password: "abcdef"
```

**設定說明：**
- `name`：使用者名稱（用於日誌記錄）
- `account`：帳號識別碼（檔名需包含此字串）
- `password`：對應的 Excel 檔案開啟密碼

### 3. 執行程式
```bash
python main.py
```

### 4. 查看結果
- 解密後的檔案會儲存在 `output/` 資料夾
- 執行日誌會儲存在 `log/` 資料夾，檔名格式：`execution_log_YYYYMMDD_HHMMSS.txt`

## 檔案命名規則

程式會根據檔名中包含的 `account` 字串來匹配對應的密碼。例如：
- 檔案：`測試店家A_testaccountA_Order.all.20250602_20250702.xlsx`
- 對應：`account: testaccountA` 的密碼設定

## 工具腳本

### 目錄樹生成工具
```bash
python scripts/tree.py
```
- 自動掃描專案資料夾結構
- 產生 `tree.txt` 檔案，方便文件記錄
- 自動排除 `.git`、`__pycache__` 等系統資料夾

## 執行日誌說明

程式會自動產生詳細的執行日誌，包含：
- ✅ 成功處理的檔案
- ❌ 處理失敗的檔案及錯誤原因
- ⚠️ 未找到對應密碼的檔案
- ⚠️ 設定檔中有但未找到對應檔案的帳號

## 注意事項

1. **檔案安全**：每次執行前會自動清空 `log/` 資料夾
2. **密碼保護**：請妥善保管 `passwords.yaml` 檔案，避免密碼外洩
3. **檔案格式**：僅支援 Excel 檔案格式（.xlsx, .xls）
4. **檔名規則**：檔名必須包含對應的 `account` 字串才能正確匹配

## 錯誤處理

程式具備完善的錯誤處理機制：
- 自動跳過隱藏檔案（以 `.` 開頭）
- 詳細記錄處理失敗的原因
- 支援多帳號密碼自動比對
- 提供完整的執行狀態回饋

## 技術架構

- **核心解密**：使用 `msoffcrypto-tool` 函式庫
- **設定管理**：YAML 格式設定檔
- **檔案處理**：Python `pathlib` 模組
- **日誌記錄**：自建日誌系統，支援中文編碼

## 授權說明

本專案由宜加寵物用品 Bruce 開發，版本 1.0 (2025-07-02)