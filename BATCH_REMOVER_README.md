# Excel 批次密碼移除工具

## 功能說明

這是一個專門用於批次處理 Excel 檔案密碼移除的工具，會自動：

1. **掃描 input 資料夾**中的所有 Excel 檔案（.xlsx, .xls）
2. **根據檔名匹配帳號**，使用 `passwords.yaml` 中的密碼
3. **移除 Excel 檔案密碼**保護
4. **重新命名檔案**為：`{name}_{account}_{原始檔案名稱}.xlsx`
5. **儲存到 output 資料夾**並產生詳細日誌

## 檔案命名規則

### 輸入檔案
- 檔名必須包含對應的 `account` 字串
- 例如：`petboss5566_report.xlsx` 會匹配到 `account: petboss5566`

### 輸出檔案
- 格式：`{name}_{account}_{原始檔案名稱}.xlsx`
- 例如：`萌寵要當家_petboss5566_petboss5566_report.xlsx`

## 使用方式

### 方法一：使用批次檔（推薦）
```bash
# 雙擊執行
run_batch_remover.bat
```

### 方法二：手動執行
```bash
python scripts/batch_password_remover.py
```

## 使用步驟

### 1. 準備檔案
將需要處理的 Excel 檔案放入 `input/` 資料夾：
```
input/
├── petboss5566_report.xlsx
├── dogcatclub5566_data.xlsx
├── petstar5566_orders.xlsx
└── ...
```

### 2. 確認設定
確保 `passwords.yaml` 包含對應的帳號設定：
```yaml
excel_files:
  - name: 萌寵要當家
    account: petboss5566
    password: "725389"
  - name: 汪喵日總匯
    account: dogcatclub5566
    password: "692389"
  - name: 毛寵星人樂園
    account: petstar5566
    password: "693289"
```

### 3. 執行程式
雙擊 `run_batch_remover.bat` 或執行 `python scripts/batch_password_remover.py`

### 4. 查看結果
處理後的檔案會儲存在 `output/` 資料夾：
```
output/
├── 萌寵要當家_petboss5566_petboss5566_report.xlsx
├── 汪喵日總匯_dogcatclub5566_dogcatclub5566_data.xlsx
├── 毛寵星人樂園_petstar5566_petstar5566_orders.xlsx
└── ...
```

## 執行日誌

程式會自動產生詳細的執行日誌，儲存在 `log/` 資料夾：
- 檔名格式：`batch_removal_log_YYYYMMDD_HHMMSS.txt`
- 包含處理統計、成功檔案清單、失敗檔案清單

### 日誌範例
```
🔓 正在處理：petboss5566_report.xlsx
✅ 萌寵要當家 (petboss5566) - 成功移除密碼：萌寵要當家_petboss5566_petboss5566_report.xlsx

🔓 正在處理：dogcatclub5566_data.xlsx
✅ 汪喵日總匯 (dogcatclub5566) - 檔案本身無密碼：汪喵日總匯_dogcatclub5566_dogcatclub5566_data.xlsx

==================================================
📊 處理統計
總檔案數：2
成功處理：2
處理失敗：0

✅ 成功處理的檔案：
  petboss5566_report.xlsx → 萌寵要當家_petboss5566_petboss5566_report.xlsx
  dogcatclub5566_data.xlsx → 汪喵日總匯_dogcatclub5566_dogcatclub5566_data.xlsx
```

## 功能特色

### 🔍 智慧檔案匹配
- 自動根據檔名中的 `account` 匹配對應密碼
- 支援多個帳號同時處理

### 🔓 多種處理模式
- **有密碼檔案**：使用密碼移除保護
- **無密碼檔案**：直接複製並重新命名
- **未加密檔案**：自動識別並處理

### 📝 詳細日誌記錄
- 即時顯示處理進度
- 記錄成功和失敗的檔案
- 提供統計摘要

### 🛡️ 安全處理
- 使用專業的 `msoffcrypto-tool` 函式庫
- 保持原始檔案不變
- 錯誤處理和恢復機制

## 注意事項

1. **檔案命名**：輸入檔案檔名必須包含對應的 `account` 字串
2. **密碼設定**：確保 `passwords.yaml` 中的密碼正確
3. **檔案格式**：僅支援 Excel 檔案（.xlsx, .xls）
4. **資料夾結構**：程式會自動建立 `output/` 和 `log/` 資料夾
5. **原始檔案**：原始檔案不會被修改，處理後的檔案會儲存在 `output/` 資料夾

## 錯誤處理

### 常見錯誤及解決方案

1. **找不到對應帳號**
   - 檢查檔案檔名是否包含正確的 `account` 字串
   - 確認 `passwords.yaml` 中的設定

2. **密碼錯誤**
   - 檢查 `passwords.yaml` 中的密碼是否正確
   - 確認檔案是否真的有密碼保護

3. **檔案損壞**
   - 檢查原始檔案是否可以正常開啟
   - 嘗試手動開啟檔案確認完整性

## 與主程式的差異

| 功能 | 主程式 (main.py) | 批次移除工具 (scripts/batch_password_remover.py) |
|------|------------------|------------------------------------------|
| 壓縮檔案支援 | ✅ | ❌ |
| 檔案重新命名 | ❌ | ✅ |
| 處理範圍 | Excel + 壓縮檔案 | 僅 Excel 檔案 |
| 命名格式 | 保持原檔名 | `{name}_{account}_{原檔名}` |
| 使用場景 | 一般處理 | 批次整理和歸檔 |

## 技術架構

- **核心解密**：使用 `msoffcrypto-tool` 函式庫
- **檔案處理**：Python `pathlib` 模組
- **設定管理**：YAML 格式設定檔
- **日誌記錄**：自建日誌系統，支援中文編碼
- **錯誤處理**：完整的異常處理機制 