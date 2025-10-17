# Excel 密碼移除工具

一個專為處理 Excel 檔案密碼保護而設計的自動化工具，支援批次處理、壓縮檔案解壓縮和智能檔案重新命名。

## 🚀 主要功能

- **🔓 自動密碼破解**：使用預設密碼本自動破解 Excel 檔案密碼
- **📦 壓縮檔案處理**：支援 ZIP/RAR 壓縮檔案的解壓縮和處理
- **🏷️ 智能重新命名**：根據店家資訊自動重新命名檔案
- **📊 詳細日誌記錄**：生成完整的處理報告和錯誤日誌
- **🔄 檔案衝突處理**：自動備份重複檔案，避免覆蓋

## 📁 檔案命名規則

處理後的檔案會按照以下格式重新命名：

```
{shop_id}_{shop_account}_{shop_name}_{執行日期時間}.xlsx
```

**範例：**
- `SH0021_yogurtmeow168_優格小喵_20250116_143052.xlsx`
- `SH0001_petboss5566_萌寵要當家_20250116_143052.xlsx`

## 🛠️ 系統需求

- **作業系統**：Windows 10/11
- **Python**：3.7 或以上版本
- **依賴套件**：
  - `msoffcrypto-tool` - Excel 密碼處理
  - `rarfile` - RAR 檔案解壓縮
  - `zipfile` - ZIP 檔案解壓縮（Python 內建）

## 📦 安裝說明

1. **下載專案**
   ```bash
   git clone <repository-url>
   cd excel_password_remover
   ```

2. **安裝依賴**
   ```bash
   pip install -r requirements.txt
   ```

3. **設定環境**
   - 確保 `mapping/shops_master.json` 檔案存在
   - 建立必要的資料夾結構

## 🎯 使用方法

### 方法一：直接執行（推薦）

**Windows 用戶：**
```bash
# 雙擊執行
main.bat
```

**PowerShell 用戶：**
```powershell
# 執行 PowerShell 腳本
.\menu.ps1
```

### 方法二：Python 直接執行

```bash
python scripts/batch_password_remover.py
```

## 📂 資料夾結構

```
excel_password_remover/
├── input/                    # 放置需要處理的檔案
├── output/                   # 處理後的檔案輸出位置
├── log/                      # 執行日誌檔案
├── backup/                   # 檔案衝突備份
├── mapping/                  # 店家資料和密碼本
│   ├── shops_master.json     # 店家資料和密碼
│   └── A02_Shops_Master - Shops_Master.csv
├── scripts/                  # Python 腳本檔案
│   ├── batch_password_remover.py  # 主要處理腳本
│   ├── csv_to_json.py        # CSV 轉 JSON 工具
│   └── TreeMaker.py          # 目錄樹生成工具
├── main.bat                  # Windows 批次檔
├── menu.ps1                  # PowerShell 腳本
└── requirements.txt          # Python 依賴套件
```

## 🔧 支援的檔案格式

### Excel 檔案
- `.xlsx` - Excel 2007 及以上版本
- `.xls` - Excel 97-2003 版本

### 壓縮檔案
- `.zip` - ZIP 壓縮檔案
- `.rar` - RAR 壓縮檔案

## 📋 處理流程

1. **載入資料**：讀取 `mapping/shops_master.json` 中的店家資料和密碼
2. **掃描檔案**：檢查 `input/` 目錄中的所有檔案
3. **解壓縮**：處理壓縮檔案並提取 Excel 檔案
4. **密碼破解**：嘗試所有可用密碼破解 Excel 檔案
5. **重新命名**：根據店家資訊重新命名檔案
6. **輸出結果**：將處理後的檔案移動到 `output/` 目錄
7. **生成日誌**：記錄處理結果和錯誤資訊

## ⚙️ 配置說明

### 店家資料格式

`mapping/shops_master.json` 檔案包含以下資訊：

```json
{
  "password_index": {
    "密碼": {
      "platform": "平台名稱",
      "shop_id": "店家ID",
      "shop_account": "店家帳號",
      "shop_name": "店家名稱",
      "shop_status": "狀態",
      "Universal Password": "通用密碼",
      "Report Download Password": "報表下載密碼"
    }
  }
}
```

### 密碼本更新

如需更新密碼本，請：

1. 修改 `mapping/A02_Shops_Master - Shops_Master.csv`
2. 執行 `python scripts/csv_to_json.py` 轉換為 JSON 格式
3. 重新執行程式

## 📊 輸出結果

### 成功處理
- 檔案會重新命名並移動到 `output/` 目錄
- 生成詳細的處理日誌

### 處理失敗
- 錯誤資訊會記錄在日誌中
- 原始檔案保持不變

### 檔案衝突
- 重複檔案會自動備份到 `backup/` 目錄
- 使用時間戳記避免覆蓋

## 🐛 故障排除

### 常見問題

1. **Python 未安裝**
   - 請安裝 Python 3.7 或以上版本
   - 確保 Python 已加入系統 PATH

2. **依賴套件缺失**
   ```bash
   pip install -r requirements.txt
   ```

3. **RAR 檔案無法解壓縮**
   - 安裝 UnRAR 工具
   - 或使用 Chocolatey：`choco install unrar -y`

4. **密碼本檔案不存在**
   - 確保 `mapping/shops_master.json` 檔案存在
   - 檢查檔案路徑和權限

### 日誌檢查

處理完成後，請檢查 `log/` 目錄中的日誌檔案：
- 查看處理結果統計
- 檢查錯誤訊息
- 確認檔案重新命名情況

## 📝 更新日誌

### v2.0.0 (2025-01-16)
- ✅ 整合所有功能到單一腳本
- ✅ 更新檔案命名規則
- ✅ 簡化操作流程
- ✅ 移除互動式選單

### v1.0.0
- ✅ 基本密碼移除功能
- ✅ 壓縮檔案處理
- ✅ 批次處理支援

## 📄 授權條款

本專案採用 MIT 授權條款，詳見 LICENSE 檔案。

## 🤝 貢獻指南

歡迎提交 Issue 和 Pull Request 來改善這個工具！

## 📞 聯絡資訊

如有問題或建議，請透過以下方式聯絡：
- 提交 GitHub Issue
- 發送電子郵件

---

**注意**：使用本工具時請確保遵守相關法律法規，僅用於合法的檔案處理需求。