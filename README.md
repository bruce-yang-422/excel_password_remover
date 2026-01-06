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
{shop_id}_{shop_account}_{shop_name}_{執行日期時間}_{流水號}.xlsx
```

**範例：**
- `SH0021_yogurtmeow168_優格小喵_20250116_143052_01.xlsx`
- `SH0001_petboss5566_萌寵要當家_20250116_143052_01.xlsx`
- `MOSP01_TP0007661_愛喵樂MO+_20250116_143052_01.xls`

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
   git clone https://github.com/bruce-yang-422/excel_password_remover.git
   cd excel_password_remover
   ```

2. **安裝依賴**
   ```bash
   pip install -r requirements.txt
   ```

3. **設定環境**
   - 確保 `mapping/shops_master.json` 檔案存在
   - 建立必要的資料夾結構

4. **設定 UnRAR.exe（處理 RAR 檔案需要）**
   - 下載 UnRAR for Windows：https://www.rarlab.com/rar_add.htm
   - 解壓縮下載的檔案，找到 `UnRAR.exe`
   - 將 `UnRAR.exe` 複製到 `scripts/` 資料夾中
   - **注意**：如果不需要處理 RAR 檔案，可以跳過此步驟

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
│   ├── Shopee_files/         # 蝦皮平台檔案
│   ├── MOMO_files/           # MOMO 平台檔案
│   ├── PChome_files/         # PChome 平台檔案
│   ├── Yahoo_files/          # Yahoo 平台檔案
│   ├── ETMall_files/         # ETMall 平台檔案
│   ├── mo_store_plus_files/  # MO Store Plus 平台檔案
│   └── coupang_files/        # Coupang 平台檔案
├── output/                   # 處理後的檔案輸出位置
├── log/                      # 執行日誌檔案
├── temp/                     # 臨時檔案目錄
├── mapping/                  # 店家資料和密碼本
│   ├── shops_master.json     # 店家資料和密碼
│   ├── csv_to_json_converter.py  # CSV 轉 JSON 工具
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
2. **掃描檔案**：檢查 `input/` 目錄及平台資料夾中的所有檔案
3. **平台識別**：根據檔案所在資料夾識別對應平台
4. **解壓縮**：處理壓縮檔案並提取 Excel 檔案
5. **密碼破解**：使用平台特定密碼破解 Excel 檔案
6. **重新命名**：使用統一格式 `{shop_id}_{shop_account}_{shop_name}_{執行日期時間}_{流水號}` 重新命名檔案
7. **輸出結果**：將處理後的檔案移動到 `output/` 目錄
8. **生成日誌**：記錄處理結果和錯誤資訊

## ⚙️ 配置說明

### 店家資料格式

**⚠️ 安全提醒：** 以下範例僅為格式說明，實際使用時請替換為真實的密碼和店家資訊。

`mapping/shops_master.json` 檔案包含以下資訊：

```json
{
  "platform_index": {
    "Shopee": {
      "實際密碼": {
        "platform": "Shopee",
        "shop_id": "SH0021",
        "shop_account": "yogurtmeow168",
        "shop_name": "優格小喵.",
        "shop_status": "Active",
        "Universal Password": "實際密碼",
        "Report Download Password": "實際密碼"
      }
    },
    "MOMO": {
      "實際密碼": {
        "platform": "MOMO",
        "shop_id": "MOMO01",
        "shop_account": "account123",
        "shop_name": "店家名稱",
        "shop_status": "Active",
        "Universal Password": "實際密碼",
        "Report Download Password": "實際密碼"
      }
    }
  }
}
```

### 密碼本更新

如需更新密碼本，請：

1. 修改 `mapping/A02_Shops_Master - Shops_Master.csv`
2. 執行 `python mapping/csv_to_json_converter.py` 轉換為 JSON 格式
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
   - 請確認 `scripts/UnRAR.exe` 檔案存在
   - 如果不存在，請下載 UnRAR for Windows：https://www.rarlab.com/rar_add.htm
   - 將 `UnRAR.exe` 複製到 `scripts/` 資料夾中
   - 重新執行程式

4. **密碼本檔案不存在**
   - 確保 `mapping/shops_master.json` 檔案存在
   - 檢查檔案路徑和權限

### 日誌檢查

處理完成後，請檢查 `log/` 目錄中的日誌檔案：
- 查看處理結果統計
- 檢查錯誤訊息
- 確認檔案重新命名情況

## 📝 更新日誌

### v3.0.0 (2025-10-17)
- ✅ 重構專案架構，改為模組化設計
- ✅ 新增平台分類資料夾支援
- ✅ 統一檔案命名規則為 `{shop_id}_{shop_account}_{shop_name}_{執行日期時間}_{流水號}`
- ✅ 新增 CSV 到 JSON 轉換工具
- ✅ 新增 PowerShell 執行腳本
- ✅ 優化平台特定密碼測試邏輯
- ✅ 保留商店名稱中的點號字符
- ✅ 更新 .gitignore 排除敏感檔案

### v2.0.0 (2025-10-16)
- ✅ 整合所有功能到單一腳本
- ✅ 更新檔案命名規則
- ✅ 簡化操作流程
- ✅ 移除互動式選單

### v1.0.0 (2025-10-15)
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