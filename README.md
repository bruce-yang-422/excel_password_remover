# excel_password_remover

## 專案簡介

本專案用於批次移除 Excel 檔案的開啟密碼，並將解密後的檔案另存至 output 資料夾。適合需大量處理受密碼保護 Excel 檔案的自動化需求。

## 主要功能
- 批次移除 input 資料夾內 Excel 檔案的開啟密碼
- 支援多組帳號密碼設定（以 YAML 檔管理）
- 執行過程自動產生詳細 log 檔
- 內建目錄樹結構生成工具（tree.py）

## 安裝需求
- Python 3.7 以上
- 依賴套件：
  - msoffcrypto-tool
  - PyYAML
  - openpyxl
  - 其餘請參考 requirements.txt

安裝套件：
```bash
pip install -r requirements.txt
```

## 使用方式
1. 將待解密的 Excel 檔案放入 `input/` 資料夾。
2. 準備 `passwords.yaml`（格式見下方範例），放於專案根目錄。
3. 執行主程式：
   ```bash
   python main.py
   ```
4. 解密後檔案將輸出至 `output/`，執行紀錄於 `log/`。

## passwords.yaml 格式範例
```yaml
excel_files:
  - account: user001
    name: 王小明
    password: 123456
  - account: user002
    name: 李小華
    password: abcdef
```
- `account`：檔名需包含此字串以對應密碼
- `name`：使用者名稱（僅供 log 記錄）
- `password`：對應 Excel 檔案的開啟密碼

## 目錄樹結構工具（tree.py）
- 執行 `python scripts/tree.py` 會自動掃描專案資料夾，產生 `tree.txt`，方便文件記錄與結構概覽。

## 其他說明
- 每次執行會自動清空 log 資料夾
- 支援多帳號密碼自動比對
- 若有未對應檔案或密碼錯誤，log 會有詳細紀錄

