# excel_password_remover

## 專案簡介

本專案用於批次移除 Excel 檔案的開啟密碼，適合大量處理寵物商店訂單等 Excel 報表。透過 `passwords.yaml` 設定檔，結合自動化腳本，能快速將加密檔案解密並輸出至指定資料夾。

## 專案結構

```
excel_password_remover
├── input/         # 存放待解密的 Excel 檔案
├── output/        # 存放已解密的 Excel 檔案
├── passwords.yaml # Excel 檔案帳號與密碼設定
├── main.py        # 主程式，批次移除密碼
├── scripts/       # 輔助腳本
│   ├── remover.py # 密碼移除核心函式
│   ├── utils.py   # 設定檔與路徑處理
│   └── tree.py    # 目錄結構產生器
├── requirements.txt # 依賴套件
└── README.md      # 使用說明
```

## 安裝方式

1. 安裝 Python 3.8 以上版本。
2. 安裝必要套件：

```bash
pip install -r requirements.txt
```

## 使用說明

1. 將待解密的 Excel 檔案放入 `input/` 資料夾。
2. 編輯 `passwords.yaml`，格式如下：

```yaml
excel_files:
  - name: 檔案描述
    account: 檔名關鍵字
    password: "密碼"
  # ...可多組
```

3. 執行主程式：

```bash
python main.py
```

4. 解密後的檔案將自動輸出至 `output/` 資料夾。

## 腳本說明

- `main.py`：主控流程，讀取設定檔、比對檔案、呼叫解密。
- `scripts/remover.py`：使用 msoffcrypto-tool 解密 Excel。
- `scripts/utils.py`：讀取 YAML 設定、處理路徑。
- `scripts/tree.py`：產生專案目錄樹（執行 `python scripts/tree.py` 會產生 tree.txt）。

## 其他

- 支援 xlsx/xlsm/xlsb 等 Office Excel 格式。
- 若找不到對應檔案或密碼錯誤，終端機會顯示警告。

## 致謝

- [msoffcrypto-tool](https://github.com/nolze/msoffcrypto-tool)
- [PyYAML](https://pyyaml.org/)

如有問題歡迎提 issue 或聯絡專案作者。

