@echo off
chcp 65001 >nul
echo 正在啟動 Excel 密碼移除工具...
echo.

REM 檢查 Python 是否安裝
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 錯誤：找不到 Python
    echo 請先安裝 Python：https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 檢查虛擬環境是否存在
if not exist ".venv" (
    echo 正在建立虛擬環境...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo 錯誤：無法建立虛擬環境
        pause
        exit /b 1
    )
)

REM 啟動虛擬環境並安裝依賴
echo 正在啟動虛擬環境並安裝依賴...
call .venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo 錯誤：無法啟動虛擬環境
    pause
    exit /b 1
)

REM 安裝依賴
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo 錯誤：無法安裝依賴套件
    pause
    exit /b 1
)

REM 執行主程式
echo 正在執行主程式...
python main.py
if %errorlevel% neq 0 (
    echo 錯誤：程式執行失敗
    pause
    exit /b 1
)

echo.
echo 程式執行完成！
pause 