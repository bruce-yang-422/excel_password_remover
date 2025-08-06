@echo off
chcp 65001 >nul
title Excel 批次密碼移除工具

echo.
echo ========================================
echo    Excel 批次密碼移除工具
echo ========================================
echo.

:: 檢查 Python 是否安裝
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 錯誤：未找到 Python，請先安裝 Python 3.7 或以上版本
    echo.
    echo 請前往 https://www.python.org/downloads/ 下載並安裝 Python
    pause
    exit /b 1
)

echo ✅ Python 環境檢查通過
echo.

:: 檢查虛擬環境是否存在
if not exist ".venv" (
    echo 📦 建立虛擬環境...
    python -m venv .venv
    if errorlevel 1 (
        echo ❌ 建立虛擬環境失敗
        pause
        exit /b 1
    )
)

:: 啟動虛擬環境
echo 🔄 啟動虛擬環境...
call .venv\Scripts\activate.bat

:: 檢查依賴套件
echo 📋 檢查依賴套件...
pip show msoffcrypto-tool >nul 2>&1
if errorlevel 1 (
    echo 📦 安裝依賴套件...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ❌ 安裝依賴套件失敗
        pause
        exit /b 1
    )
)

echo ✅ 環境設定完成
echo.

:: 檢查 input 資料夾
if not exist "input" (
    echo 📁 建立 input 資料夾...
    mkdir input
    echo ⚠️  請將需要處理的 Excel 檔案放入 input 資料夾
    echo.
    pause
    exit /b 0
)

:: 檢查 input 資料夾是否有檔案
dir /b input\*.xlsx input\*.xls >nul 2>&1
if errorlevel 1 (
    echo ⚠️  input 資料夾中沒有找到 Excel 檔案
    echo    請將 .xlsx 或 .xls 檔案放入 input 資料夾
    echo.
    pause
    exit /b 0
)

echo 🚀 開始執行批次密碼移除...
echo.

:: 執行程式
python scripts/batch_password_remover.py

echo.
echo ========================================
echo    執行完成
echo ========================================
pause 