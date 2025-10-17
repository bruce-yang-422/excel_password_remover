# Excel 密碼移除工具 - PowerShell 版本
# 設定控制台編碼為 UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "    Excel 密碼移除工具" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# 檢查 Python 環境
Write-Host "檢查 Python 環境..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "錯誤：找不到 Python，請安裝 Python 3.7 或更高版本" -ForegroundColor Red
        Write-Host "請訪問 https://www.python.org/downloads/ 下載並安裝 Python" -ForegroundColor Red
        Read-Host "按 Enter 鍵退出"
        exit 1
    }
    Write-Host "Python 環境檢查通過：$pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "錯誤：無法檢查 Python 環境" -ForegroundColor Red
    Read-Host "按 Enter 鍵退出"
    exit 1
}

# 檢查依賴
Write-Host "檢查依賴套件..." -ForegroundColor Yellow
try {
    $msoffcrypto = python -c "import msoffcrypto; print('msoffcrypto-tool 已安裝')" 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "安裝依賴套件..." -ForegroundColor Yellow
        pip install -r requirements.txt
        if ($LASTEXITCODE -ne 0) {
            Write-Host "依賴套件安裝失敗" -ForegroundColor Red
            Read-Host "按 Enter 鍵退出"
            exit 1
        }
    } else {
        Write-Host "依賴套件檢查通過" -ForegroundColor Green
    }
} catch {
    Write-Host "依賴套件檢查失敗" -ForegroundColor Red
    Read-Host "按 Enter 鍵退出"
    exit 1
}

# 檢查檔案
Write-Host "檢查檔案..." -ForegroundColor Yellow
if (-not (Test-Path "input")) {
    Write-Host "建立 input 資料夾..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path "input" | Out-Null
    Write-Host "請將要處理的 Excel 檔案放入 input 資料夾" -ForegroundColor Cyan
    Write-Host ""
    Read-Host "按 Enter 鍵繼續"
}

# 檢查根目錄檔案
$rootFiles = Get-ChildItem -Path "input" -Include "*.xlsx", "*.xls", "*.zip", "*.rar" -ErrorAction SilentlyContinue
$rootFilesExist = $rootFiles.Count -gt 0

# 檢查平台資料夾檔案
$platformFolders = @("Shopee_files", "MOMO_files", "PChome_files", "Yahoo_files", "ETMall_files", "mo_store_plus_files", "coupang_files")
$platformFilesExist = $false

foreach ($folder in $platformFolders) {
    $folderPath = "input\$folder"
    if (Test-Path $folderPath) {
        $files = Get-ChildItem -Path $folderPath -Include "*.xlsx", "*.xls", "*.zip", "*.rar" -ErrorAction SilentlyContinue
        if ($files.Count -gt 0) {
            $platformFilesExist = $true
            break
        }
    }
}

if (-not $rootFilesExist -and -not $platformFilesExist) {
    Write-Host "在 input 資料夾或平台資料夾中找不到檔案" -ForegroundColor Red
    Write-Host "請將 .xlsx、.xls、.zip 或 .rar 檔案放入：" -ForegroundColor Yellow
    Write-Host "  - input 資料夾（根目錄）" -ForegroundColor Cyan
    Write-Host "  - input\Shopee_files 資料夾" -ForegroundColor Cyan
    Write-Host "  - input\MOMO_files 資料夾" -ForegroundColor Cyan
    Write-Host "  - input\PChome_files 資料夾" -ForegroundColor Cyan
    Write-Host "  - input\Yahoo_files 資料夾" -ForegroundColor Cyan
    Write-Host "  - input\ETMall_files 資料夾" -ForegroundColor Cyan
    Write-Host "  - input\mo_store_plus_files 資料夾" -ForegroundColor Cyan
    Write-Host "  - input\coupang_files 資料夾" -ForegroundColor Cyan
    Write-Host ""
    Read-Host "按 Enter 鍵退出"
    exit 1
}

Write-Host "正在啟動批次密碼移除工具..." -ForegroundColor White
Write-Host ""

try {
    python scripts/batch_password_remover.py
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "    執行完成" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
}
catch {
    Write-Host "錯誤：$($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "按任意鍵退出..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")