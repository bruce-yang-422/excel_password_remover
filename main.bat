@echo off
title Excel Password Remover

:: Check and set up the environment once at the beginning
call :SETUP_ENVIRONMENT
if %errorlevel% neq 0 (
echo A critical error occurred during environment setup. Exiting.
echo Press any key to exit...
pause >nul
exit /b 1
)

:: Directly execute batch password removal
goto BATCH_REMOVER

:SETUP_ENVIRONMENT
echo.
echo ========================================
echo     Checking and setting up environment...
echo ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
echo Error: Python not found, please install Python 3.7 or higher
echo.
echo Please visit https://www.python.org/downloads/ to download and install Python
echo.
echo Press any key to exit...
pause >nul
exit /b 1
) else (
echo Python environment check passed
)

echo.

if not exist ".venv" (
echo Creating virtual environment...
python -m venv .venv
if errorlevel 1 (
echo Failed to create virtual environment
echo Press any key to exit...
pause >nul
exit /b 1
)
) else (
echo Virtual environment already exists
)

echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo Checking dependencies...
if not exist "requirements.txt" (
echo Error: "requirements.txt" file not found.
echo Please create this file with required dependencies.
echo Press any key to exit...
pause >nul
exit /b 1
)

pip show msoffcrypto-tool >nul 2>&1
if errorlevel 1 (
echo Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
echo Failed to install dependencies
echo Press any key to exit...
pause >nul
exit /b 1
)
) else (
echo Dependencies already installed
)

echo.
echo Environment setup completed!
echo.
exit /b 0


:BATCH_REMOVER
cls
echo.
echo ========================================
echo     Excel Password Removal Tool
echo     Unified File Naming Version
echo ========================================
echo.
if not exist "input" (
echo Creating input folder...
mkdir input
echo Please put Excel files to process in input folder
echo.
echo Press any key to continue...
pause
goto MAIN_MENU
)

:: Check for files in root input folder
dir /b input\*.xlsx input\*.xls input\*.zip input\*.rar >nul 2>&1
if errorlevel 1 (
    set ROOT_FILES_EXIST=1
) else (
    set ROOT_FILES_EXIST=0
)

:: Check for files in platform folders
set PLATFORM_FILES_EXIST=1
for %%d in (Shopee_files MOMO_files PChome_files Yahoo_files ETMall_files mo_store_plus_files coupang_files) do (
    if exist "input\%%d" (
        dir /b "input\%%d\*.xlsx" "input\%%d\*.xls" "input\%%d\*.zip" "input\%%d\*.rar" >nul 2>&1
        if not errorlevel 1 (
            set PLATFORM_FILES_EXIST=0
            goto :files_found
        )
    )
)

:files_found
if %ROOT_FILES_EXIST% neq 0 if %PLATFORM_FILES_EXIST% neq 0 (
echo No files found in input folder or platform folders
echo Please put .xlsx, .xls, .zip or .rar files in:
echo   - input folder (root directory)
echo   - input\Shopee_files folder (Shopee platform files)
echo   - input\MOMO_files folder (MOMO platform files)
echo   - input\PChome_files folder (PChome platform files)
echo   - input\Yahoo_files folder (Yahoo platform files)
echo   - input\ETMall_files folder (ETMall platform files)
echo   - input\mo_store_plus_files folder (MO Store Plus platform files)
echo   - input\coupang_files folder (Coupang platform files)
echo.
echo File naming rule: {shop_id}_{shop_account}_{shop_name}_{execution_date_time}_{serial_number}
echo.
echo Press any key to continue...
pause
goto MAIN_MENU
)

echo Starting batch password removal...
echo.

python scripts/batch_password_remover.py
if errorlevel 1 (
echo Script execution failed
echo Press any key to exit...
pause
exit /b 1
)

echo.
echo ========================================
echo     Execution completed
echo ========================================
echo.
echo Press any key to exit...
pause >nul
exit /b 0
