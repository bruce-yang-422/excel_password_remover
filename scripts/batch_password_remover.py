#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 密碼移除工具 - 平台分類版本 v3.0.0

主要功能：
    🔓 自動破解 Excel 檔案密碼保護
    📦 處理壓縮檔案（ZIP/RAR）並解壓縮
    🏷️ 統一檔案重新命名：{shop_name}_{shop_id}_{shop_account}_{執行日期時間}_{流水號}
    📊 生成詳細處理報告和日誌
    🔄 自動處理檔案衝突和備份
    🎯 平台分類處理和密碼測試

支援格式：
    - Excel: .xlsx, .xls
    - 壓縮檔: .zip, .rar

支援平台：
    - Shopee (蝦皮)
    - MOMO (MOMO 購物)
    - PChome (PChome 購物)
    - Yahoo (Yahoo 購物)
    - ETMall (東森購物)
    - MO Store Plus
    - Coupang

使用方法：
    python scripts/batch_password_remover.py
    或直接執行 main.bat / menu.ps1

資料夾結構：
    input/
    ├── Shopee_files/         # 蝦皮平台檔案
    ├── MOMO_files/           # MOMO 平台檔案
    ├── PChome_files/         # PChome 平台檔案
    ├── Yahoo_files/          # Yahoo 平台檔案
    ├── ETMall_files/         # ETMall 平台檔案
    ├── mo_store_plus_files/  # MO Store Plus 平台檔案
    └── coupang_files/        # Coupang 平台檔案

處理流程：
    1. 載入店家資料和密碼本 (mapping/shops_master.json)
    2. 掃描 input/ 目錄及平台資料夾中的檔案
    3. 根據檔案所在資料夾識別對應平台
    4. 解壓縮壓縮檔案並提取 Excel 檔案
    5. 使用平台特定密碼破解 Excel 檔案
    6. 統一重新命名檔案並移動到 output/ 目錄
    7. 生成處理報告和詳細日誌

輸出結果：
    - output/ 目錄：處理後的檔案
    - log/ 目錄：詳細處理日誌
    - temp/ 目錄：臨時檔案

檔案命名規則：
    {shop_name}_{shop_id}_{shop_account}_{執行日期時間}_{流水號}
    範例：
    - 優格小喵_SH0021_yogurtmeow168_20250116_143052_01.xlsx
    - 愛喵樂MO+_MOSP01_TP0007661_20250116_143052_01.xls
    - 東森購物_ETM001_541767_20250116_143052_01.xls

平台特定功能：
    - 平台資料夾自動識別
    - 平台特定密碼測試
    - 蝦皮平台 Order.all 檔案過濾
    - 商店名稱點號保留

注意事項：
    - 確保 mapping/shops_master.json 檔案存在
    - 支援檔案名稱衝突自動處理
    - 自動備份重複檔案
    - 支援密碼保護的壓縮檔案
    - 平台密碼僅測試對應平台，避免混淆
"""

import sys
from pathlib import Path
import datetime
import shutil
import tempfile
import json
import zipfile
try:
    import rarfile  # type: ignore
except Exception as _rar_import_error:  # ImportError or other env errors
    rarfile = None  # fallback sentinel
    _RAR_IMPORT_ERROR = _rar_import_error
import msoffcrypto  # type: ignore
from typing import Dict, List, Optional, Union, Any

# =============================================================================
# 工具函數模組 (來自 utils.py)
# =============================================================================

def get_base_path() -> Path:
    """
    取得專案根目錄路徑
    支援 Python 模式、exe 單文件模式、exe 目錄模式
    """
    if getattr(sys, 'frozen', False):
        # PyInstaller 單文件模式：使用 _MEIPASS 臨時目錄
        if hasattr(sys, '_MEIPASS'):
            # 單文件模式：資源在臨時目錄，但資料檔案應該在 exe 同目錄
            return Path(sys.executable).parent
        else:
            # exe 目錄模式：exe 所在目錄
            return Path(sys.executable).parent
    else:
        # Python 模式：腳本所在目錄的父目錄（專案根目錄）
        return Path(__file__).parent.parent

# ==========================================
# 初始化 UnRAR 路徑 (支援打包模式)
# ==========================================
def init_unrar_tool():
    """
    設定 rarfile 的工具路徑。
    支援：開發環境 (scripts/UnRAR.exe) 與 打包環境 (sys._MEIPASS/UnRAR.exe)
    """
    if rarfile is None:
        return

    # 1. 判斷是否為 PyInstaller 打包後的環境
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # 打包模式：工具會被解壓到臨時目錄 (_MEIPASS)
        base_path = Path(sys._MEIPASS)
        unrar_path = base_path / "UnRAR.exe"
    else:
        # 開發模式：工具在 scripts 資料夾 (相對於此腳本)
        base_path = Path(__file__).parent
        unrar_path = base_path / "UnRAR.exe"

    # 2. 強制設定 rarfile 的工具路徑
    # 注意：如果不設定，rarfile 會去系統 PATH 找，找不到就會報錯
    rarfile.UNRAR_TOOL = str(unrar_path)
    
    # Debug 訊息 (可選)
    # print(f"[DEBUG] UnRAR path set to: {unrar_path}")

def load_passwords(json_filename: str = "mapping/shops_master.json") -> Dict[str, List[Dict[str, Any]]]:
    """
    讀取 mapping/shops_master.json
    exe 模式與 py 模式皆從執行檔所在資料夾讀取
    """
    base_path = get_base_path()
    json_path = base_path / json_filename

    if not json_path.exists():
        raise FileNotFoundError(f"找不到 {json_filename}: {json_path}")

    with json_path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    
    # 轉換 JSON 資料為原本 passwords.yaml 的格式
    return convert_json_to_passwords_format(data)

def convert_json_to_passwords_format(json_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    將 JSON 資料轉換為新的 platform_index 格式
    直接返回原始 JSON 資料，因為現在使用 platform_index 結構
    """
    # 直接返回原始 JSON 資料，因為現在使用新的 platform_index 結構
    return json_data

# =============================================================================
# Excel 密碼移除核心模組 (來自 remover.py)
# =============================================================================

def remove_password(input_path: Union[str, Path], output_path: Union[str, Path], password: str) -> bool:
    """
    使用 msoffcrypto-tool 解開 Excel 開啟密碼，另存為 output_path
    若檔案未加密，直接複製
    """
    with open(input_path, "rb") as f_in:
        office_file = msoffcrypto.OfficeFile(f_in)
        try:
            office_file.load_key(password=password)
            with open(output_path, "wb") as f_out:
                office_file.decrypt(f_out)
        except msoffcrypto.exceptions.FileFormatError as e:
            error_msg = str(e)
            if "Unencrypted document" in error_msg:
                # 檔案未加密，直接複製
                shutil.copyfile(input_path, output_path)
            elif "Record not found" in error_msg:
                # Record not found 可能表示檔案已解密或格式異常，直接複製
                shutil.copyfile(input_path, output_path)
            else:
                raise
    
    return True

# =============================================================================
# 壓縮檔案處理核心模組 (來自 compression.py)
# =============================================================================

def extract_zip(zip_path: Union[str, Path], extract_to: Union[str, Path], password: Optional[str] = None) -> List[str]:
    """
    解壓縮 ZIP 檔案
    
    Args:
        zip_path: ZIP 檔案路徑
        extract_to: 解壓縮目標資料夾
        password: 密碼（可選）
    
    Returns:
        list: 解壓縮的檔案列表
    """
    extracted_files = []
    
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            if password:
                zip_ref.setpassword(password.encode('utf-8'))
            
            # 解壓縮所有檔案
            zip_ref.extractall(extract_to)
            
            # 記錄解壓縮的檔案
            for file_info in zip_ref.infolist():
                if not file_info.is_dir():
                    extracted_files.append(file_info.filename)
                    
    except zipfile.BadZipFile:
        raise Exception(f"無效的 ZIP 檔案：{zip_path}")
    except RuntimeError as e:
        if "Bad password" in str(e):
            raise Exception(f"ZIP 檔案密碼錯誤：{zip_path}")
        else:
            raise Exception(f"解壓縮 ZIP 檔案時發生錯誤：{e}")
    
    return extracted_files

def extract_rar(rar_path: Union[str, Path], extract_to: Union[str, Path], password: Optional[str] = None) -> List[str]:
    """
    解壓縮 RAR 檔案
    
    Args:
        rar_path: RAR 檔案路徑
        extract_to: 解壓縮目標資料夾
        password: 密碼（可選）
    
    Returns:
        list: 解壓縮的檔案列表
    """
    extracted_files = []

    # 依賴檢查
    if rarfile is None:
        raise Exception(
            "缺少依賴 'rarfile'，請先執行: python -m pip install -r requirements.txt"
        )

    try:
        with rarfile.RarFile(rar_path, 'r') as rar_ref:
            if password:
                rar_ref.setpassword(password)
            
            # 解壓縮所有檔案
            rar_ref.extractall(extract_to)
            
            # 記錄解壓縮的檔案
            for file_info in rar_ref.infolist():
                if not file_info.is_dir():
                    extracted_files.append(file_info.filename)
                    
    except rarfile.BadRarFile:
        raise Exception(f"無效的 RAR 檔案：{rar_path}")
    except rarfile.RarCannotExec:
        # 系統找不到解壓工具，提示安裝 WinRAR
        raise Exception(
            "系統未找到可用的 RAR 解壓工具。\n"
            "請安裝 WinRAR 以支援 RAR 檔案解壓縮功能。\n"
            "下載網址：https://www.winrar.com.tw/"
        )
    except Exception as e:  # 處理其他所有異常，包括密碼錯誤
        if "password" in str(e).lower():
            raise Exception(f"RAR 檔案密碼錯誤：{rar_path}")
        else:
            raise Exception(f"RAR 檔案處理錯誤：{rar_path} - {e}")
    
    return extracted_files

def extract_compressed_file(file_path: Union[str, Path], extract_to: Union[str, Path], password: Optional[str] = None) -> List[str]:
    """
    根據檔案副檔名自動選擇解壓縮方法
    
    Args:
        file_path: 壓縮檔案路徑
        extract_to: 解壓縮目標資料夾
        password: 密碼（可選）
    
    Returns:
        list: 解壓縮的檔案列表
    """
    file_path = Path(file_path)
    extract_to = Path(extract_to)
    
    # 確保目標資料夾存在
    extract_to.mkdir(parents=True, exist_ok=True)
    
    # 根據副檔名選擇解壓縮方法
    if file_path.suffix.lower() == '.zip':
        return extract_zip(file_path, extract_to, password)
    elif file_path.suffix.lower() == '.rar':
        return extract_rar(file_path, extract_to, password)
    else:
        raise Exception(f"不支援的壓縮檔案格式：{file_path.suffix}")

def is_compressed_file(file_path: Union[str, Path]) -> bool:
    """
    檢查檔案是否為支援的壓縮檔案格式
    
    Args:
        file_path: 檔案路徑
    
    Returns:
        bool: 是否為支援的壓縮檔案
    """
    file_path = Path(file_path)
    return file_path.suffix.lower() in ['.zip', '.rar']

# =============================================================================
# 主要處理邏輯
# =============================================================================

def test_password(file_path: Union[str, Path], password: str) -> tuple[bool, str]:
    """測試密碼是否正確，使用臨時檔案"""
    temp_path = None
    try:
        # 建立臨時檔案來測試，使用原始檔案的副檔名
        file_ext = Path(file_path).suffix.lower()
        with tempfile.NamedTemporaryFile(suffix=file_ext, delete=False) as temp_file:
            temp_path = temp_file.name
        
        # 嘗試移除密碼到臨時檔案
        remove_password(file_path, temp_path, password)
        
        # 如果成功，刪除臨時檔案
        Path(temp_path).unlink()
        return True, "encrypted"  # 返回成功和檔案類型
        
    except Exception as e:
        # 如果失敗，清理臨時檔案（如果存在）
        try:
            if temp_path and Path(temp_path).exists():
                Path(temp_path).unlink()
        except:
            pass
        
        error_msg = str(e)
        # 某些錯誤應該被視為成功（檔案已解密或未加密）
        if "Unencrypted document" in error_msg or "Record not found" in error_msg:
            print(f"   [OK] 檔案已解密或未加密：{error_msg}")
            return True, "unencrypted"  # 返回成功和檔案類型
        else:
            print(f"   [FAIL] 密碼測試失敗：{e}")
            return False, "failed"

def generate_unique_filename(output_dir: Union[str, Path], base_name: str, file_ext: str, timestamp: Optional[str] = None) -> str:
    """
    生成唯一的檔案名稱，統一使用流水號從 01 開始
    
    Args:
        output_dir: 輸出目錄
        base_name: 基礎檔案名稱（不包含副檔名）
        file_ext: 檔案副檔名
        timestamp: 時間戳（可選，如果不提供則自動生成）
    
    Returns:
        str: 唯一的檔案名稱（統一包含流水號）
    """
    if timestamp is None:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # 統一從流水號 01 開始
    sequence = 1
    while sequence < 100:  # 最多 99 個流水號
        filename = f"{base_name}_{timestamp}_{sequence:02d}{file_ext}"
        output_path = Path(output_dir) / filename
        
        if not output_path.exists():
            return filename
        
        sequence += 1
    
    # 如果還是重複，加上微秒時間戳
    microsecond = datetime.datetime.now().microsecond
    filename = f"{base_name}_{timestamp}_{microsecond:06d}{file_ext}"
    return filename

def handle_file_conflict(output_path: Union[str, Path], backup_dir: Union[str, Path]) -> bool:
    """
    處理檔案衝突，將舊檔案移到備份資料夾
    
    Args:
        output_path: 目標檔案路徑
        backup_dir: 備份資料夾路徑
    
    Returns:
        bool: 是否成功處理衝突
    """
    # 確保 Path 物件
    output_path = Path(output_path)
    backup_dir = Path(backup_dir)
    
    try:
        if output_path.exists():
            # 確保備份資料夾存在
            backup_dir.mkdir(parents=True, exist_ok=True)
            
            # 建立備份檔案名稱（加上時間戳）
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"{output_path.stem}_{timestamp}{output_path.suffix}"
            backup_path = backup_dir / backup_filename
            
            # 移動舊檔案到備份資料夾
            shutil.move(str(output_path), str(backup_path))
            print(f"[BACKUP] 檔案衝突處理：{output_path.name} → 備份至 {backup_path}")
            return True
    except Exception as e:
        print(f"[WARN] 處理檔案衝突失敗：{e}")
        return False
    
    return True

def process_compressed_files(input_dir: Union[str, Path], output_dir: Union[str, Path], temp_dir: Union[str, Path], compressed_accounts: List[Dict[str, Any]], log_lines: List[str]) -> List[Path]:
    """
    處理壓縮檔案，展開到 temp 資料夾等待處理
    
    Args:
        input_dir: 輸入資料夾
        output_dir: 輸出資料夾
        temp_dir: 臨時資料夾
        compressed_accounts: 壓縮檔案帳號設定
        log_lines: 日誌行列表
    
    Returns:
        list: 解壓縮後的 Excel 檔案路徑列表
    """
    # 確保 Path 物件
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    temp_dir = Path(temp_dir)
    
    extracted_excel_files = []
    
    # 掃描壓縮檔案
    compressed_files = []
    for file_path in input_dir.iterdir():
        if file_path.is_file() and is_compressed_file(file_path):
            compressed_files.append(file_path)
    
    if not compressed_files:
        return extracted_excel_files
    
    print(f"[ZIP] 發現 {len(compressed_files)} 個壓縮檔案")
    
    # 處理每個壓縮檔案
    for file_path in compressed_files:
        filename = file_path.name
        print(f"\n[ZIP] 正在處理壓縮檔案：{filename}")
        
        # 嘗試所有已知密碼
        success = False
        # 追蹤是否已處理此檔案（用於除錯和統計）
        processed_this_file = False  # 用於追蹤處理狀態
        for account_info in compressed_accounts:
            password = account_info.get("password")
            name = account_info["name"]
            platform = account_info.get("platform", "")
            
            if not password:
                continue
            
            # 特例：如果壓縮檔有密碼，跳過 Shopee 平台的密碼嘗試
            if platform == "Shopee":
                print(f"   [SKIP] 跳過 Shopee 平台密碼：{name}")
                continue
                
            try:
                # 建立臨時解壓縮資料夾
                temp_extract_dir = temp_dir / f"temp_{filename}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
                temp_extract_dir.mkdir(parents=True, exist_ok=True)
                
                # 解壓縮檔案到臨時資料夾
                extracted_files = extract_compressed_file(file_path, temp_extract_dir, password)
                
                # 處理解壓縮後的檔案，直接移動到 output 並重新命名
                for extracted_file in extracted_files:
                    extracted_path = temp_extract_dir / extracted_file
                    if extracted_path.is_file():
                        # 檢查壓縮檔案名稱中間是否包含 TP0007 (MOMO 系列)
                        if "TP0007" in filename:
                            # 重新命名為 MO_Store_Plus_[原始檔名]
                            new_filename = f"MO_Store_Plus_{extracted_file}"
                        else:
                            # 重新命名為 [shop_name]_[shop_id]_[shop_account]_[執行日期時間]_[流水號]
                            shop_id = account_info.get("shop_id", "UNKNOWN")
                            shop_account = account_info.get("account", "UNKNOWN")
                            # 只替換空格，保留點號
                            safe_name = name.replace(' ', '_')
                            file_ext = Path(extracted_file).suffix
                            base_name = f"{safe_name}_{shop_id}_{shop_account}"
                            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                            new_filename = generate_unique_filename(output_dir, base_name, file_ext, timestamp)
                        
                        # 處理檔名衝突
                        output_path = output_dir / new_filename
                        backup_dir = output_dir / "backup"
                        handle_file_conflict(output_path, backup_dir)
                        
                        # 如果是 Excel 檔案，加入處理列表（包括 MO_Store_Plus 檔案）
                        # 注意：檔案暫時不移動到 output，等密碼移除成功後再移動
                        if output_path.suffix.lower() in ['.xlsx', '.xls']:
                            extracted_excel_files.append(extracted_path)  # 使用原始路徑，不是 output 路徑
                
                # 暫時不刪除臨時資料夾，等所有檔案都處理完成後再清理
                # 記錄臨時資料夾路徑，供後續清理使用
                print(f"[TEMP] 臨時資料夾保留：{temp_extract_dir}")
                print(f"[TEMP] 等待所有檔案處理完成後再清理")
                
                success_msg = f"[OK] {name} - 成功解壓縮：{filename} → {len(extracted_files)} 個檔案"
                log_lines.append(success_msg)
                print(success_msg)
                print(f"   [FILE] 解壓縮檔案：{len(extracted_files)} 個")
                print(f"   [STAT] Excel 檔案：{len([f for f in extracted_files if Path(f).suffix.lower() in ['.xlsx', '.xls']])} 個")
                
                success = True
                processed_this_file = True  # 標記此檔案已處理（用於統計）
                print(f"[DEBUG] 檔案 {filename} 處理完成，狀態：{processed_this_file}")
                break  # 解壓縮成功後跳出，避免重複處理同一個壓縮檔案
                
            except Exception as e:
                error_msg = f"[FAIL] {name} - 解壓縮失敗：{filename} - {e}"
                log_lines.append(error_msg)
                print(f"   {error_msg}")
                # 清理臨時資料夾
                try:
                    if 'temp_extract_dir' in locals() and temp_extract_dir.exists():
                        shutil.rmtree(temp_extract_dir)
                except Exception:
                    pass  # 忽略清理錯誤
                continue
        
        if not success:
            error_msg = f"[FAIL] 所有密碼都無法解壓縮：{filename}"
            log_lines.append(error_msg)
            print(error_msg)
        else:
            print(f"[SUCCESS] 壓縮檔案 {filename} 處理完成")
    
    return extracted_excel_files

def process_platform_compressed_files(compressed_files: List[Path], output_dir: Path, temp_dir: Path, platform_index: Dict[str, Any], platform_name: str, log_lines: List[str]) -> List[Path]:
    """
    處理平台資料夾中的壓縮檔案
    """
    extracted_excel_files = []
    
    # 根據平台名稱確定平台類型
    platform_type = platform_name.replace("_files", "").replace("zip", "shopee").replace("xlsx", "shopee")
    
    for compressed_file in compressed_files:
        filename = compressed_file.name
        print(f"[EXTRACT] 正在處理壓縮檔案：{filename}")
        
        # 建立臨時解壓縮目錄
        temp_extract_dir = temp_dir / f"extract_{filename}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        temp_extract_dir.mkdir(exist_ok=True)
        
        try:
            # 先嘗試使用平台密碼解壓縮
            extracted_files = []
            if platform_type in platform_index:
                passwords = platform_index[platform_type]
                print(f"[EXTRACT] 嘗試使用 {platform_type} 平台的 {len(passwords)} 個密碼解壓縮")
                
                for password in passwords.keys():
                    try:
                        print(f"[EXTRACT] 嘗試密碼：{password}")
                        if compressed_file.suffix.lower() == '.zip':
                            extracted_files = extract_zip(compressed_file, temp_extract_dir, password)
                        elif compressed_file.suffix.lower() == '.rar':
                            extracted_files = extract_rar(compressed_file, temp_extract_dir, password)
                        print(f"[EXTRACT] 使用密碼 {password} 成功解壓縮 {len(extracted_files)} 個檔案")
                        break
                    except Exception as e:
                        print(f"[EXTRACT] 密碼 {password} 解壓縮失敗：{e}")
                        continue
            
            # 如果密碼解壓縮失敗，嘗試無密碼解壓縮
            if not extracted_files:
                print(f"[EXTRACT] 密碼解壓縮失敗，嘗試無密碼解壓縮")
                if compressed_file.suffix.lower() == '.zip':
                    extracted_files = extract_zip(compressed_file, temp_extract_dir)
                elif compressed_file.suffix.lower() == '.rar':
                    extracted_files = extract_rar(compressed_file, temp_extract_dir)
                else:
                    print(f"[SKIP] 不支援的壓縮格式：{compressed_file.suffix}")
                    continue
                
                print(f"[EXTRACT] 無密碼解壓縮成功 {len(extracted_files)} 個檔案")
            
            # 處理解壓縮出來的 Excel 檔案
            for extracted_filename in extracted_files:
                extracted_file_path = temp_extract_dir / extracted_filename
                if extracted_file_path.exists() and extracted_file_path.suffix.lower() in ['.xlsx', '.xls']:
                    print(f"[EXTRACT] 發現 Excel 檔案：{extracted_filename}")
                    
                    # 嘗試使用該平台的密碼破解
                    success = try_platform_passwords(extracted_file_path, platform_index, platform_type, output_dir, log_lines)
                    if success:
                        extracted_excel_files.append(extracted_file_path)
                    else:
                        print(f"[EXTRACT] 無法破解 {extracted_filename}，將加入一般處理流程")
                        extracted_excel_files.append(extracted_file_path)
            
        except Exception as e:
            error_msg = f"[EXTRACT] 解壓縮 {filename} 失敗：{e}"
            log_lines.append(error_msg)
            print(error_msg)
            continue
    
    return extracted_excel_files

def try_platform_passwords(file_path: Path, platform_index: Dict[str, Any], platform_type: str, output_dir: Path, log_lines: List[str]) -> bool:
    """
    嘗試使用指定平台的密碼破解檔案（僅限該平台密碼）
    """
    filename = file_path.name
    
    # 獲取該平台的密碼
    if platform_type in platform_index:
        passwords = platform_index[platform_type]
        print(f"[PLATFORM] 僅使用 {platform_type} 平台的 {len(passwords)} 個密碼進行測試")
        
        for password, shop_info in passwords.items():
            print(f"[TEST] 測試 {platform_type} 平台密碼：{password}")
            success, file_type = test_password(file_path, password)
            if success:
                # 密碼正確，建立檔案
                shop_name = shop_info.get("shop_name", "")
                shop_id = shop_info.get("shop_id", "UNKNOWN")
                shop_account = shop_info.get("shop_account", "UNKNOWN")
                
                print(f"[SUCCESS] {platform_type} 平台密碼 {password} 破解成功，對應商店：{shop_name} ({shop_account})")
                
                file_ext = file_path.suffix.lower()
                
                # 統一使用標準格式：{shop_name}_{shop_id}_{shop_account}_{執行日期時間}_{流水號}
                # 只替換空格，保留點號
                safe_name = shop_name.replace(' ', '_')
                base_name = f"{safe_name}_{shop_id}_{shop_account}"
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                new_filename = generate_unique_filename(output_dir, base_name, file_ext, timestamp)

                output_path = output_dir / new_filename
                
                # 處理檔名衝突
                backup_dir = output_dir / "backup"
                handle_file_conflict(output_path, backup_dir)
                
                try:
                    if file_type == "encrypted":
                        # 加密檔案，進行密碼移除
                        remove_password(file_path, output_path, password)
                    else:
                        # 未加密檔案，直接複製
                        shutil.copyfile(file_path, output_path)
                    success_msg = f"[OK] 使用 {platform_type} 平台 {shop_name} ({shop_account}) 密碼成功處理：{new_filename}"
                    log_lines.append(success_msg)
                    print(success_msg)
                    return True
                except Exception as e:
                    error_msg = f"[FAIL] 使用 {platform_type} 平台 {shop_name} ({shop_account}) 密碼處理失敗：{e}"
                    log_lines.append(error_msg)
                    print(error_msg)
                    continue
            else:
                print(f"[FAIL] {platform_type} 平台密碼 {password} 測試失敗")
    else:
        print(f"[WARN] 找不到 {platform_type} 平台的密碼設定")
    
    return False

def process_root_compressed_files(compressed_files: List[Path], output_dir: Path, temp_dir: Path, platform_index: Dict[str, Any], log_lines: List[str]) -> List[Path]:
    """
    處理根目錄中的壓縮檔案（使用所有平台密碼）
    """
    extracted_excel_files = []
    
    for compressed_file in compressed_files:
        filename = compressed_file.name
        print(f"[EXTRACT] 正在處理根目錄壓縮檔案：{filename}")
        
        # 建立臨時解壓縮目錄
        temp_extract_dir = temp_dir / f"extract_{filename}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        temp_extract_dir.mkdir(exist_ok=True)
        
        try:
            # 嘗試解壓縮檔案
            if compressed_file.suffix.lower() == '.zip':
                extracted_files = extract_zip(compressed_file, temp_extract_dir)
            elif compressed_file.suffix.lower() == '.rar':
                extracted_files = extract_rar(compressed_file, temp_extract_dir)
            else:
                print(f"[SKIP] 不支援的壓縮格式：{compressed_file.suffix}")
                continue
            
            print(f"[EXTRACT] 成功解壓縮 {len(extracted_files)} 個檔案")
            
            # 處理解壓縮出來的 Excel 檔案
            for extracted_filename in extracted_files:
                extracted_file_path = temp_extract_dir / extracted_filename
                if extracted_file_path.exists() and extracted_file_path.suffix.lower() in ['.xlsx', '.xls']:
                    print(f"[EXTRACT] 發現 Excel 檔案：{extracted_filename}")
                    # 加入一般處理流程，讓程式嘗試所有平台密碼
                    extracted_excel_files.append(extracted_file_path)
            
        except Exception as e:
            error_msg = f"[EXTRACT] 解壓縮 {filename} 失敗：{e}"
            log_lines.append(error_msg)
            print(error_msg)
            continue
    
    return extracted_excel_files

def main():
    """主程式：批次處理 Excel 檔案密碼移除"""
    
    # 初始化 UnRAR 工具路徑
    init_unrar_tool()
    
    # 取得專案根目錄（使用統一的函數）
    project_root = get_base_path().resolve()

    input_dir = project_root / "input"
    output_dir = project_root / "output"
    output_dir.mkdir(exist_ok=True)
    temp_dir = project_root / "temp"
    temp_dir.mkdir(exist_ok=True)
    log_dir = project_root / "log"
    passwords_path = "mapping/shops_master.json"

    # 清空並建立 log 資料夾
    if log_dir.exists():
        shutil.rmtree(log_dir)
    log_dir.mkdir(exist_ok=True)

    # 建立 log 檔案
    log_path = log_dir / f"batch_removal_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    # 讀取密碼設定
    try:
        data = load_passwords(passwords_path)
        platform_index = data.get("platform_index", {})
        shops_data = data.get("shops", [])
        
        # 建立帳號到商店資訊的映射
        excel_accounts = {}
        for shop in shops_data:
            account = shop.get("shop_account", "")
            if account:
                excel_accounts[account] = shop
        
        compressed_accounts = data.get("compressed_files", [])
        print(f"[OK] 成功載入 {len(excel_accounts)} 個 Excel 帳號設定")
        print(f"[OK] 成功載入 {len(compressed_accounts)} 個壓縮檔案設定")
        print(f"[OK] 成功載入 {len(platform_index)} 個平台索引")
    except Exception as e:
        print(f"[FAIL] 載入 mapping/shops_master.json 失敗：{e}")
        return

    log_lines = []
    processed_files = []
    failed_files = []

    # 處理壓縮檔案
    extracted_excel_files = process_compressed_files(input_dir, output_dir, temp_dir, compressed_accounts, log_lines)

    # 掃描 input 資料夾中的 Excel 檔案（支援平台分類資料夾）
    excel_files = []
    platform_folders = ["Shopee_files", "MOMO_files", "PChome_files", "Yahoo_files", "ETMall_files", "mo_store_plus_files", "coupang_files"]
    
    # 掃描平台資料夾
    for folder_name in platform_folders:
        folder_path = input_dir / folder_name
        if folder_path.exists() and folder_path.is_dir():
            print(f"[SCAN] 掃描平台資料夾：{folder_name}")
            folder_excel_files = []
            folder_compressed_files = []
            
            for file_path in folder_path.iterdir():
                if file_path.is_file():
                    file_ext = file_path.suffix.lower()
                    filename = file_path.name
                    
                    # 特殊處理：蝦皮平台只處理包含 "Order.all" 的檔案
                    if folder_name == "Shopee_files" and "Order.all" not in filename:
                        print(f"[SKIP] 蝦皮平台跳過非 Order.all 檔案：{filename}")
                        continue
                    
                    if file_ext in ['.xlsx', '.xls']:
                        folder_excel_files.append(file_path)
                    elif file_ext in ['.zip', '.rar']:
                        folder_compressed_files.append(file_path)
            
            print(f"[SCAN] 在 {folder_name} 中發現 {len(folder_excel_files)} 個 Excel 檔案，{len(folder_compressed_files)} 個壓縮檔案")
            excel_files.extend(folder_excel_files)
            
            # 處理該資料夾中的壓縮檔案
            if folder_compressed_files:
                print(f"[EXTRACT] 開始處理 {folder_name} 中的壓縮檔案...")
                extracted_files = process_platform_compressed_files(folder_compressed_files, output_dir, temp_dir, platform_index, folder_name, log_lines)
                excel_files.extend(extracted_files)
    
    # 掃描 input 根目錄中的檔案（向後相容）
    root_excel_files = []
    root_compressed_files = []
    for file_path in input_dir.iterdir():
        if file_path.is_file():
            file_ext = file_path.suffix.lower()
            if file_ext in ['.xlsx', '.xls']:
                root_excel_files.append(file_path)
            elif file_ext in ['.zip', '.rar']:
                root_compressed_files.append(file_path)
    
    if root_excel_files:
        print(f"[SCAN] 在 input 根目錄中發現 {len(root_excel_files)} 個 Excel 檔案")
        excel_files.extend(root_excel_files)
    
    if root_compressed_files:
        print(f"[SCAN] 在 input 根目錄中發現 {len(root_compressed_files)} 個壓縮檔案")
        # 處理根目錄中的壓縮檔案（使用所有平台密碼）
        print(f"[EXTRACT] 開始處理根目錄中的壓縮檔案...")
        extracted_files = process_root_compressed_files(root_compressed_files, output_dir, temp_dir, platform_index, log_lines)
        excel_files.extend(extracted_files)

    # 合併所有需要處理的 Excel 檔案
    all_excel_files = excel_files
    print(f"[FILES] 總計發現 {len(excel_files)} 個 Excel 檔案")

    # 處理每個 Excel 檔案
    for file_path in all_excel_files:
        filename = file_path.name
        print(f"\n[PROCESS] 正在處理：{filename}")

        # 根據檔案所在資料夾確定平台
        file_platform = None
        for folder_name in platform_folders:
            if folder_name in str(file_path):
                file_platform = folder_name.replace("_files", "")
                break
        
        if file_platform:
            print(f"[PLATFORM] 檔案來自平台資料夾：{file_platform}")
        else:
            print(f"[PLATFORM] 檔案來自根目錄，將嘗試所有平台")

        # 尋找匹配的帳號
        matched_account = None
        print(f"[MATCH] 正在匹配檔案：{filename}")
        
        # 特殊處理：MO_Store_Plus 檔案，嘗試所有有密碼的帳號
        if "MO_Store_Plus" in filename or file_platform == "mo_store_plus":
            print(f"   [MATCH] MO_Store_Plus 檔案，將嘗試所有有密碼的帳號")
            matched_account = "MO_Store_Plus"  # 標記為特殊處理
        else:
            # 根據平台篩選帳號
            target_accounts = excel_accounts
            if file_platform and file_platform in platform_index:
                # 只檢查該平台的帳號
                platform_accounts = {}
                for password, shop_info in platform_index[file_platform].items():
                    account = shop_info.get("shop_account", "")
                    if account:
                        platform_accounts[account] = shop_info
                target_accounts = platform_accounts
                print(f"   [MATCH] 限制在 {file_platform} 平台的 {len(target_accounts)} 個帳號中匹配")
            
            for account, account_info in target_accounts.items():
                name = account_info.get("shop_name", "")
                print(f"   [MATCH] 檢查帳號：{account}, 店家名稱：{name}")
                # 先嘗試匹配帳號
                if account in filename:
                    print(f"   [OK] 帳號匹配成功：{account}")
                    matched_account = account
                    break
                # 如果帳號匹配失敗，嘗試匹配店家名稱
                elif name and name in filename:
                    print(f"   [OK] 店家名稱匹配成功：{name}")
                    matched_account = account
                    break
                else:
                    print(f"   [FAIL] 無匹配：帳號 '{account}' 和店家名稱 '{name}' 都不在檔案名中")

        success = False
        
        if matched_account:
            if matched_account == "MO_Store_Plus":
                # 特殊處理：MO_Store_Plus 檔案，僅使用 mo_store_plus 平台密碼
                print(f"[WARN] MO_Store_Plus 檔案，僅使用 mo_store_plus 平台密碼破解：{filename}")
                success = try_platform_passwords(file_path, platform_index, "mo_store_plus", output_dir, log_lines)
            else:
                # 找到對應帳號，僅使用該平台的密碼
                account_info = excel_accounts[matched_account]
                shop_name = account_info.get("shop_name", "")
                shop_id = account_info.get("shop_id", "UNKNOWN")
                shop_account = account_info.get("shop_account", "UNKNOWN")
                
                # 根據檔案所在平台，僅使用該平台的密碼
                if file_platform and file_platform in platform_index:
                    print(f"[PLATFORM] 檔案來自 {file_platform} 平台，僅使用該平台密碼")
                    success = try_platform_passwords(file_path, platform_index, file_platform, output_dir, log_lines)
                    
                    if success:
                        # 成功處理，記錄到 processed_files
                        processed_files.append((filename, "已處理", shop_name, matched_account))
                else:
                    print(f"[WARN] 無法確定檔案平台，跳過處理：{filename}")
                    success = False

        # 如果還沒有成功，僅嘗試檔案所在平台的密碼
        if not success:
            if file_platform and file_platform in platform_index:
                print(f"[WARN] 嘗試使用 {file_platform} 平台密碼破解：{filename}")
                success = try_platform_passwords(file_path, platform_index, file_platform, output_dir, log_lines)
                
                if success:
                    # 成功處理，記錄到 processed_files
                    processed_files.append((filename, "已處理", "平台檔案", file_platform))
                else:
                    # 失敗，記錄到 failed_files
                    error_msg = f"[FAIL] {file_platform} 平台密碼無法破解：{filename}"
                    log_lines.append(error_msg)
                    failed_files.append((filename, error_msg))
                    print(error_msg)
            else:
                print(f"[WARN] 無法確定檔案平台，跳過處理：{filename}")
                error_msg = f"[FAIL] 無法確定檔案平台：{filename}"
                log_lines.append(error_msg)
                failed_files.append((filename, error_msg))

        # 如果所有密碼都無法破解
        if not success:
            error_msg = f"[FAIL] 所有密碼都無法破解：{filename}"
            log_lines.append(error_msg)
            failed_files.append((filename, error_msg))
            print(error_msg)

    # 寫入詳細日誌
    log_lines.append("\n" + "="*50)
    log_lines.append("[STAT] 處理統計")
    log_lines.append(f"總檔案數：{len(all_excel_files)}")
    log_lines.append(f"成功處理：{len(processed_files)}")
    log_lines.append(f"處理失敗：{len(failed_files)}")

    if processed_files:
        log_lines.append("\n[OK] 成功處理的檔案：")
        for original, new_name, name, account in processed_files:
            log_lines.append(f"  {original} → {new_name}")

    if failed_files:
        log_lines.append("\n[FAIL] 處理失敗的檔案：")
        for filename, error in failed_files:
            log_lines.append(f"  {filename}: {error}")

    # 寫入 log 檔案
    with log_path.open("w", encoding="utf-8") as log_file:
        log_file.write("\n".join(log_lines))

    # 輸出結果摘要
    print(f"\n" + "="*50)
    print(f"[STAT] 處理完成！")
    print(f"總檔案數：{len(all_excel_files)}")
    print(f"成功處理：{len(processed_files)}")
    print(f"處理失敗：{len(failed_files)}")
    print(f"[LOG] 詳細日誌：{log_path}")
    
    # 清理 temp 資料夾中的所有臨時檔案
    print(f"\n[CLEANUP] 開始清理 temp 資料夾...")
    temp_files_cleaned = 0
    temp_dirs_cleaned = 0
    
    # 清理 temp 資料夾中的所有檔案和資料夾
    if temp_dir.exists():
        for item in temp_dir.iterdir():
            try:
                if item.is_file():
                    item.unlink()
                    temp_files_cleaned += 1
                elif item.is_dir():
                    shutil.rmtree(item)
                    temp_dirs_cleaned += 1
            except Exception as e:
                print(f"[ERROR] 清理失敗：{item.name} - {e}")
    
    print(f"[CLEANUP] 總共清理了 {temp_files_cleaned} 個臨時檔案和 {temp_dirs_cleaned} 個臨時資料夾")

    if processed_files:
        print(f"\n[OK] 成功處理的檔案已重新命名並儲存至：{output_dir}")
        for original, new_name, name, account in processed_files:
            print(f"  {original} → {new_name}")

    if failed_files:
        print(f"\n[FAIL] 處理失敗的檔案：")
        for filename, error in failed_files:
            print(f"  {filename}: {error}")


if __name__ == "__main__":
    main()
    print("\n[OK] 執行完畢") 