# scripts/compression.py

import zipfile
import rarfile
import os
from pathlib import Path
import shutil

def extract_zip(zip_path, extract_to, password=None):
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

def extract_rar(rar_path, extract_to, password=None):
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
    except rarfile.BadPasswordError:
        raise Exception(f"RAR 檔案密碼錯誤：{rar_path}")
    except Exception as e:
        raise Exception(f"解壓縮 RAR 檔案時發生錯誤：{e}")
    
    return extracted_files

def extract_compressed_file(file_path, extract_to, password=None):
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

def is_compressed_file(file_path):
    """
    檢查檔案是否為支援的壓縮檔案格式
    
    Args:
        file_path: 檔案路徑
    
    Returns:
        bool: 是否為支援的壓縮檔案
    """
    file_path = Path(file_path)
    return file_path.suffix.lower() in ['.zip', '.rar'] 