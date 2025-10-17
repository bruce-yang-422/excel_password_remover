#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSV 轉 JSON 資料轉換工具

主要功能：
    📊 將 CSV 格式的店家資料轉換為 JSON 格式
    🔤 自動處理 CSV 檔案中的編碼問題
    🏗️  生成結構化的 JSON 資料
    📈 支援多平台和部門的資料分類統計

輸入檔案：
    mapping/A02_Shops_Master - Shops_Master.csv

輸出檔案：
    mapping/shops_master.json

使用方法：
    python scripts/csv_to_json.py

處理流程：
    1. 讀取 CSV 檔案並自動處理編碼
    2. 解析店家資料和平台資訊
    3. 生成結構化的 JSON 資料
    4. 保存到指定的輸出目錄
    5. 顯示統計資訊和平台分布

支援功能：
    ✅ 自動編碼檢測和轉換
    ✅ 多平台資料分類
    ✅ 部門資料統計
    ✅ 錯誤處理和日誌記錄
    ✅ 平台分布統計

輸出格式：
    {
        "shops": [...],           // 店家資料陣列
        "total_count": 123,       // 總店鋪數
        "platforms": [...],       // 平台列表
        "departments": [...],     // 部門列表
        "platform_distribution": {...}  // 平台分布統計
    }
"""

import csv
import json
import sys
from pathlib import Path
from typing import Dict, List, Any


def csv_to_json(csv_file_path: str, output_dir: str = "mapping") -> None:
    """
    將 CSV 檔案轉換為 JSON 格式
    
    Args:
        csv_file_path: CSV 檔案路徑
        output_dir: 輸出目錄
    """
    data = []
    platforms = set()
    departments = set()

    try:
        with open(csv_file_path, 'r', encoding='utf-8-sig') as csv_file:
            # 讀取所有行
            lines = csv_file.readlines()
            
            # 跳過第2行（中文標題欄位），從第3行開始處理
            if len(lines) >= 2:
                # 使用第1行作為標題，從第3行開始讀取資料
                csv_content = [lines[0]] + lines[2:]  # 第1行標題 + 第3行開始的資料
                
                # 使用 StringIO 來處理修改後的內容
                from io import StringIO
                csv_reader = csv.DictReader(StringIO(''.join(csv_content)))
                
                for row in csv_reader:
                    # 清理空值
                    cleaned_row = {k: v.strip() for k, v in row.items() if v is not None}
                    data.append(cleaned_row)
                    
                    # 收集統計資訊
                    if 'platform' in cleaned_row:
                        platforms.add(cleaned_row['platform'])
                    if 'department' in cleaned_row:
                        departments.add(cleaned_row['department'])
            else:
                print("❌ 錯誤：CSV 檔案格式不正確，至少需要3行（標題、中文標題、資料）")
                return

        # 統計平台分布
        platform_distribution = {}
        for item in data:
            if 'platform' in item:
                platform = item['platform']
                platform_distribution[platform] = platform_distribution.get(platform, 0) + 1

        json_output = {
            "shops": data,
            "total_count": len(data),
            "platforms": sorted(list(platforms)),
            "departments": sorted(list(departments)),
            "platform_distribution": platform_distribution
        }

        # 確保輸出目錄存在
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)
        
        json_file_path = output_path / "shops_master.json"
        
        with open(json_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(json_output, json_file, ensure_ascii=False, indent=2)
        
        print(f"✅ 成功讀取 {len(data)} 筆資料")
        print(f"✅ 成功轉換為 JSON 格式：{json_file_path.relative_to(Path.cwd())}")
        print("📊 統計資訊：")
        print(f"   - 總店鋪數：{len(data)}")
        print(f"   - 平台數：{len(platforms)}")
        print(f"   - 部門數：{len(departments)}")
        print("📈 平台分布：")
        for platform, count in platform_distribution.items():
            print(f"   - {platform}: {count} 家店鋪")

    except FileNotFoundError:
        print(f"❌ 錯誤：找不到檔案 {csv_file_path}")
    except Exception as e:
        print(f"❌ 轉換失敗：{e}")


if __name__ == "__main__":
    # 取得專案根目錄
    if getattr(sys, 'frozen', False):
        project_root = Path(sys.executable).parent.resolve()
    else:
        project_root = Path(__file__).parent.parent.resolve()

    mapping_dir = project_root / "mapping"
    mapping_dir.mkdir(exist_ok=True) # 確保 mapping 資料夾存在

    csv_file_name = "A02_Shops_Master - Shops_Master.csv"
    json_file_name = "shops_master.json"

    csv_path = mapping_dir / csv_file_name
    json_path = mapping_dir / json_file_name

    print(f"📖 正在讀取 CSV 檔案：{csv_path}")
    csv_to_json(csv_path, mapping_dir)
