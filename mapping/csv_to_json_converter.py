#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSV 轉 JSON 轉換腳本
將 mapping/A02_Shops_Master - Shops_Master.csv 轉換為 mapping/shops_master.json
並調整結構以平台為主要索引，密碼為次要索引
"""

import csv
import json
from pathlib import Path

def convert_csv_to_json():
    """將 CSV 檔案轉換為 JSON 格式"""
    
    # 檔案路徑
    csv_file = Path("mapping/A02_Shops_Master - Shops_Master.csv")
    json_file = Path("mapping/shops_master.json")
    
    # 檢查 CSV 檔案是否存在
    if not csv_file.exists():
        print(f"[ERROR] CSV 檔案不存在: {csv_file}")
        return False
    
    # 讀取 CSV 檔案
    shops_data = []
    
    with open(csv_file, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        
        for row in reader:
            # 跳過標題行（如果有的話）
            if row.get('platform') == '平台':
                continue
            
            # 提取需要的欄位
            shop_info = {
                "platform": row.get('platform', ''),
                "shop_id": row.get('shop_id', ''),
                "shop_account": row.get('shop_account', ''),
                "shop_name": row.get('shop_name', ''),
                "shop_status": row.get('shop_status', ''),
                "Universal Password": row.get('Universal Password', ''),
                "Report Download Password": row.get('Report Download Password', '')
            }
            
            # 只保留有密碼的記錄
            if shop_info["Universal Password"] or shop_info["Report Download Password"]:
                shops_data.append(shop_info)
    
    # 建立以平台為主要索引，密碼為次要索引的結構
    platform_index = {}
    
    for shop in shops_data:
        platform = shop.get("platform", "")
        universal_pwd = shop.get("Universal Password", "")
        report_pwd = shop.get("Report Download Password", "")
        
        # 如果平台不存在，先建立平台索引
        if platform not in platform_index:
            platform_index[platform] = {}
        
        # 為 Universal Password 建立索引
        if universal_pwd and universal_pwd != "無":
            platform_index[platform][universal_pwd] = {
                "platform": shop["platform"],
                "shop_id": shop["shop_id"],
                "shop_account": shop["shop_account"],
                "shop_name": shop["shop_name"],
                "shop_status": shop["shop_status"],
                "Universal Password": universal_pwd,
                "Report Download Password": report_pwd
            }
        
        # 為 Report Download Password 建立索引
        if report_pwd and report_pwd != "無":
            platform_index[platform][report_pwd] = {
                "platform": shop["platform"],
                "shop_id": shop["shop_id"],
                "shop_account": shop["shop_account"],
                "shop_name": shop["shop_name"],
                "shop_status": shop["shop_status"],
                "Universal Password": universal_pwd,
                "Report Download Password": report_pwd
            }
    
    # 建立最終的 JSON 結構
    result = {
        "platform_index": platform_index,
        "shops": shops_data
    }
    
    # 寫入 JSON 檔案
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"[SUCCESS] 轉換完成！")
    print(f"[STAT] 處理了 {len(shops_data)} 個商店記錄")
    print(f"[PLATFORM] 建立了 {len(platform_index)} 個平台索引")
    
    # 計算總密碼數量
    total_passwords = sum(len(passwords) for passwords in platform_index.values())
    print(f"[PASSWORD] 總共建立了 {total_passwords} 個密碼索引")
    print(f"[FILE] 輸出檔案: {json_file}")
    
    return True

def show_password_index():
    """顯示平台和密碼索引的內容"""
    
    json_file = Path("mapping/shops_master.json")
    
    if not json_file.exists():
        print(f"[ERROR] JSON 檔案不存在: {json_file}")
        return
    
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    platform_index = data.get("platform_index", {})
    
    print(f"\n[PLATFORM] 平台索引 (共 {len(platform_index)} 個平台):")
    print("=" * 80)
    
    total_passwords = 0
    for platform, passwords in platform_index.items():
        print(f"\n平台: {platform}")
        print(f"密碼數量: {len(passwords)}")
        print("-" * 40)
        
        for i, (password, info) in enumerate(passwords.items(), 1):
            print(f"  {i:2d}. 密碼: {password}")
            print(f"      商店ID: {info['shop_id']}")
            print(f"      帳號: {info['shop_account']}")
            print(f"      名稱: {info['shop_name']}")
            print(f"      狀態: {info['shop_status']}")
            print()
        
        total_passwords += len(passwords)
    
    print(f"\n[SUMMARY] 總計: {len(platform_index)} 個平台, {total_passwords} 個密碼")

if __name__ == "__main__":
    print("[START] 開始轉換 CSV 到 JSON...")
    
    if convert_csv_to_json():
        show_password_index()
    else:
        print("[ERROR] 轉換失敗")
