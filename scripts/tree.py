#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
File: tree.py
用途: 目錄樹生成器
說明: 簡單的專案目錄樹結構生成工具，建立專案資料夾結構的視覺化表示
     將目錄樹結構儲存到 tree.txt，適用於文件記錄和專案結構概覽
     自動掃描專案資料夾並生成層級式目錄結構
重要提醒: 輸出檔案為根目錄的 tree.txt
作者: 多元寵物用品ETL團隊
版本: 1.0 (2024-06-24)
"""
import os

EXCLUDE_DIRS = {'.git', '__pycache__', '.venv', '.idea', '.vscode', 'env', 'venv', 'node_modules', '.mypy_cache', '.pytest_cache'}

def print_tree(root, prefix="", file=None):
    entries = [e for e in os.listdir(root) if e not in EXCLUDE_DIRS]
    entries.sort()
    for idx, entry in enumerate(entries):
        path = os.path.join(root, entry)
        connector = "└── " if idx == len(entries) - 1 else "├── "
        file.write(prefix + connector + entry + "\n")
        if os.path.isdir(path):
            extension = "    " if idx == len(entries) - 1 else "│   "
            print_tree(path, prefix + extension, file)

if __name__ == "__main__":
    root_path = os.getcwd()
    folder_name = os.path.basename(root_path)
    if not folder_name:  # 處理根目錄情況
        folder_name = root_path
    with open("tree.txt", "w", encoding="utf-8") as f:
        f.write(folder_name + "\n")  # 顯示實際目錄名稱
        print_tree(root_path, file=f)
    print(f"樹狀圖已輸出到 tree.txt（根目錄為：{folder_name}）")
