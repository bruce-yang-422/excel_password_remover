#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
File: TreeMaker.py
用途: 增強版目錄樹生成器
說明: 簡單的專案目錄樹結構生成工具，建立專案資料夾結構的視覺化表示
     將目錄樹結構儲存到 tree.txt，適用於文件記錄和專案結構概覽
     自動掃描專案資料夾並生成層級式目錄結構
重要提醒: 輸出檔案為根目錄的 tree.txt
Authors: 楊翔志 & AI Collective
Studio: tranquility-base
版本: 1.1 (2025-07-14)
"""
import os
import argparse
from pathlib import Path

EXCLUDE_DIRS = {
    '.git', '__pycache__', '.mypy_cache', '.pytest_cache', '.DS_Store', 
    'dist', 'build', '.coverage', '.tox', '.eggs'
}

EXCLUDE_FILES = {
    '.gitignore', '.gitattributes', '.DS_Store', 'Thumbs.db',
    '*.pyc', '*.pyo', '*.pyd', '__pycache__'
}

def should_exclude_dir(dir_name: str) -> bool:
    """檢查目錄是否應該被排除"""
    return dir_name in EXCLUDE_DIRS or dir_name.startswith('.')

def should_exclude_file(file_name: str) -> bool:
    """檢查檔案是否應該被排除"""
    return (file_name in EXCLUDE_FILES or 
            file_name.startswith('.') or 
            file_name.endswith(('.pyc', '.pyo', '.pyd')))

def generate_tree(directory: Path, prefix: str = "", is_last: bool = True, max_depth: int = None, current_depth: int = 0) -> str:
    """
    生成目錄樹結構
    
    Args:
        directory: 目錄路徑
        prefix: 前綴字串
        is_last: 是否為最後一個項目
        max_depth: 最大深度限制
        current_depth: 當前深度
    
    Returns:
        目錄樹字串
    """
    if max_depth is not None and current_depth >= max_depth:
        return ""
    
    tree_lines = []
    
    # 獲取目錄中的所有項目，排序並分組
    try:
        items = list(directory.iterdir())
        items.sort(key=lambda x: (x.is_file(), x.name.lower()))
    except PermissionError:
        return f"{prefix}└── [權限不足]\n"
    
    # 過濾掉需要排除的項目
    filtered_items = []
    for item in items:
        if item.is_dir() and not should_exclude_dir(item.name):
            filtered_items.append(item)
        elif item.is_file() and not should_exclude_file(item.name):
            filtered_items.append(item)
    
    for i, item in enumerate(filtered_items):
        is_last_item = i == len(filtered_items) - 1
        
        # 選擇適當的符號
        if is_last_item:
            current_prefix = "└── "
            next_prefix = prefix + "    "
        else:
            current_prefix = "├── "
            next_prefix = prefix + "│   "
        
        # 添加項目名稱
        if item.is_dir():
            tree_lines.append(f"{prefix}{current_prefix}{item.name}/")
            # 遞歸處理子目錄
            subtree = generate_tree(item, next_prefix, is_last_item, max_depth, current_depth + 1)
            tree_lines.append(subtree)
        else:
            tree_lines.append(f"{prefix}{current_prefix}{item.name}")
    
    return "\n".join(tree_lines) + "\n" if tree_lines else ""

def main():
    """主函數"""
    parser = argparse.ArgumentParser(description="生成專案目錄樹結構")
    parser.add_argument("--path", "-p", type=str, default=".", help="要掃描的目錄路徑 (預設: 當前目錄)")
    parser.add_argument("--output", "-o", type=str, default="tree.txt", help="輸出檔案名稱 (預設: tree.txt)")
    parser.add_argument("--max-depth", "-d", type=int, help="最大掃描深度")
    parser.add_argument("--exclude", "-e", nargs="*", help="額外排除的目錄或檔案")
    
    args = parser.parse_args()
    
    # 處理路徑
    target_dir = Path(args.path).resolve()
    if not target_dir.exists():
        print(f"❌ 錯誤：目錄 '{target_dir}' 不存在")
        return
    
    if not target_dir.is_dir():
        print(f"❌ 錯誤：'{target_dir}' 不是一個目錄")
        return
    
    # 處理額外排除項目
    if args.exclude:
        for item in args.exclude:
            if item.startswith('.'):
                EXCLUDE_DIRS.add(item)
            else:
                EXCLUDE_FILES.add(item)
    
    print(f"🌳 正在生成目錄樹...")
    print(f"📁 掃描目錄：{target_dir}")
    print(f"📄 輸出檔案：{args.output}")
    if args.max_depth:
        print(f"📏 最大深度：{args.max_depth}")
    
    # 生成目錄樹
    tree_content = f"專案目錄樹結構\n"
    tree_content += f"掃描路徑：{target_dir}\n"
    tree_content += f"生成時間：{os.popen('date /t & time /t').read().strip()}\n"
    tree_content += "=" * 50 + "\n\n"
    tree_content += generate_tree(target_dir, max_depth=args.max_depth)
    
    # 寫入檔案
    try:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(tree_content)
        print(f"✅ 目錄樹已成功生成：{args.output}")
        print(f"📊 檔案大小：{os.path.getsize(args.output)} bytes")
    except Exception as e:
        print(f"❌ 寫入檔案失敗：{e}")

if __name__ == "__main__":
    main()
