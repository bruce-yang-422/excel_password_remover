# main.py

import os
from scripts.remover import remove_password
from scripts.utils import load_passwords, build_output_path

def main():
    # 讀取 passwords.yaml
    data = load_passwords("passwords.yaml")
    
    for item in data["excel_files"]:
        name = item["name"]
        account = item["account"]
        password = item["password"]

        # 尋找 input/ 中符合帳號關鍵字的檔案
        input_dir = "input"
        matched_files = [f for f in os.listdir(input_dir) if account in f]
        
        if not matched_files:
            print(f"⚠️ 找不到 {account} 對應檔案")
            continue

        input_path = os.path.join(input_dir, matched_files[0])
        output_path = build_output_path(input_path)
        
        print(f"🔓 移除 {name} ({account}) 的密碼中...")

        try:
            remove_password(input_path, output_path, password)
            print(f"✅ 已輸出至 {output_path}")
        except Exception as e:
            print(f"❌ {name} ({account}) 移除失敗: {e}")

if __name__ == "__main__":
    main()
