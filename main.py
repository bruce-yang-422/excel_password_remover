# main.py

from pathlib import Path
from scripts.remover import remove_password
from scripts.compression import extract_compressed_file, is_compressed_file
from scripts.utils import load_passwords
import datetime
import shutil
import sys
import os

print("🔧 DEBUG | sys.executable", sys.executable)
print("🔧 DEBUG | __file__", __file__)


def main():
    if getattr(sys, 'frozen', False):
        project_root = Path(sys.executable).parent.resolve()
    else:
        project_root = Path(__file__).parent.resolve()

    input_dir = project_root / "input"
    output_dir = project_root / "output"
    output_dir.mkdir(exist_ok=True)
    log_dir = project_root / "log"
    passwords_path = "passwords.yaml"  # 僅傳檔名，讓 load_passwords 自行判定路徑

    # 🔥 每次執行前清空 log 資料夾
    if log_dir.exists():
        shutil.rmtree(log_dir)
    log_dir.mkdir(exist_ok=True)

    # 定義 log 檔案名稱
    log_path = log_dir / f"execution_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    # 讀取 passwords.yaml
    data = load_passwords(passwords_path)
    accounts = {item["account"]: item["name"] for item in data["excel_files"]}
    
    # 建立壓縮檔案密碼對照表（根據密碼匹配）
    compressed_passwords = {}
    if "compressed_files" in data:
        for item in data["compressed_files"]:
            password = item.get("password")
            if password:
                compressed_passwords[password] = {
                    "name": item["name"],
                    "account": item.get("account")
                }

    log_lines = []
    processed_accounts = set()  # ✅ 初始化 processed_accounts

    for input_path in input_dir.iterdir():
        if not input_path.is_file():
            continue

        filename = input_path.name

        # 🔥 排除隱藏檔案（以 . 開頭）
        if filename.startswith("."):
            continue

        # 檢查是否為壓縮檔案
        if is_compressed_file(input_path):
            print(f"📦 發現壓縮檔案：{filename}")
            
            # 嘗試所有已知密碼來解壓縮
            extracted_files = []
            extract_dir = None
            matched_password_info = None
            
            for password, password_info in compressed_passwords.items():
                try:
                    # 建立解壓縮目標資料夾
                    extract_dir = output_dir / f"{filename}_{password_info['name']}_extracted"
                    
                    # 嘗試使用此密碼解壓縮
                    extracted_files = extract_compressed_file(input_path, extract_dir, password)
                    matched_password_info = password_info
                    print(f"✅ {filename} 使用密碼 {password} 解壓縮成功，匹配到：{password_info['name']}")
                    break
                    
                except Exception as e:
                    # 密碼錯誤，繼續嘗試下一個
                    continue
            
            # 如果所有密碼都失敗，嘗試無密碼解壓縮
            if not matched_password_info:
                try:
                    extract_dir = output_dir / f"{filename}_no_password_extracted"
                    extracted_files = extract_compressed_file(input_path, extract_dir)
                    print(f"✅ {filename} 無密碼解壓縮成功")
                    log_lines.append(f"✅ {filename} 無密碼解壓縮成功，共 {len(extracted_files)} 個檔案")
                except Exception as e:
                    log_lines.append(f"❌ {filename} 所有密碼嘗試失敗，解壓縮失敗：{e}")
                    print(f"❌ {filename} 所有密碼嘗試失敗，無法解壓縮")
                    continue
            else:
                log_lines.append(f"✅ {matched_password_info['name']} ({filename}) 已成功解壓縮，共 {len(extracted_files)} 個檔案")
                print(f"✅ {matched_password_info['name']} ({filename}) 解壓縮成功，檔案位於：{extract_dir}")
            
            # 處理解壓縮後的 Excel 檔案
            for extracted_file in extract_dir.rglob("*"):
                if extracted_file.is_file() and extracted_file.suffix.lower() in ['.xlsx', '.xls']:
                    excel_filename = extracted_file.name
                    
                    # 如果有匹配到壓縮檔案密碼，優先使用該 account
                    if matched_password_info:
                        compressed_account = matched_password_info.get("account")
                        if compressed_account and compressed_account in excel_filename:
                            account = compressed_account
                            account_name = matched_password_info["name"]
                            
                            # 取得 Excel 密碼
                            password_item = next(item for item in data["excel_files"] if item["account"] == account)
                            excel_password = password_item.get("password")
                            
                            # 建立 Excel 輸出路徑
                            excel_output_path = output_dir / excel_filename
                            
                            try:
                                if excel_password:
                                    # 若 yaml 有提供密碼，嘗試移除
                                    remove_password(extracted_file, excel_output_path, excel_password)
                                    log_lines.append(f"✅ 從 {matched_password_info['name']} 解壓縮的 {account_name} ({account}) Excel 檔案已成功移除密碼")
                                else:
                                    # yaml 無提供密碼，先嘗試用空密碼移除
                                    try:
                                        remove_password(extracted_file, excel_output_path, "")
                                        log_lines.append(f"✅ 從 {matched_password_info['name']} 解壓縮的 {account_name} ({account}) Excel 檔案無密碼，已直接複製")
                                    except Exception as e:
                                        if "Unencrypted document" in str(e):
                                            # 若檔案根本沒加密，直接複製
                                            shutil.copy2(extracted_file, excel_output_path)
                                            log_lines.append(f"✅ 從 {matched_password_info['name']} 解壓縮的 {account_name} ({account}) Excel 檔案本身無密碼，已直接複製")
                                        else:
                                            log_lines.append(f"❌ 從 {matched_password_info['name']} 解壓縮的 {account_name} ({account}) Excel 檔案有密碼但 yaml 未提供密碼")
                                            continue
                                
                                processed_accounts.add(account)
                                
                            except Exception as e:
                                if "Unencrypted document" in str(e):
                                    # 若檔案本身無密碼，直接複製
                                    shutil.copy2(extracted_file, excel_output_path)
                                    log_lines.append(f"✅ 從 {matched_password_info['name']} 解壓縮的 {account_name} ({account}) Excel 檔案無密碼，已直接複製")
                                else:
                                    log_lines.append(f"❌ 從 {matched_password_info['name']} 解壓縮的 {account_name} ({account}) Excel 檔案處理失敗：{e}")
                            continue
                    
                    # 檢查檔名是否包含任何其他 account
                    matched_accounts = [account for account in accounts if account in excel_filename]
                    
                    if matched_accounts:
                        account = matched_accounts[0]
                        account_name = accounts[account]
                        
                        # 取得 Excel 密碼
                        password_item = next(item for item in data["excel_files"] if item["account"] == account)
                        excel_password = password_item.get("password")
                        
                        # 建立 Excel 輸出路徑
                        excel_output_path = output_dir / excel_filename
                        
                        try:
                            if excel_password:
                                # 若 yaml 有提供密碼，嘗試移除
                                remove_password(extracted_file, excel_output_path, excel_password)
                                log_lines.append(f"✅ 從解壓縮檔案中找到的 {account_name} ({account}) Excel 檔案已成功移除密碼")
                            else:
                                # yaml 無提供密碼，先嘗試用空密碼移除
                                try:
                                    remove_password(extracted_file, excel_output_path, "")
                                    log_lines.append(f"✅ 從解壓縮檔案中找到的 {account_name} ({account}) Excel 檔案無密碼，已直接複製")
                                except Exception as e:
                                    if "Unencrypted document" in str(e):
                                        # 若檔案根本沒加密，直接複製
                                        shutil.copy2(extracted_file, excel_output_path)
                                        log_lines.append(f"✅ 從解壓縮檔案中找到的 {account_name} ({account}) Excel 檔案本身無密碼，已直接複製")
                                    else:
                                        log_lines.append(f"❌ 從解壓縮檔案中找到的 {account_name} ({account}) Excel 檔案有密碼但 yaml 未提供密碼")
                                        continue
                            
                            processed_accounts.add(account)
                            
                        except Exception as e:
                            if "Unencrypted document" in str(e):
                                # 若檔案本身無密碼，直接複製
                                shutil.copy2(extracted_file, excel_output_path)
                                log_lines.append(f"✅ 從解壓縮檔案中找到的 {account_name} ({account}) Excel 檔案無密碼，已直接複製")
                            else:
                                log_lines.append(f"❌ 從解壓縮檔案中找到的 {account_name} ({account}) Excel 檔案處理失敗：{e}")
            
            continue

        # 檢查檔名是否包含任何 account（原有的 Excel 檔案處理邏輯）
        matched_accounts = [account for account in accounts if account in filename]

        if not matched_accounts:
            log_lines.append(f"⚠️ 找不到對應 account，已跳過檔案：{filename}")
            continue

        account = matched_accounts[0]
        name = accounts[account]
        output_path = output_dir / filename

        # 取得 password，若不存在則設為 None
        password_item = next(item for item in data["excel_files"] if item["account"] == account)
        password = password_item.get("password")

        print(f"🔓 正在處理 {name} ({account})...")

        try:
            if password:
                # 若 yaml 有提供密碼，嘗試移除
                remove_password(input_path, output_path, password)
                log_lines.append(f"✅ {name} ({account}) 已成功移除密碼，輸出至 output 資料夾。")
            else:
                # yaml 無提供密碼，先嘗試用空密碼移除
                try:
                    remove_password(input_path, output_path, "")
                    log_lines.append(f"✅ {name} ({account}) 檔案無密碼，已直接複製到 output 資料夾。")
                except Exception as e:
                    if "Unencrypted document" in str(e):
                        # 若檔案根本沒加密，直接複製
                        shutil.copy2(input_path, output_path)
                        log_lines.append(f"✅ {name} ({account}) 檔案本身無密碼，已直接複製到 output 資料夾。")
                    else:
                        log_lines.append(f"❌ {name} ({account}) 檔案有密碼但 yaml 未提供密碼，無法處理。")
                        print(f"❌ [提醒] {name} ({account}) 檔案有密碼但 yaml 未提供，已跳過。")
                        continue

            processed_accounts.add(account)

            # ✅ 檢查輸出檔案是否存在且非 0 byte
            if not output_path.exists() or output_path.stat().st_size == 0:
                log_lines.append(f"❌ {name} ({account}) 輸出檔案異常，檔案不存在或大小為 0 byte。")
                print(f"❌ [錯誤] {name} ({account}) 輸出檔案異常，請確認。")

        except Exception as e:
            if "Unencrypted document" in str(e):
                # 若檔案本身無密碼，直接複製
                shutil.copy2(input_path, output_path)
                log_lines.append(f"✅ {name} ({account}) 檔案無密碼，已直接複製到 output 資料夾。")
            else:
                log_lines.append(f"❌ {name} ({account}) 處理失敗，錯誤原因：{e}")
                print(f"❌ [錯誤] {name} ({account}) 處理失敗，錯誤原因：{e}")

    # ⚠️ 不再檢查 yaml 中未找到對應檔案，因為沒意義

    # 輸出 log 檔案
    with log_path.open("w", encoding="utf-8") as log_file:
        log_file.write("\n".join(log_lines))

    print(f"\n📄 執行 log 已產生：{log_path}")


if __name__ == "__main__":
    main()
    input("\n✅ 執行完畢，請按 Enter 關閉視窗...")
