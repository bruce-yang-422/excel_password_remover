# scripts/batch_password_remover.py
# 批次密碼移除工具 - 使用 passwords.yaml 密碼批次破解並移除 Excel 檔案密碼

import sys
from pathlib import Path
try:
    from .remover import remove_password
    from .utils import load_passwords
except ImportError:
    # 當直接執行時使用絕對導入
    import sys
    from pathlib import Path
    sys.path.append(str(Path(__file__).parent))
    from remover import remove_password
    from utils import load_passwords
import datetime
import shutil
import tempfile

def test_password(file_path, password):
    """測試密碼是否正確，使用臨時檔案"""
    try:
        # 建立臨時檔案來測試
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
            temp_path = temp_file.name
        
        # 嘗試移除密碼到臨時檔案
        remove_password(file_path, temp_path, password)
        
        # 如果成功，刪除臨時檔案
        Path(temp_path).unlink()
        return True
        
    except Exception:
        # 如果失敗，清理臨時檔案（如果存在）
        try:
            Path(temp_path).unlink()
        except:
            pass
        return False

def main():
    """主程式：批次處理 Excel 檔案密碼移除"""
    
    # 取得專案根目錄
    if getattr(sys, 'frozen', False):
        project_root = Path(sys.executable).parent.resolve()
    else:
        project_root = Path(__file__).parent.parent.resolve()

    input_dir = project_root / "input"
    output_dir = project_root / "output"
    output_dir.mkdir(exist_ok=True)
    log_dir = project_root / "log"
    passwords_path = "passwords.yaml"

    # 清空並建立 log 資料夾
    if log_dir.exists():
        shutil.rmtree(log_dir)
    log_dir.mkdir(exist_ok=True)

    # 建立 log 檔案
    log_path = log_dir / f"batch_removal_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    # 讀取密碼設定
    try:
        data = load_passwords(passwords_path)
        excel_accounts = {item["account"]: item for item in data["excel_files"]}
        print(f"✅ 成功載入 {len(excel_accounts)} 個帳號設定")
    except Exception as e:
        print(f"❌ 載入 passwords.yaml 失敗：{e}")
        return

    log_lines = []
    processed_files = []
    failed_files = []

    # 掃描 input 資料夾中的 Excel 檔案（只掃描當前目錄，不遞迴）
    excel_files = []
    for file_path in input_dir.iterdir():
        if file_path.is_file() and file_path.suffix.lower() in ['.xlsx', '.xls']:
            excel_files.append(file_path)

    print(f"📁 發現 {len(excel_files)} 個 Excel 檔案")

    # 處理每個 Excel 檔案
    for file_path in excel_files:
        filename = file_path.name
        print(f"\n🔓 正在處理：{filename}")

        # 尋找匹配的帳號
        matched_account = None
        for account, account_info in excel_accounts.items():
            if account in filename:
                matched_account = account
                break

        success = False
        
        if matched_account:
            # 找到對應帳號，檢查密碼是否有效
            account_info = excel_accounts[matched_account]
            name = account_info["name"]
            password = account_info.get("password")

            # 如果密碼是無效的通用密碼，跳過直接嘗試所有密碼
            if password == "password:":
                print(f"⚠️ 通用帳號密碼無效，直接嘗試所有密碼破解：{filename}")
            else:
                # 測試該帳號的密碼
                if test_password(file_path, password):
                    # 密碼正確，建立檔案
                    file_ext = file_path.suffix.lower()
                    # 清理檔案名稱，移除或替換特殊字符
                    safe_name = name.replace('.', '_').replace(' ', '_')
                    safe_account = matched_account.replace('.', '_').replace(' ', '_')
                    new_filename = f"{safe_name}_{safe_account}_{filename}"
                    
                    # 確保副檔名正確
                    if not new_filename.lower().endswith(file_ext):
                        new_filename = new_filename.replace(file_ext, '') + file_ext

                    output_path = output_dir / new_filename
                    
                    try:
                        remove_password(file_path, output_path, password)
                        success_msg = f"✅ {name} ({matched_account}) - 成功移除密碼：{new_filename}"
                        log_lines.append(success_msg)
                        processed_files.append((filename, new_filename, name, matched_account))
                        print(success_msg)
                        success = True
                    except Exception as e:
                        error_msg = f"❌ {name} ({matched_account}) - 處理失敗：{e}"
                        log_lines.append(error_msg)
                        failed_files.append((filename, error_msg))
                        print(error_msg)
                        # 不標記為已處理，讓程式繼續嘗試其他密碼

        # 如果還沒有成功，嘗試所有密碼
        if not success:
            print(f"⚠️ 嘗試使用所有密碼破解：{filename}")
            
            for account, account_info in excel_accounts.items():
                password = account_info.get("password")
                if password:  # 只跳過空密碼
                    # 測試密碼是否正確
                    if test_password(file_path, password):
                        # 密碼正確，建立檔案
                        # 如果有匹配的帳號，使用匹配的帳號名稱；否則使用密碼對應的帳號
                        if matched_account:
                            matched_name = excel_accounts[matched_account]["name"]
                            name = matched_name
                            account_for_naming = matched_account
                        else:
                            name = account_info["name"]
                            account_for_naming = account
                        
                        file_ext = file_path.suffix.lower()
                        # 清理檔案名稱，移除或替換特殊字符
                        safe_name = name.replace('.', '_').replace(' ', '_')
                        safe_account = account_for_naming.replace('.', '_').replace(' ', '_')
                        new_filename = f"{safe_name}_{safe_account}_{filename}"
                        
                        # 確保副檔名正確
                        if not new_filename.lower().endswith(file_ext):
                            new_filename = new_filename.replace(file_ext, '') + file_ext

                        output_path = output_dir / new_filename
                        
                        try:
                            remove_password(file_path, output_path, password)
                            success_msg = f"✅ 使用 {account_info['name']} ({account}) 密碼成功破解：{new_filename}"
                            log_lines.append(success_msg)
                            processed_files.append((filename, new_filename, name, account_for_naming))
                            print(success_msg)
                            success = True
                            break
                        except Exception as e:
                            error_msg = f"❌ 使用 {account_info['name']} ({account}) 密碼處理失敗：{e}"
                            log_lines.append(error_msg)
                            failed_files.append((filename, error_msg))
                            print(error_msg)
                            success = True  # 標記為已處理，避免重複嘗試
                            break

        # 如果所有密碼都無法破解
        if not success:
            error_msg = f"❌ 所有密碼都無法破解：{filename}"
            log_lines.append(error_msg)
            failed_files.append((filename, error_msg))
            print(error_msg)

    # 寫入詳細日誌
    log_lines.append("\n" + "="*50)
    log_lines.append("📊 處理統計")
    log_lines.append(f"總檔案數：{len(excel_files)}")
    log_lines.append(f"成功處理：{len(processed_files)}")
    log_lines.append(f"處理失敗：{len(failed_files)}")

    if processed_files:
        log_lines.append("\n✅ 成功處理的檔案：")
        for original, new_name, name, account in processed_files:
            log_lines.append(f"  {original} → {new_name}")

    if failed_files:
        log_lines.append("\n❌ 處理失敗的檔案：")
        for filename, error in failed_files:
            log_lines.append(f"  {filename}: {error}")

    # 寫入 log 檔案
    with log_path.open("w", encoding="utf-8") as log_file:
        log_file.write("\n".join(log_lines))

    # 輸出結果摘要
    print(f"\n" + "="*50)
    print(f"📊 處理完成！")
    print(f"總檔案數：{len(excel_files)}")
    print(f"成功處理：{len(processed_files)}")
    print(f"處理失敗：{len(failed_files)}")
    print(f"📄 詳細日誌：{log_path}")

    if processed_files:
        print(f"\n✅ 成功處理的檔案已重新命名並儲存至：{output_dir}")
        for original, new_name, name, account in processed_files:
            print(f"  {original} → {new_name}")

    if failed_files:
        print(f"\n❌ 處理失敗的檔案：")
        for filename, error in failed_files:
            print(f"  {filename}: {error}")


if __name__ == "__main__":
    main()
    input("\n✅ 執行完畢，請按 Enter 關閉視窗...") 