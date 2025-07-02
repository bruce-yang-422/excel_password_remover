# main.py

from pathlib import Path
from scripts.remover import remove_password
from scripts.utils import load_passwords
import datetime
import shutil
import sys

print("🔧 DEBUG | sys.executable", sys.executable)
print("🔧 DEBUG | __file__", __file__)


def main():
    if getattr(sys, 'frozen', False):
        project_root = Path(sys.executable).parent.resolve()
    else:
        project_root = Path(__file__).parent.resolve()

    input_dir = project_root / "input"
    output_dir = project_root / "output"
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

    log_lines = []
    processed_accounts = set() # ✅ 初始化 processed_accounts

    for input_path in input_dir.iterdir():
        if not input_path.is_file():
            continue

        filename = input_path.name

        # 🔥 排除隱藏檔案（以 . 開頭）
        if filename.startswith("."):
            continue

        # 檢查檔名是否包含任何 account
        matched_accounts = [account for account in accounts if account in filename]

        if not matched_accounts:
            log_lines.append(f"⚠️ 未在 passwords.yaml 中找到 account，檔案: {filename}")
            continue

        account = matched_accounts[0]
        name = accounts[account]
        output_path = output_dir / filename
        password = next(item["password"] for item in data["excel_files"] if item["account"] == account)

        print(f"🔓 正在處理 {name} ({account})...")

        try:
            remove_password(input_path, output_path, password)
            log_lines.append(f"✅ {name} ({account}) 處理成功，輸出至 {output_path}")
            processed_accounts.add(account)
        except Exception as e:
            log_lines.append(f"❌ {name} ({account}) 處理失敗: {e}")

    # 檢查 yaml 內是否有未被處理的 account
    for account, name in accounts.items():
        if account not in processed_accounts:
            log_lines.append(f"⚠️ {name} ({account}) 在 yaml 中，但未找到對應檔案")

    # 輸出 log 檔案
    with log_path.open("w", encoding="utf-8") as log_file:
        log_file.write("\n".join(log_lines))

    print(f"\n📄 執行 log 已產生：{log_path}")

if __name__ == "__main__":
    main()
    input("\n✅ 執行完畢，請按 Enter 關閉視窗...")
