# scripts/utils.py

import yaml
from pathlib import Path
import sys

def load_passwords(yaml_filename):
    """
    讀取 passwords.yaml
    exe 模式與 py 模式皆從執行檔所在資料夾讀取
    """
    if getattr(sys, 'frozen', False):
        # exe mode
        base_path = Path(sys.executable).parent
    else:
        # python mode
        base_path = Path(__file__).parent.parent

    yaml_path = base_path / yaml_filename

    # 🔧 debug print
    print("🔧 DEBUG | load_passwords")
    print("sys.executable:", sys.executable)
    print("base_path:", base_path)
    print("yaml_path:", yaml_path)

    if not yaml_path.exists():
        raise FileNotFoundError(f"找不到 passwords.yaml: {yaml_path}")

    with yaml_path.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f)
