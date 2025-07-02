# scripts/utils.py

import yaml
from pathlib import Path

def load_passwords(yaml_path):
    """
    讀取 passwords.yaml
    """
    with open(yaml_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def build_output_path(input_path):
    """
    回傳與 input 檔案相同檔名，但路徑為 output/
    """
    input_file = Path(input_path).name
    return str(Path("output") / input_file)
