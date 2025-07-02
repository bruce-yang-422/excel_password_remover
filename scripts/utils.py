# scripts/utils.py

import yaml
from pathlib import Path

def load_passwords(yaml_path):
    """
    讀取 passwords.yaml
    支援 Path 或 str 輸入
    """
    yaml_path = Path(yaml_path)
    with yaml_path.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f)
