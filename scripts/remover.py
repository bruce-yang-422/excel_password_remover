# scripts/remover.py

import msoffcrypto
import shutil

def remove_password(input_path, output_path, password):
    """
    使用 msoffcrypto-tool 解開 Excel 開啟密碼，另存為 output_path
    若檔案未加密，直接複製
    """
    with open(input_path, "rb") as f_in:
        office_file = msoffcrypto.OfficeFile(f_in)
        try:
            office_file.load_key(password=password)
            with open(output_path, "wb") as f_out:
                office_file.decrypt(f_out)
        except msoffcrypto.exceptions.FileFormatError as e:
            if "Unencrypted document" in str(e):
                # 檔案未加密，直接複製
                shutil.copyfile(input_path, output_path)
            else:
                raise
