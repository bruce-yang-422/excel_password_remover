# scripts/remover.py

import msoffcrypto

def remove_password(input_path, output_path, password):
    """
    使用 msoffcrypto-tool 解開 Excel 開啟密碼，另存為 output_path
    """
    with open(input_path, "rb") as f_in:
        office_file = msoffcrypto.OfficeFile(f_in)
        office_file.load_key(password=password)
        with open(output_path, "wb") as f_out:
            office_file.decrypt(f_out)
