import shutil
import os

src = "referensi/SWARA_MAIRCA_DETAIL_RUMUS.xlsx"
dst = "output/SWARA_MAIRCA_DETAIL_RUMUS.xlsx"

try:
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)
    print(f"Successfully copied Excel file to: {dst}")
except Exception as e:
    print(f"Error copying file: {e}")
