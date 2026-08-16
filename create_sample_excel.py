"""
ساخت یک فایل اکسل نمونه (ips.sample.xlsx) با فرمت درست، برای شروع سریع.
اجرا: python create_sample_excel.py
"""

from pathlib import Path

import openpyxl

OUTPUT_PATH = Path(__file__).with_name("ips.sample.xlsx")

ROWS = [
    ("Web-Server-01", "10.0.0.10", ""),
    ("DB-Server-01", "10.0.0.20", ""),
    ("Branch-Isfahan-Net", "10.20.0.0/24", ""),
    ("Only-On-FW-Main", "10.30.0.5", "FW-Main"),
]


def main() -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "IPs"
    ws.append(["Name", "IP", "Firewall"])
    for row in ROWS:
        ws.append(row)

    for column_cells in ws.columns:
        max_length = max(len(str(cell.value)) for cell in column_cells if cell.value is not None)
        ws.column_dimensions[column_cells[0].column_letter].width = max_length + 4

    wb.save(OUTPUT_PATH)
    print(f"ساخته شد: {OUTPUT_PATH}")
    print('برای شروع، این فایل رو کپی کن به "ips.xlsx" (یا هر مسیری که توی config.json گذاشتی) و ویرایشش کن.')


if __name__ == "__main__":
    main()
