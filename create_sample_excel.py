"""
ساخت یک فایل اکسل نمونه (ips.sample.xlsx) با فرمت درست، برای شروع سریع.
اجرا: python create_sample_excel.py
"""

from pathlib import Path

import openpyxl

OUTPUT_PATH = Path(__file__).with_name("ips.sample.xlsx")

ADDRESS_ROWS = [
    ("Web-Server-01", "10.0.0.10", ""),
    ("DB-Server-01", "10.0.0.20", ""),
    ("Branch-Isfahan-Net", "10.20.0.0/24", ""),
    ("Only-On-FW-Main", "10.30.0.5", "FW-Main"),
]

TUNNEL_ROWS = [
    ("VPN-To-Branch-Isfahan", "203.0.113.10", "ChangeThisPreSharedKey!", "10.0.0.0/24", "10.20.0.0/24", "wan1", ""),
]


def _autosize(ws) -> None:
    for column_cells in ws.columns:
        max_length = max(len(str(cell.value)) for cell in column_cells if cell.value is not None)
        ws.column_dimensions[column_cells[0].column_letter].width = max_length + 4


def main() -> None:
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "IPs"
    ws.append(["Name", "IP", "Firewall"])
    for row in ADDRESS_ROWS:
        ws.append(row)
    _autosize(ws)

    tunnels_ws = wb.create_sheet("Tunnels")
    tunnels_ws.append(["Name", "RemoteGateway", "PSK", "LocalSubnet", "RemoteSubnet", "Interface", "Firewall"])
    for row in TUNNEL_ROWS:
        tunnels_ws.append(row)
    _autosize(tunnels_ws)

    wb.save(OUTPUT_PATH)
    print(f"ساخته شد: {OUTPUT_PATH}")
    print('برای شروع، این فایل رو کپی کن به "ips.xlsx" (یا هر مسیری که توی config.json گذاشتی) و ویرایشش کن.')


if __name__ == "__main__":
    main()
