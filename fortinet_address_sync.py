"""
نظارت روی یک فایل اکسل (Name/IP/Firewall) و همگام‌سازی خودکار address objectهای
یک یا چند فایروال Fortinet از طریق SSH — هر بار که فایل اکسل آپدیت می‌شود:
  - ردیف‌های جدید  -> ساخته می‌شوند
  - ردیف‌هایی که IP‌شان از قبل هست -> رد می‌شوند (بدون خطا)
  - ردیف‌هایی که از اکسل پاک شده‌اند و قبلاً توسط همین اسکریپت ساخته شده بودند
    -> از فایروال هم حذف می‌شوند
  - بین هر دستور یک تاخیر (پیش‌فرض ۳ ثانیه) گذاشته می‌شود تا فایروال هنگ نکند
  - بعد از هر اجرا یک گزارش اکسل ذخیره می‌شود

اجرا (به‌صورت دستی/کنسول):
    python fortinet_address_sync.py

برای اجرای دائمی به‌صورت سرویس ویندوز از fortinet_sync_service.py استفاده کن.

قبل از اجرا:
    1. pip install -r requirements.txt
    2. کپی از config.json.example به config.json و پر کردن اطلاعات واقعی
"""

from __future__ import annotations

import ipaddress
import json
import logging
import sys
import time
from datetime import datetime
from pathlib import Path

import openpyxl
from colorama import Fore, Style, init as colorama_init
from netmiko import ConnectHandler
from netmiko.exceptions import NetmikoAuthenticationException, NetmikoTimeoutException

CONFIG_PATH = Path(__file__).with_name("config.json")
REPORT_HEADERS = ["Time", "Firewall", "Name", "IP", "Action", "Message"]


# ---------------------------------------------------------------------------
# تنظیمات و لاگ
# ---------------------------------------------------------------------------

def resolve_path(value: str) -> str:
    """مسیرهای نسبی رو به پوشه‌ی خود پروژه (کنار config.json) وصل می‌کنه، نه به
    working directory فعلی — چون وقتی به‌صورت سرویس ویندوز اجرا میشه، CWD قابل اعتماد نیست."""
    p = Path(value)
    return str(p if p.is_absolute() else (CONFIG_PATH.parent / p))


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        print(
            f"فایل config.json پیدا نشد. اول از config.json.example یک کپی به اسم "
            f"config.json بساز و اطلاعات فایروال(ها)/اکسل رو داخلش پر کن.\n"
            f"مسیر مورد انتظار: {CONFIG_PATH}"
        )
        sys.exit(1)
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    if not cfg.get("firewalls"):
        print("توی config.json حداقل یک فایروال باید توی لیست 'firewalls' تعریف بشه.")
        sys.exit(1)

    cfg["excel"]["path"] = resolve_path(cfg["excel"]["path"])
    cfg["state_file"] = resolve_path(cfg.get("state_file", "sync_state.json"))
    cfg["report_file"] = resolve_path(cfg.get("report_file", "reports/sync_report.xlsx"))
    cfg["log_file"] = resolve_path(cfg.get("log_file", "fortinet_sync.log"))
    return cfg


class ColorFormatter(logging.Formatter):
    LEVEL_COLORS = {
        logging.DEBUG: Fore.CYAN,
        logging.WARNING: Fore.YELLOW,
        logging.ERROR: Fore.RED,
        logging.CRITICAL: Fore.RED + Style.BRIGHT,
    }
    KEYWORD_COLORS = [
        ("اضافه شد", Fore.GREEN),
        ("آپدیت شد", Fore.BLUE),
        ("حذف شد", Fore.MAGENTA),
        ("رد شد", Fore.YELLOW),
        ("خطا", Fore.RED),
    ]

    def format(self, record: logging.LogRecord) -> str:
        message = super().format(record)
        color = self.LEVEL_COLORS.get(record.levelno, "")
        text = record.getMessage()
        for keyword, kw_color in self.KEYWORD_COLORS:
            if keyword in text:
                color = kw_color
                break
        return f"{color}{message}{Style.RESET_ALL}" if color else message


def setup_logging(log_file: str, colorize: bool = True) -> None:
    if colorize:
        colorama_init()

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(
        ColorFormatter("%(asctime)s [%(levelname)s] %(message)s")
        if colorize
        else logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    )

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.handlers.clear()
    root.addHandler(console_handler)
    root.addHandler(file_handler)


def print_banner() -> None:
    print(f"{Fore.CYAN}{Style.BRIGHT}")
    print("=" * 52)
    print("  Fortinet Excel Address Sync")
    print("=" * 52)
    print(Style.RESET_ALL)


def print_summary(fw_name: str, stats: dict) -> None:
    print(f"{Fore.CYAN}{'-' * 40}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}خلاصه sync — فایروال: {fw_name}{Style.RESET_ALL}")
    print(f"  {Fore.GREEN}+ اضافه شد   : {stats['added']}{Style.RESET_ALL}")
    print(f"  {Fore.BLUE}~ آپدیت شد   : {stats['updated']}{Style.RESET_ALL}")
    print(f"  {Fore.YELLOW}= رد شد      : {stats['skipped']}{Style.RESET_ALL}")
    print(f"  {Fore.MAGENTA}- حذف شد     : {stats['deleted']}{Style.RESET_ALL}")
    print(f"  {Fore.RED}x خطا        : {stats['failed']}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}{'-' * 40}{Style.RESET_ALL}")


# ---------------------------------------------------------------------------
# خواندن اکسل
# ---------------------------------------------------------------------------

def file_fingerprint(path: Path) -> str:
    stat = path.stat()
    return f"{stat.st_mtime_ns}:{stat.st_size}"


def read_excel_entries(cfg: dict) -> tuple[list[dict], bool]:
    excel_cfg = cfg["excel"]
    wb = openpyxl.load_workbook(excel_cfg["path"], data_only=True)
    sheet_name = excel_cfg.get("sheet_name")
    ws = wb[sheet_name] if sheet_name else wb.active

    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    headers = {}
    for idx, value in enumerate(header_row):
        if value:
            headers[str(value).strip().lower()] = idx

    name_col = excel_cfg.get("name_column", "Name").strip().lower()
    ip_col = excel_cfg.get("ip_column", "IP").strip().lower()
    fw_col = excel_cfg.get("firewall_column", "Firewall").strip().lower()

    if ip_col not in headers:
        raise ValueError(
            f"ستون '{excel_cfg.get('ip_column')}' توی هدر اکسل پیدا نشد. "
            f"ستون‌های موجود: {list(headers.keys())}"
        )

    ip_idx = headers[ip_col]
    name_idx = headers.get(name_col)
    fw_idx = headers.get(fw_col)

    entries = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if ip_idx >= len(row):
            continue
        ip_value = row[ip_idx]
        if not ip_value:
            continue
        ip_value = str(ip_value).strip()

        name_value = ""
        if name_idx is not None and name_idx < len(row) and row[name_idx]:
            name_value = str(row[name_idx]).strip()
        if not name_value:
            name_value = f"AUTO_{ip_value.replace('/', '_').replace(':', '_')}"

        fw_value = ""
        if fw_idx is not None and fw_idx < len(row) and row[fw_idx]:
            fw_value = str(row[fw_idx]).strip()

        entries.append({"name": name_value, "ip": ip_value, "firewall": fw_value})

    return entries, fw_idx is not None


def entries_for_firewall(entries: list[dict], fw_name: str, has_firewall_column: bool) -> list[dict]:
    if not has_firewall_column:
        return entries
    fw_name_lower = fw_name.strip().lower()
    result = []
    for entry in entries:
        fw_value = entry["firewall"].strip().lower()
        if not fw_value or fw_value in ("all", "*") or fw_value == fw_name_lower:
            result.append(entry)
    return result


def warn_unknown_firewalls(entries: list[dict], fw_names: list[str], has_firewall_column: bool) -> None:
    if not has_firewall_column:
        return
    known = {n.strip().lower() for n in fw_names} | {"", "all", "*"}
    unknown = {entry["firewall"] for entry in entries if entry["firewall"].strip().lower() not in known}
    for value in unknown:
        logging.warning(f"مقدار ستون Firewall ناشناخته و رد شد (به هیچ فایروالی نمی‌خوره): '{value}'")


def normalize_subnet(ip_value: str) -> tuple[str, str]:
    ip_value = ip_value.strip()
    if "/" in ip_value:
        network = ipaddress.ip_network(ip_value, strict=False)
        return str(network.network_address), str(network.netmask)
    ipaddress.ip_address(ip_value)  # اعتبارسنجی
    return ip_value, "255.255.255.255"


# ---------------------------------------------------------------------------
# ارتباط با فایروال
# ---------------------------------------------------------------------------

def fetch_existing_objects(conn) -> dict[str, tuple[str, str]]:
    output = conn.send_command("show firewall address", read_timeout=30)
    objects: dict[str, tuple[str, str]] = {}
    current_name = None
    for raw_line in output.splitlines():
        line = raw_line.strip()
        if line.startswith('edit "'):
            current_name = line[len('edit "'):].rstrip('"')
        elif line.startswith("set subnet") and current_name:
            parts = line.split()
            if len(parts) >= 4:
                objects[current_name] = (parts[2], parts[3])
            elif len(parts) == 3 and "/" in parts[2]:
                net = ipaddress.ip_network(parts[2], strict=False)
                objects[current_name] = (str(net.network_address), str(net.netmask))
        elif line == "next":
            current_name = None
    return objects


def create_address_object(conn, name: str, network: str, netmask: str) -> None:
    commands = [
        "config firewall address",
        f'edit "{name}"',
        f"set subnet {network} {netmask}",
        "next",
        "end",
    ]
    conn.send_config_set(commands, read_timeout=30)


def delete_address_object(conn, name: str) -> None:
    commands = [
        "config firewall address",
        f'delete "{name}"',
        "end",
    ]
    conn.send_config_set(commands, read_timeout=30)


# ---------------------------------------------------------------------------
# state (اینکه کدوم آبجکت‌ها رو خود اسکریپت ساخته)
# ---------------------------------------------------------------------------

def load_state(path: Path) -> dict:
    if not path.exists():
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_state(path: Path, state: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# گزارش اکسل
# ---------------------------------------------------------------------------

def append_report(report_path: Path, rows: list[dict]) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    if report_path.exists():
        wb = openpyxl.load_workbook(report_path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "SyncReport"
        ws.append(REPORT_HEADERS)

    for row in rows:
        ws.append([row.get(h, "") for h in REPORT_HEADERS])

    wb.save(report_path)


# ---------------------------------------------------------------------------
# منطق اصلی sync برای یک فایروال
# ---------------------------------------------------------------------------

def sync_firewall(
    fw_cfg: dict,
    entries: list[dict],
    state: dict,
    report_rows: list[dict],
    delay: float,
    delete_removed: bool,
) -> dict:
    fw_name = fw_cfg.get("name", fw_cfg["host"])
    logging.info(f"── فایروال: {fw_name} ({fw_cfg['host']}) ──")

    stats = {"added": 0, "updated": 0, "skipped": 0, "deleted": 0, "failed": 0}
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    device = {
        "device_type": fw_cfg.get("device_type", "fortinet"),
        "host": fw_cfg["host"],
        "username": fw_cfg["username"],
        "password": fw_cfg["password"],
        "port": fw_cfg.get("port", 22),
    }

    try:
        conn = ConnectHandler(**device)
    except (NetmikoTimeoutException, NetmikoAuthenticationException) as exc:
        logging.error(f"[{fw_name}] اتصال SSH ناموفق: {exc}")
        stats["failed"] += 1
        report_rows.append({"Time": now_str, "Firewall": fw_name, "Name": "", "IP": "",
                             "Action": "ConnectFailed", "Message": str(exc)})
        return stats

    managed_prev = state.get(fw_name, {})
    managed_now: dict[str, list[str]] = {}

    try:
        existing = fetch_existing_objects(conn)
        existing_subnets = set(existing.values())
        logging.info(f"[{fw_name}] {len(existing)} آبجکت آدرس موجود روی فایروال.")

        desired: dict[str, tuple[str, str]] = {}
        ip_display: dict[str, str] = {}
        for entry in entries:
            try:
                network, netmask = normalize_subnet(entry["ip"])
            except ValueError as exc:
                logging.warning(f"[{fw_name}] IP نامعتبر رد شد ({entry['ip']}): {exc}")
                stats["failed"] += 1
                report_rows.append({"Time": now_str, "Firewall": fw_name, "Name": entry["name"],
                                     "IP": entry["ip"], "Action": "InvalidIP", "Message": str(exc)})
                continue
            desired[entry["name"]] = (network, netmask)
            ip_display[entry["name"]] = entry["ip"]

        for name, subnet in desired.items():
            network, netmask = subnet
            ip_value = ip_display[name]
            current = existing.get(name)

            if current == subnet:
                managed_now[name] = list(subnet)
                stats["skipped"] += 1
                logging.info(f"[{fw_name}] رد شد (از قبل درست روی فایروال هست): {name} -> {ip_value}")
                report_rows.append({"Time": now_str, "Firewall": fw_name, "Name": name,
                                     "IP": ip_value, "Action": "Skipped", "Message": "از قبل موجود بود"})
                continue

            if current is None and subnet in existing_subnets:
                stats["skipped"] += 1
                logging.warning(f"[{fw_name}] رد شد (این IP از قبل با اسم دیگه‌ای ثبت شده): {name} -> {ip_value}")
                report_rows.append({"Time": now_str, "Firewall": fw_name, "Name": name,
                                     "IP": ip_value, "Action": "Skipped", "Message": "IP با اسم دیگه‌ای موجود بود"})
                continue

            try:
                create_address_object(conn, name, network, netmask)
                existing_subnets.add(subnet)
                managed_now[name] = list(subnet)
                if current is None:
                    stats["added"] += 1
                    logging.info(f"[{fw_name}] اضافه شد: {name} -> {ip_value}")
                    report_rows.append({"Time": now_str, "Firewall": fw_name, "Name": name,
                                         "IP": ip_value, "Action": "Added", "Message": ""})
                else:
                    stats["updated"] += 1
                    logging.info(f"[{fw_name}] آپدیت شد: {name} -> {ip_value}")
                    report_rows.append({"Time": now_str, "Firewall": fw_name, "Name": name,
                                         "IP": ip_value, "Action": "Updated", "Message": ""})
            except Exception as exc:
                stats["failed"] += 1
                logging.error(f"[{fw_name}] خطا هنگام ثبت {name} ({ip_value}): {exc}")
                report_rows.append({"Time": now_str, "Firewall": fw_name, "Name": name,
                                     "IP": ip_value, "Action": "Failed", "Message": str(exc)})

            time.sleep(delay)  # فاصله بین دستورات تا فایروال هنگ نکنه

        if delete_removed:
            to_delete = [n for n in managed_prev if n not in desired]
            for name in to_delete:
                if name not in existing:
                    continue  # از قبل هم روی فایروال نبوده، چیزی برای حذف نیست
                try:
                    delete_address_object(conn, name)
                    stats["deleted"] += 1
                    logging.info(f"[{fw_name}] حذف شد (از اکسل پاک شده بود): {name}")
                    report_rows.append({"Time": now_str, "Firewall": fw_name, "Name": name,
                                         "IP": "", "Action": "Deleted", "Message": "از اکسل حذف شده بود"})
                except Exception as exc:
                    stats["failed"] += 1
                    logging.error(f"[{fw_name}] خطا هنگام حذف {name} (احتمالاً داخل یه پالیسی/گروه استفاده شده): {exc}")
                    managed_now[name] = managed_prev[name]  # نگه‌دار برای تلاش دوباره دفعه بعد
                    report_rows.append({"Time": now_str, "Firewall": fw_name, "Name": name,
                                         "IP": "", "Action": "DeleteFailed", "Message": str(exc)})
                time.sleep(delay)
        else:
            for name, subnet in managed_prev.items():
                if name not in desired:
                    managed_now.setdefault(name, subnet)
    finally:
        conn.disconnect()

    state[fw_name] = managed_now
    return stats


# ---------------------------------------------------------------------------
# چرخه اصلی
# ---------------------------------------------------------------------------

def sync_once(cfg: dict) -> None:
    delay = cfg["sync"].get("delay_between_commands_seconds", 3)
    delete_removed = cfg["sync"].get("delete_removed_objects", True)
    state_path = Path(cfg.get("state_file", "sync_state.json"))
    report_path = Path(cfg.get("report_file", "reports/sync_report.xlsx"))

    try:
        entries, has_fw_col = read_excel_entries(cfg)
    except Exception as exc:
        logging.error(f"خواندن اکسل ناموفق بود: {exc}")
        return

    if not entries:
        logging.info("هیچ ردیفی توی اکسل پیدا نشد.")
        return

    firewalls = cfg["firewalls"]
    fw_names = [fw.get("name", fw["host"]) for fw in firewalls]
    warn_unknown_firewalls(entries, fw_names, has_fw_col)

    state = load_state(state_path)
    report_rows: list[dict] = []

    for fw_cfg in firewalls:
        fw_name = fw_cfg.get("name", fw_cfg["host"])
        fw_entries = entries_for_firewall(entries, fw_name, has_fw_col)
        if not fw_entries:
            logging.info(f"[{fw_name}] هیچ ردیفی برای این فایروال نیست، رد شد.")
            continue
        stats = sync_firewall(fw_cfg, fw_entries, state, report_rows, delay, delete_removed)
        print_summary(fw_name, stats)

    save_state(state_path, state)
    if report_rows:
        append_report(report_path, report_rows)
        logging.info(f"گزارش اکسل ذخیره/آپدیت شد: {report_path}")


def sleep_or_stop(seconds: float, stop_event=None) -> bool:
    """می‌خوابه، مگر اینکه stop_event ست بشه. برمی‌گردونه True یعنی باید متوقف بشه."""
    if stop_event is None:
        time.sleep(seconds)
        return False
    import win32event
    result = win32event.WaitForSingleObject(stop_event, int(seconds * 1000))
    return result == win32event.WAIT_OBJECT_0


def watch(cfg: dict, stop_event=None) -> None:
    excel_path = Path(cfg["excel"]["path"])
    poll_interval = cfg["sync"].get("poll_interval_seconds", 5)

    if not excel_path.exists():
        logging.error(f"فایل اکسل پیدا نشد: {excel_path}")
        if stop_event is None:
            sys.exit(1)
        return

    logging.info(f"شروع نظارت روی {excel_path} (هر {poll_interval} ثانیه بررسی می‌شود)")
    last_fingerprint = None

    while True:
        try:
            current_fingerprint = file_fingerprint(excel_path)
            if current_fingerprint != last_fingerprint:
                if last_fingerprint is None:
                    logging.info("اجرای اولیه — همگام‌سازی با وضعیت فعلی اکسل ...")
                else:
                    logging.info("تغییر در فایل اکسل شناسایی شد — شروع sync ...")
                    sleep_or_stop(1, stop_event)  # مطمئن بشیم اکسل کامل ذخیره شده
                sync_once(cfg)
                last_fingerprint = file_fingerprint(excel_path)
        except Exception as exc:
            logging.error(f"خطای غیرمنتظره: {exc}")

        if sleep_or_stop(poll_interval, stop_event):
            logging.info("حلقه نظارت متوقف شد.")
            break


def main() -> None:
    cfg = load_config()
    setup_logging(cfg.get("log_file", "fortinet_sync.log"))
    print_banner()
    try:
        watch(cfg)
    except KeyboardInterrupt:
        logging.info("متوقف شد توسط کاربر.")


if __name__ == "__main__":
    main()
