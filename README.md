# Fortinet Excel Address Sync

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)

**[🇬🇧 English](#-english)** | **[🇮🇷 فارسی](#-فارسی)**

---

## 🇬🇧 English

Watches an Excel sheet (`Name` / `IP` / `Firewall`) and auto-syncs
FortiGate/FortiOS firewall **address objects** over SSH whenever the sheet is
saved — adds new IPs, skips ones that already exist, removes objects deleted
from the sheet, supports multiple firewalls at once, and can run as a Windows
service. No more manually typing `config firewall address` into the FortiOS
CLI for every new IP a colleague drops in a spreadsheet.

Built with [netmiko](https://github.com/ktbyers/netmiko) +
[openpyxl](https://openpyxl.readthedocs.io/).

### Features

| Feature | Description |
|---|---|
| Automatic Excel watch | Any change (Save) to the Excel file is detected automatically and a sync starts |
| Skips duplicate IPs | If an IP already exists on the firewall, it's skipped without error and the script moves on |
| Auto-delete | A row removed from Excel gets its address object deleted from the firewall too — but only if this script created it |
| Delay between commands | A pause (default 3s) between each command so the firewall isn't hammered |
| Multiple firewalls at once | An optional `Firewall` column in Excel lets you target a specific firewall or all of them |
| Excel report | After every sync, `reports/sync_report.xlsx` is updated with per-row details (Added/Skipped/Deleted/Failed) |
| Runs as a Windows service | Starts automatically on boot and stays running in the background |
| Colorized output | Console logs are color-coded (green=added, blue=updated, yellow=skipped, red=error) |
| **IPsec site-to-site tunnels** | An optional `Tunnels` sheet auto-builds a full route-based IPsec VPN per row: Phase1, Phase2, static route, address objects, and both traffic-flow policies |

> ⚠️ **Test in a lab or maintenance window first.** This tool can create/delete firewall
> policies and VPN tunnels, not just address objects. A mistake here can cut off site
> connectivity. Always dry-run against a test FortiGate (or GNS3/EVE-NG) before pointing
> it at production.

### Quick start (beginners)

1. Install [Python 3.10+](https://www.python.org/downloads/) (check **Add python.exe to PATH** during install).
2. Download/clone this folder.
3. Right-click PowerShell, **Run as Administrator**, `cd` into this folder, and run:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\setup.ps1
   ```

4. The script asks for your firewall details (host/user/password) and Excel path, then generates `config.json` and a sample `ips.xlsx`.
5. Open `ips.xlsx`, fill in your real rows, and Save.
6. Run it:

   ```powershell
   .\.venv\Scripts\python.exe fortinet_address_sync.py
   ```

From now on, every time you Save the Excel file, new IPs get created and changes get pushed to the firewall.

### Manual setup (advanced users)

```powershell
pip install -r requirements.txt
copy config.json.example config.json
notepad config.json
python fortinet_address_sync.py
```

### Excel file format

First sheet, header in row 1. Column names are configurable (see below).

| Name              | IP           | Firewall  |
|-------------------|--------------|-----------|
| Web-Server-01     | 10.0.0.10    |           |
| Branch-Isfahan    | 10.20.0.0/24 |           |
| Only-On-FW-Main   | 10.30.0.5    | FW-Main   |

- **Name** is optional — if left empty, a name is auto-generated from the IP (`AUTO_10.0.0.10`).
- **IP** can be a single address (`1.2.3.4` → treated as `/32`) or a subnet (`10.20.0.0/24`).
- **Firewall** is optional — only relevant if you have multiple firewalls in `config.json`. Empty or `ALL` applies to every firewall; otherwise it must exactly match a firewall's `name` field in `config.json`.

Run `python create_sample_excel.py` to generate a sample file (`ips.sample.xlsx`) in this format.

### IPsec tunnels (`Tunnels` sheet)

Add a second sheet called `Tunnels` to the same Excel file (the sample generator already
creates one) with these columns:

| Name | RemoteGateway | PSK | LocalSubnet | RemoteSubnet | Interface | Firewall |
|---|---|---|---|---|---|---|
| VPN-To-Branch-Isfahan | 203.0.113.10 | ChangeThisPreSharedKey! | 10.0.0.0/24 | 10.20.0.0/24 | wan1 | |

For each row, the script builds a complete **route-based** site-to-site tunnel:

1. Two address objects (`<Name>_local`, `<Name>_remote`) for the local/remote subnets
2. `vpn ipsec phase1-interface` (IKEv2, AES256-SHA256, DH group 14 by default)
3. `vpn ipsec phase2-interface`
4. A static route to the remote subnet via the tunnel interface
5. Two firewall policies (`<Name>_out`, `<Name>_in`) allowing traffic in both directions

Each piece is created only if missing (so a partially-applied tunnel resumes cleanly next
run), and removing a row deletes everything in the reverse order (policies → route → phase2
→ phase1 → address objects). Only tunnels this script created are ever touched — see
[Security notes](#security-notes).

Defaults (WAN/LAN interface names, IKE proposal, DH group) are set once in `config.json`
under `"ipsec"` and apply to every row; `Interface` in the sheet is optional and overrides
the WAN interface per-row if you need to.

### Config file (`config.json`)

```jsonc
{
  "firewalls": [
    {
      "name": "FW-Main",       // used in Excel's Firewall column and in logs
      "host": "192.168.1.1",
      "port": 22,
      "username": "admin",
      "password": "CHANGE_ME",
      "device_type": "fortinet"
    }
  ],
  "excel": {
    "path": "ips.xlsx",
    "sheet_name": null,        // null = first sheet
    "name_column": "Name",
    "ip_column": "IP",
    "firewall_column": "Firewall"
  },
  "ipsec": {
    "sheet_name": "Tunnels",
    "name_column": "Name",
    "remote_gateway_column": "RemoteGateway",
    "psk_column": "PSK",
    "local_subnet_column": "LocalSubnet",
    "remote_subnet_column": "RemoteSubnet",
    "interface_column": "Interface",
    "firewall_column": "Firewall",
    "default_wan_interface": "wan1",
    "default_lan_interface": "internal",
    "ike_version": "2",
    "phase1_proposal": "aes256-sha256",
    "phase2_proposal": "aes256-sha256",
    "dhgrp": "14"
  },
  "sync": {
    "poll_interval_seconds": 5,            // how often to check if the Excel file changed
    "delay_between_commands_seconds": 3,   // delay between each firewall command
    "delete_removed_objects": true         // delete firewall objects whose row was removed from Excel?
  },
  "state_file": "sync_state.json",   // tracks which objects this script created (for safe deletion)
  "report_file": "reports/sync_report.xlsx",
  "log_file": "fortinet_sync.log"
}
```

Remove the `"ipsec"` block entirely (or just don't add a `Tunnels` sheet) if you only want
address-object sync. You can list multiple firewalls in the `firewalls` array — each sync
connects to all of them in turn.

### Running permanently as a Windows service

```powershell
# Run as Administrator
.\.venv\Scripts\python.exe fortinet_sync_service.py install
.\.venv\Scripts\python.exe fortinet_sync_service.py start
```

Managing the service:

```powershell
net stop FortinetExcelSync
net start FortinetExcelSync
.\.venv\Scripts\python.exe fortinet_sync_service.py remove
```

### Security notes

- The firewall password (in `config.json`) **and** every tunnel's pre-shared key (in the
  `Tunnels` sheet) are stored in plain text. Never commit `config.json` or a real `ips.xlsx` —
  both are already in `.gitignore`, but stay careful.
- Auto-delete (`delete_removed_objects`) only removes objects that **this script itself
  created** (tracked in `sync_state.json`) — address objects, and for tunnels, the full
  Phase1/Phase2/route/policy/address set. Anything you created manually on the firewall is
  never touched. Set `delete_removed_objects` to `false` to disable auto-delete entirely.
- If an object is referenced by a policy or group, the firewall will refuse to delete it — the
  script logs the failure and keeps it in `sync_state.json` to retry next time.
- Tunnel policies are created with `service "ALL"` and no NAT between the two subnets — the
  standard baseline for a site-to-site VPN. Tighten them manually in FortiOS afterward if you
  need narrower rules; this script won't touch a tunnel's policies once created unless the row
  is removed from Excel.

### Troubleshooting

| Problem | Solution |
|---|---|
| `config.json not found` | Copy from `config.json.example` or run `setup.ps1` |
| SSH connection fails | Check host/port/username/password; make sure SSH is enabled on the firewall's management interface |
| IP column not found | Column names in the Excel header must exactly match `name_column`/`ip_column` in `config.json` |
| Firewall doesn't add/delete anything | Check `fortinet_sync.log` — it usually includes FortiOS's own error message |
| `Tunnels` sheet is ignored | Make sure the `"ipsec"` block exists in `config.json` and the sheet name matches `ipsec.sheet_name` (default `Tunnels`) |
| Tunnel created but no traffic passes | Check the peer's Phase1/Phase2 proposal and PSK match exactly; also verify routing/NAT on the remote side — this script only configures the local FortiGate |

### License

MIT — see [LICENSE](LICENSE).

---

## 🇮🇷 فارسی

اسکریپتی که یک فایل اکسل (Name / IP) رو زیر نظر می‌گیره و به محض این‌که آپدیتش کنی،
خودش با SSH به فایروال (یا فایروال‌های) Fortinet وصل می‌شه و address objectها رو
همگام می‌کنه — بدون این‌که خودت مجبور باشی دستی وارد CLI فایروال بشی.

### قابلیت‌ها

| قابلیت | توضیح |
|---|---|
| نظارت خودکار روی اکسل | هر تغییری روی فایل اکسل (Save) خودش تشخیص داده می‌شه و sync شروع می‌شه |
| رد کردن IPهای تکراری | اگه IP از قبل روی فایروال هست، بدون خطا رد می‌شه و می‌ره سراغ بعدی |
| حذف خودکار | ردیفی که از اکسل پاک بشه، اگه خودِ اسکریپت اون رو ساخته بود، از فایروال هم حذف می‌شه |
| تاخیر بین دستورات | یک وقفه (پیش‌فرض ۳ ثانیه) بین هر دستور تا فایروال زیر فشار نره |
| چند فایروال هم‌زمان | با یه ستون اختیاری `Firewall` توی اکسل، هر ردیف رو می‌تونی به یک فایروال خاص یا همه‌شون بفرستی |
| گزارش اکسل | بعد از هر sync یه فایل `reports/sync_report.xlsx` با جزئیات هر ردیف (اضافه/رد/حذف/خطا) آپدیت می‌شه |
| اجرا به‌صورت سرویس ویندوز | با بالا اومدن ویندوز خودش استارت می‌شه و همیشه در پس‌زمینه فعاله |
| خروجی رنگی | لاگ کنسول رنگیه (سبز=اضافه شد، آبی=آپدیت، زرد=رد شد، قرمز=خطا) |
| **تانل IPsec سایت‌به‌سایت** | یه شیت اختیاری `Tunnels` — برای هر ردیف یه IPsec VPN کامل route-based می‌سازه: Phase1، Phase2، static route، address objectها و هر دو پالیسی عبور ترافیک |

> ⚠️ **اول توی یه محیط تست یا maintenance window امتحان کن.** این ابزار Policy و VPN Tunnel هم
> می‌سازه/حذف می‌کنه، نه فقط address object. یه اشتباه اینجا می‌تونه ارتباط بین سایت‌ها رو قطع کنه.
> همیشه اول روی یه FortiGate تست (یا GNS3/EVE-NG) امتحان کن، بعد بزن روی محیط واقعی.

### شروع سریع (برای مبتدی‌ها)

1. [Python 3.10+](https://www.python.org/downloads/) رو نصب کن (موقع نصب تیک **Add python.exe to PATH** رو بزن).
2. این پوشه رو دانلود/کلون کن.
3. روی PowerShell راست‌کلیک کن، **Run as Administrator** رو بزن، بیا توی همین پوشه و بزن:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\setup.ps1
   ```

4. اسکریپت ازت اطلاعات فایروال (آی‌پی/یوزر/پسورد) و مسیر فایل اکسل رو می‌پرسه و خودش `config.json` و یه `ips.xlsx` نمونه می‌سازه.
5. `ips.xlsx` رو باز کن، ردیف‌های واقعی خودت رو بنویس، Save کن.
6. اجرا کن:

   ```powershell
   .\.venv\Scripts\python.exe fortinet_address_sync.py
   ```

از این به بعد، هر بار اکسل رو Save کنی، خودش IPهای جدید رو می‌سازه و تغییرات رو روی فایروال اعمال می‌کنه.

### راه‌اندازی دستی (برای حرفه‌ای‌ها)

```powershell
pip install -r requirements.txt
copy config.json.example config.json
notepad config.json
python fortinet_address_sync.py
```

### فرمت فایل اکسل

شیت اول، هدر توی ردیف اول. اسم ستون‌ها قابل تنظیمه (پایین رو ببین).

| Name              | IP           | Firewall  |
|-------------------|--------------|-----------|
| Web-Server-01     | 10.0.0.10    |           |
| Branch-Isfahan    | 10.20.0.0/24 |           |
| Only-On-FW-Main   | 10.30.0.5    | FW-Main   |

- **Name** اختیاریه — اگه خالی باشه، خودکار از روی IP اسم می‌سازه (`AUTO_10.0.0.10`).
- **IP** می‌تونه تک آی‌پی (`1.2.3.4` → به‌صورت `/32`) یا ساب‌نت (`10.20.0.0/24`) باشه.
- **Firewall** اختیاریه — فقط وقتی چند فایروال توی `config.json` تعریف کرده باشی کاربرد داره. خالی یا `ALL` یعنی به همه‌ی فایروال‌ها اعمال بشه؛ در غیر این صورت باید دقیقاً برابر فیلد `name` یکی از فایروال‌ها توی `config.json` باشه.

برای شروع سریع: `python create_sample_excel.py` یه فایل نمونه (`ips.sample.xlsx`) با همین فرمت می‌سازه.

### تانل‌های IPsec (شیت `Tunnels`)

یه شیت دوم به اسم `Tunnels` به همون فایل اکسل اضافه کن (اسکریپت نمونه‌ساز خودش این شیت رو هم می‌سازه) با این ستون‌ها:

| Name | RemoteGateway | PSK | LocalSubnet | RemoteSubnet | Interface | Firewall |
|---|---|---|---|---|---|---|
| VPN-To-Branch-Isfahan | 203.0.113.10 | ChangeThisPreSharedKey! | 10.0.0.0/24 | 10.20.0.0/24 | wan1 | |

برای هر ردیف، اسکریپت یه تانل سایت‌به‌سایت **route-based** کامل می‌سازه:

1. دو تا address object (`<Name>_local`, `<Name>_remote`) برای ساب‌نت‌های لوکال/ریموت
2. `vpn ipsec phase1-interface` (پیش‌فرض IKEv2، AES256-SHA256، DH group 14)
3. `vpn ipsec phase2-interface`
4. یه static route به ساب‌نت ریموت از طریق اینترفیس تانل
5. دو پالیسی فایروال (`<Name>_out`, `<Name>_in`) برای عبور ترافیک هر دو جهت

هر تیکه فقط وقتی که وجود نداره ساخته می‌شه (یعنی اگه یه تانل نصفه‌ونیمه ساخته شده باشه، دفعه بعد از همون‌جا ادامه پیدا می‌کنه)، و پاک کردن ردیف از اکسل همه‌چیز رو با ترتیب برعکس حذف می‌کنه (پالیسی‌ها ← route ← phase2 ← phase1 ← address objectها). فقط تانل‌هایی که خودِ اسکریپت ساخته لمس می‌شن — بخش [نکات امنیتی](#نکات-امنیتی) رو ببین.

مقادیر پیش‌فرض (اسم اینترفیس WAN/LAN، proposal، DH group) یه‌بار توی `config.json` زیر `"ipsec"` تنظیم می‌شن و برای همه‌ی ردیف‌ها اعمال می‌شن؛ ستون `Interface` توی شیت اختیاریه و فقط اگه لازم بود اینترفیس WAN رو برای اون ردیف خاص عوض می‌کنه.

### فایل تنظیمات (`config.json`)

```jsonc
{
  "firewalls": [
    {
      "name": "FW-Main",       // اسم دلخواه، برای ستون Firewall توی اکسل و لاگ‌ها استفاده می‌شه
      "host": "192.168.1.1",
      "port": 22,
      "username": "admin",
      "password": "CHANGE_ME",
      "device_type": "fortinet"
    }
  ],
  "excel": {
    "path": "ips.xlsx",
    "sheet_name": null,        // null = شیت اول
    "name_column": "Name",
    "ip_column": "IP",
    "firewall_column": "Firewall"
  },
  "ipsec": {
    "sheet_name": "Tunnels",
    "name_column": "Name",
    "remote_gateway_column": "RemoteGateway",
    "psk_column": "PSK",
    "local_subnet_column": "LocalSubnet",
    "remote_subnet_column": "RemoteSubnet",
    "interface_column": "Interface",
    "firewall_column": "Firewall",
    "default_wan_interface": "wan1",
    "default_lan_interface": "internal",
    "ike_version": "2",
    "phase1_proposal": "aes256-sha256",
    "phase2_proposal": "aes256-sha256",
    "dhgrp": "14"
  },
  "sync": {
    "poll_interval_seconds": 5,            // هر چند ثانیه چک کنه اکسل تغییر کرده یا نه
    "delay_between_commands_seconds": 3,   // تاخیر بین هر دستور روی فایروال
    "delete_removed_objects": true         // ردیف‌های حذف‌شده از اکسل، از فایروال هم حذف بشن؟
  },
  "state_file": "sync_state.json",   // ردپای آبجکت‌هایی که خودِ اسکریپت ساخته (برای حذف امن)
  "report_file": "reports/sync_report.xlsx",
  "log_file": "fortinet_sync.log"
}
```

اگه فقط address sync می‌خوای، کل بخش `"ipsec"` رو حذف کن (یا اصلاً شیت `Tunnels` رو نساز). می‌تونی چند تا فایروال توی آرایه‌ی `firewalls` بذاری — هر sync، به‌ترتیب به همه‌شون وصل می‌شه.

### اجرای دائمی به‌صورت سرویس ویندوز

```powershell
# Run as Administrator
.\.venv\Scripts\python.exe fortinet_sync_service.py install
.\.venv\Scripts\python.exe fortinet_sync_service.py start
```

مدیریت سرویس:

```powershell
net stop FortinetExcelSync
net start FortinetExcelSync
.\.venv\Scripts\python.exe fortinet_sync_service.py remove
```

### نکات امنیتی

- پسورد فایروال (توی `config.json`) **و** Pre-Shared Key هر تانل (توی شیت `Tunnels`) به‌صورت متن ساده
  ذخیره می‌شن. هیچوقت `config.json` یا یه `ips.xlsx` واقعی رو commit نکن — هردو توی `.gitignore` هستن، ولی حواست باشه.
- حذف خودکار (`delete_removed_objects`) فقط چیزهایی رو حذف می‌کنه که **خودِ همین اسکریپت** ساخته
  (طبق `sync_state.json`) — چه address object، چه (برای تانل‌ها) کل مجموعه‌ی Phase1/Phase2/route/policy/address.
  هرچیزی که خودت دستی روی فایروال ساخته باشی هیچوقت لمس نمی‌شه. اگه اصلاً نمی‌خوای حذف خودکار داشته باشی،
  `delete_removed_objects` رو `false` بذار.
- اگه آبجکتی داخل یه Policy یا Group استفاده شده باشه، فایروال اجازه‌ی حذفش رو نمی‌ده — اسکریپت این خطا رو
  لاگ می‌کنه و توی `sync_state.json` نگهش می‌داره تا دفعه‌ی بعد دوباره امتحان کنه.
- پالیسی‌های تانل با `service "ALL"` و بدون NAT بین دو ساب‌نت ساخته می‌شن — تنظیمات پایه‌ی معمول یه VPN
  سایت‌به‌سایت. اگه لازمه محدودترش کنی، بعداً دستی توی FortiOS تغییرش بده؛ این اسکریپت بعد از ساخت، تا وقتی
  ردیف از اکسل حذف نشده، دیگه دست به پالیسی‌های تانل نمی‌زنه.

### عیب‌یابی

| مشکل | راه‌حل |
|---|---|
| `فایل config.json پیدا نشد` | از `config.json.example` کپی بگیر یا `setup.ps1` رو اجرا کن |
| اتصال SSH ناموفقه | آی‌پی/پورت/یوزر/پسورد رو چک کن؛ مطمئن شو SSH روی اینترفیس مدیریتی فایروال فعاله |
| ستون IP پیدا نشد | اسم ستون‌ها توی هدر اکسل باید دقیقاً با `name_column`/`ip_column` توی `config.json` یکی باشه |
| فایروال چیزی حذف/اضافه نمی‌کنه | لاگ `fortinet_sync.log` رو ببین؛ معمولاً پیام خطای خودِ FortiOS هم توش هست |
| شیت `Tunnels` نادیده گرفته می‌شه | مطمئن شو بخش `"ipsec"` توی `config.json` وجود داره و اسم شیت با `ipsec.sheet_name` (پیش‌فرض `Tunnels`) یکیه |
| تانل ساخته شد ولی ترافیک رد نمی‌شه | مطمئن شو proposal/PSK طرف مقابل دقیقاً یکیه؛ روتینگ/NAT سمت ریموت رو هم چک کن — این اسکریپت فقط سمت خودی رو تنظیم می‌کنه |

### لایسنس

MIT — فایل [LICENSE](LICENSE) رو ببین.
