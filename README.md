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

You can list multiple firewalls in the `firewalls` array — each sync connects to all of them in turn.

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

- The firewall password is stored in plain text in `config.json`. Never commit this file —
  it's already in `.gitignore`, but stay careful.
- Auto-delete (`delete_removed_objects`) only removes address objects that **this script itself
  created** (tracked in `sync_state.json`). If you created an object manually on the firewall,
  this script will never touch it. Set `delete_removed_objects` to `false` to disable auto-delete entirely.
- If an object is referenced by a policy or group, the firewall will refuse to delete it — the
  script logs the failure and keeps it in `sync_state.json` to retry next time.

### Troubleshooting

| Problem | Solution |
|---|---|
| `config.json not found` | Copy from `config.json.example` or run `setup.ps1` |
| SSH connection fails | Check host/port/username/password; make sure SSH is enabled on the firewall's management interface |
| IP column not found | Column names in the Excel header must exactly match `name_column`/`ip_column` in `config.json` |
| Firewall doesn't add/delete anything | Check `fortinet_sync.log` — it usually includes FortiOS's own error message |

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

می‌تونی چند تا فایروال توی آرایه‌ی `firewalls` بذاری — هر sync، به‌ترتیب به همه‌شون وصل می‌شه.

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

- پسورد فایروال داخل `config.json` به‌صورت متن ساده ذخیره می‌شه. این فایل رو هیچوقت commit نکن —
  توی `.gitignore` هست، ولی حواست باشه.
- حذف خودکار (`delete_removed_objects`) فقط address objectهایی رو حذف می‌کنه که **خودِ همین اسکریپت** آن‌ها را
  ساخته (طبق `sync_state.json`). اگه یه آبجکت رو خودت دستی روی فایروال ساخته باشی، هیچوقت این اسکریپت لمسش نمی‌کنه.
  اگه اصلاً نمی‌خوای حذف خودکار داشته باشی، `delete_removed_objects` رو `false` بذار.
- اگه آبجکتی داخل یه Policy یا Group استفاده شده باشه، فایروال اجازه‌ی حذفش رو نمی‌ده — اسکریپت این خطا رو
  لاگ می‌کنه و توی `sync_state.json` نگهش می‌داره تا دفعه‌ی بعد دوباره امتحان کنه.

### عیب‌یابی

| مشکل | راه‌حل |
|---|---|
| `فایل config.json پیدا نشد` | از `config.json.example` کپی بگیر یا `setup.ps1` رو اجرا کن |
| اتصال SSH ناموفقه | آی‌پی/پورت/یوزر/پسورد رو چک کن؛ مطمئن شو SSH روی اینترفیس مدیریتی فایروال فعاله |
| ستون IP پیدا نشد | اسم ستون‌ها توی هدر اکسل باید دقیقاً با `name_column`/`ip_column` توی `config.json` یکی باشه |
| فایروال چیزی حذف/اضافه نمی‌کنه | لاگ `fortinet_sync.log` رو ببین؛ معمولاً پیام خطای خودِ FortiOS هم توش هست |

### لایسنس

MIT — فایل [LICENSE](LICENSE) رو ببین.
