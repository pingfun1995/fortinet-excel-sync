# Fortinet Excel Address Sync

اسکریپتی که یک فایل اکسل (Name / IP) رو زیر نظر می‌گیره و به محض این‌که آپدیتش کنی،
خودش با SSH به فایروال (یا فایروال‌های) Fortinet وصل می‌شه و address objectها رو
همگام می‌کنه — بدون این‌که خودت مجبور باشی دستی وارد CLI فایروال بشی.

## قابلیت‌ها

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

## شروع سریع (برای مبتدی‌ها)

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

## راه‌اندازی دستی (برای حرفه‌ای‌ها)

```powershell
pip install -r requirements.txt
copy config.json.example config.json
notepad config.json
python fortinet_address_sync.py
```

## فرمت فایل اکسل

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

## فایل تنظیمات (`config.json`)

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

## اجرای دائمی به‌صورت سرویس ویندوز

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

## نکات امنیتی

- پسورد فایروال داخل `config.json` به‌صورت متن ساده ذخیره می‌شه. این فایل رو هیچوقت commit نکن —
  توی `.gitignore` هست، ولی حواست باشه.
- حذف خودکار (`delete_removed_objects`) فقط address objectهایی رو حذف می‌کنه که **خودِ همین اسکریپت** آن‌ها را
  ساخته (طبق `sync_state.json`). اگه یه آبجکت رو خودت دستی روی فایروال ساخته باشی، هیچوقت این اسکریپت لمسش نمی‌کنه.
  اگه اصلاً نمی‌خوای حذف خودکار داشته باشی، `delete_removed_objects` رو `false` بذار.
- اگه آبجکتی داخل یه Policy یا Group استفاده شده باشه، فایروال اجازه‌ی حذفش رو نمی‌ده — اسکریپت این خطا رو
  لاگ می‌کنه و توی `sync_state.json` نگهش می‌داره تا دفعه‌ی بعد دوباره امتحان کنه.

## عیب‌یابی

| مشکل | راه‌حل |
|---|---|
| `فایل config.json پیدا نشد` | از `config.json.example` کپی بگیر یا `setup.ps1` رو اجرا کن |
| اتصال SSH ناموفقه | آی‌پی/پورت/یوزر/پسورد رو چک کن؛ مطمئن شو SSH روی اینترفیس مدیریتی فایروال فعاله |
| ستون IP پیدا نشد | اسم ستون‌ها توی هدر اکسل باید دقیقاً با `name_column`/`ip_column` توی `config.json` یکی باشه |
| فایروال چیزی حذف/اضافه نمی‌کنه | لاگ `fortinet_sync.log` رو ببین؛ معمولاً پیام خطای خودِ FortiOS هم توش هست |

## لایسنس

MIT — فایل [LICENSE](LICENSE) رو ببین.
