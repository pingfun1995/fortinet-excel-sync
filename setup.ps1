# راه‌انداز تعاملی Fortinet Excel Sync — مخصوص کاربرهای مبتدی.
# اجرا: powershell -ExecutionPolicy Bypass -File .\setup.ps1

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot

function Write-Step($text) { Write-Host "`n==> $text" -ForegroundColor Cyan }
function Write-Ok($text)   { Write-Host "    OK: $text" -ForegroundColor Green }
function Write-Warn($text) { Write-Host "    ! $text" -ForegroundColor Yellow }
function Write-Err($text)  { Write-Host "    X $text" -ForegroundColor Red }

Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "  Fortinet Excel Address Sync - نصب و تنظیم اولیه" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan

# 1) بررسی پایتون
Write-Step "بررسی نصب بودن Python"
$pythonCmd = $null
foreach ($candidate in @("python", "py")) {
    try {
        & $candidate --version *> $null
        if ($LASTEXITCODE -eq 0) { $pythonCmd = $candidate; break }
    } catch {}
}
if (-not $pythonCmd) {
    Write-Err "Python پیدا نشد."
    Write-Host "    از https://www.python.org/downloads/ نسخه‌ی Python 3.10 یا بالاتر رو نصب کن،" -ForegroundColor Yellow
    Write-Host "    و موقع نصب حتماً تیک 'Add python.exe to PATH' رو بزن. بعد دوباره این اسکریپت رو اجرا کن." -ForegroundColor Yellow
    exit 1
}
Write-Ok "پیدا شد: $pythonCmd"

# 2) ساخت محیط مجازی و نصب وابستگی‌ها
Write-Step "ساخت virtual environment (.venv) و نصب پکیج‌ها"
$venvPath = Join-Path $root ".venv"
if (-not (Test-Path $venvPath)) {
    & $pythonCmd -m venv $venvPath
    Write-Ok "محیط مجازی ساخته شد."
} else {
    Write-Ok "محیط مجازی از قبل وجود داره."
}
$venvPython = Join-Path $venvPath "Scripts\python.exe"
& $venvPython -m pip install --upgrade pip *> $null
& $venvPython -m pip install -r (Join-Path $root "requirements.txt")
Write-Ok "پکیج‌ها نصب شدند."

# 3) ساخت config.json
Write-Step "تنظیم config.json"
$configPath = Join-Path $root "config.json"
if (Test-Path $configPath) {
    Write-Warn "config.json از قبل وجود داره — از این مرحله رد می‌شیم تا چیزی رو دوباره‌نویسی نکنیم."
    Write-Host "    اگه می‌خوای از اول تنظیم کنی، اول فایل config.json رو حذف/تغییر اسم بده." -ForegroundColor Yellow
} else {
    $firewalls = @()
    $addMore = $true
    $index = 1
    while ($addMore) {
        Write-Host "`n  -- فایروال شماره $index --" -ForegroundColor Cyan
        $name = Read-Host "  اسم دلخواه برای این فایروال (مثلاً FW-Main)"
        if ([string]::IsNullOrWhiteSpace($name)) { $name = "FW-$index" }
        $fwHost = Read-Host "  آی‌پی/هاست فایروال"
        $portInput = Read-Host "  پورت SSH (خالی = 22)"
        $port = 22
        if (-not [string]::IsNullOrWhiteSpace($portInput)) { $port = [int]$portInput }
        $username = Read-Host "  یوزرنیم SSH"
        $securePass = Read-Host "  پسورد SSH" -AsSecureString
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
        $password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

        $firewalls += [ordered]@{
            name         = $name
            host         = $fwHost
            port         = $port
            username     = $username
            password     = $password
            device_type  = "fortinet"
        }

        $again = Read-Host "  فایروال دیگه‌ای هم هست؟ (y/N)"
        $addMore = ($again -eq "y" -or $again -eq "Y")
        $index++
    }

    $excelPath = Read-Host "`n  مسیر فایل اکسل (خالی = ips.xlsx کنار همین اسکریپت)"
    if ([string]::IsNullOrWhiteSpace($excelPath)) { $excelPath = "ips.xlsx" }

    $delayInput = Read-Host "  تاخیر بین دستورات روی فایروال به ثانیه (خالی = 3)"
    $delay = 3
    if (-not [string]::IsNullOrWhiteSpace($delayInput)) { $delay = [int]$delayInput }

    $deleteAnswer = Read-Host "  وقتی یه ردیف از اکسل پاک شد، از فایروال هم حذف بشه؟ (Y/n)"
    $deleteRemoved = -not ($deleteAnswer -eq "n" -or $deleteAnswer -eq "N")

    $config = [ordered]@{
        firewalls = $firewalls
        excel     = [ordered]@{
            path             = $excelPath
            sheet_name       = $null
            name_column      = "Name"
            ip_column        = "IP"
            firewall_column  = "Firewall"
        }
        sync = [ordered]@{
            poll_interval_seconds        = 5
            delay_between_commands_seconds = $delay
            delete_removed_objects       = $deleteRemoved
        }
        state_file  = "sync_state.json"
        report_file = "reports/sync_report.xlsx"
        log_file    = "fortinet_sync.log"
    }

    $config | ConvertTo-Json -Depth 6 | Set-Content -Path $configPath -Encoding utf8
    Write-Ok "config.json ساخته شد."
}

# 4) ساخت فایل اکسل نمونه اگه اکسل واقعی وجود نداره
Write-Step "بررسی فایل اکسل"
$excelTarget = Join-Path $root "ips.xlsx"
if (-not (Test-Path $excelTarget)) {
    & $venvPython (Join-Path $root "create_sample_excel.py") *> $null
    Copy-Item (Join-Path $root "ips.sample.xlsx") $excelTarget
    Write-Ok "یه فایل نمونه ساخته و به‌عنوان ips.xlsx کپی شد — خودت ردیف‌های واقعی رو جایگزین کن."
} else {
    Write-Ok "ips.xlsx از قبل وجود داره، دست نمی‌زنیم."
}

# 5) جمع‌بندی
Write-Host "`n=====================================================" -ForegroundColor Green
Write-Host "  تمام شد!" -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Green
Write-Host "برای اجرای دستی (توی همین کنسول):"
Write-Host "  .\.venv\Scripts\python.exe fortinet_address_sync.py`n" -ForegroundColor Cyan
Write-Host "برای نصب به‌صورت سرویس ویندوز دائمی (Run as Administrator لازمه):"
Write-Host "  .\.venv\Scripts\python.exe fortinet_sync_service.py install"
Write-Host "  .\.venv\Scripts\python.exe fortinet_sync_service.py start`n" -ForegroundColor Cyan
Write-Host "فایل ips.xlsx رو باز کن، IPها/اسم‌هاشون رو بنویس و Save کن — بقیه‌ش خودکاره." -ForegroundColor Yellow
