"""
اجرای fortinet_address_sync.py به‌صورت یک سرویس ویندوز — همیشه در پس‌زمینه فعال
می‌ماند و با بالا آمدن ویندوز خودش استارت می‌شود.

نصب (Run as Administrator):
    pip install -r requirements.txt
    python fortinet_sync_service.py install
    python fortinet_sync_service.py start

مدیریت سرویس:
    python fortinet_sync_service.py stop
    python fortinet_sync_service.py remove
    net start FortinetExcelSync
    net stop FortinetExcelSync

لاگ‌ها هم توی فایلی که در config.json مشخص کردی نوشته می‌شن و هم توی
Event Viewer > Windows Logs > Application قابل مشاهده‌ان.
"""

import servicemanager
import win32event
import win32service
import win32serviceutil

from fortinet_address_sync import load_config, setup_logging, watch


class FortinetSyncService(win32serviceutil.ServiceFramework):
    _svc_name_ = "FortinetExcelSync"
    _svc_display_name_ = "Fortinet Excel Address Sync"
    _svc_description_ = (
        "همگام‌سازی خودکار address objectهای فایروال Fortinet از روی یک فایل اکسل"
    )

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, ""),
        )
        cfg = load_config()
        setup_logging(cfg.get("log_file", "fortinet_sync.log"), colorize=False)
        watch(cfg, stop_event=self.stop_event)


if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(FortinetSyncService)
