# -*- coding: utf-8 -*-
"""
Shared reporting configuration and path centralization for DailyWD-UK.
"""

from __future__ import annotations
from pathlib import Path
import datetime as dt
import csv
from typing import Callable, Iterable, Set

import json

# =========================== DYNAMIC CONFIG =========================== #
APP_CONFIG_PATH = Path("app_config.json")

def load_app_config() -> dict:
    defaults = {
        "REPORT_DATE_OVERRIDE": "2026-05-11",
        "DRY_RUN": True,
        "DISK_CLEANUP_DAYS": 14
    }
    if not APP_CONFIG_PATH.exists():
        with open(APP_CONFIG_PATH, "w") as f:
            json.dump(defaults, f, indent=4)
        return defaults
    try:
        with open(APP_CONFIG_PATH, "r") as f:
            return {**defaults, **json.load(f)}
    except:
        return defaults

def save_app_config(config: dict):
    with open(APP_CONFIG_PATH, "w") as f:
        json.dump(config, f, indent=4)

_config = load_app_config()
REPORT_DATE_OVERRIDE = _config.get("REPORT_DATE_OVERRIDE")
DRY_RUN_GLOBAL       = _config.get("DRY_RUN", True)

# =========================== PATHS =========================== #
# Core Folders
PROJECT_ROOT = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action")
RAW_DIR      = PROJECT_ROOT / "raw"
SETTING_DIR  = PROJECT_ROOT / "setting"

# Databases & Logs
LOCAL_DB_PATH        = Path(r"C:\Users\gsakhare\AppData\Local\KCReports\kc_reports.db")
SENT_LOG_PATH        = Path("sent_log.csv")
SENT_LOG_DETAIL_PATH = Path("sent_log_detail.csv")

# Input Files
SETTINGS_PATH = SETTING_DIR / "wd_settings.xlsx"
SUMMARY_PATH  = RAW_DIR / "summary_all_days.xlsx"

# Assets (Signatures/Logos)
ASSETS_ROOT    = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\Conversion Report")
SIG_IMAGE_PATH = ASSETS_ROOT / "signature_banner.jpg"
KC_LOGO_PATH   = ASSETS_ROOT / "KC-LOGO-1-1024x225.png"

# Output Directories (Previews)
EXEC_OUT_DIR    = RAW_DIR / "executive_digest_previews"
MGR_OUT_DIR     = RAW_DIR / "manager_digest_previews"
STAFF_OUT_DIR   = RAW_DIR / "daily_email_previews"

# Binaries
WKHTMLTOPDF_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

# =========================== SETTINGS =========================== #
TRACKING_START_DATE = "2026-01-21"


# =========================== HELPERS =========================== #

def resolve_report_date(
    available_dates: Iterable[dt.date],
    holidays: Set[dt.date],
    fallback_picker: Callable[[list[dt.date], Set[dt.date]], dt.date | None],
) -> dt.date | None:
    dates = [d for d in available_dates if d]
    if REPORT_DATE_OVERRIDE:
        try:
            forced = dt.date.fromisoformat(REPORT_DATE_OVERRIDE)
        except ValueError as exc:
            raise ValueError("REPORT_DATE_OVERRIDE must be YYYY-MM-DD") from exc

        date_set = set(dates)
        if forced not in date_set:
            # Fallback: if not in summary but we want it, we'll try to find it later
            pass
        return forced

    return fallback_picker(dates, holidays)


def _ensure_sent_log_dir():
    SENT_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)

def _ensure_csv_header(path: Path, header: list[str]):
    _ensure_sent_log_dir()
    if not path.exists() or path.stat().st_size == 0:
        with path.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(header)
        return

    try:
        first_line = path.open("r", encoding="utf-8").readline().strip()
    except Exception:
        first_line = ""
    header_line = ",".join(header)
    if not first_line.lower().startswith(header_line.lower()):
        existing = path.read_text(encoding="utf-8")
        with path.open("w", encoding="utf-8", newline="") as f:
            f.write(header_line + "\n")
            if existing:
                if not existing.endswith("\n"):
                    existing += "\n"
                f.write(existing)


def append_sent_log(report_type: str, report_date: dt.date, status: str = "ok"):
    if not isinstance(report_date, dt.date):
        return
    _ensure_csv_header(SENT_LOG_PATH, ["ReportType", "ReportDate", "TimestampUTC", "Status"])
    ts = dt.datetime.now(dt.timezone.utc).isoformat()
    with SENT_LOG_PATH.open("a", encoding="utf-8", newline="") as f:
        csv.writer(f).writerow([report_type, report_date.isoformat(), ts, status])


def append_sent_detail(report_type: str, report_date: dt.date, recipient: str, status: str = "ok", meta: str = ""):
    if not isinstance(report_date, dt.date) or not recipient:
        return
    _ensure_csv_header(SENT_LOG_DETAIL_PATH, ["ReportType", "ReportDate", "Recipient", "TimestampUTC", "Status", "Meta"])
    ts = dt.datetime.now(dt.timezone.utc).isoformat()
    with SENT_LOG_DETAIL_PATH.open("a", encoding="utf-8", newline="") as f:
        csv.writer(f).writerow([report_type, report_date.isoformat(), recipient, ts, status, meta or ""])


def load_sent_dates(report_type: str, ok_status: set[str] | tuple[str, ...] = ("ok",)) -> set[dt.date]:
    if not SENT_LOG_PATH.exists():
        return set()
    ok_set = {s.lower() for s in ok_status}
    out: set[dt.date] = set()
    with SENT_LOG_PATH.open("r", encoding="utf-8", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row or row[0].lower() == "reporttype":
                continue
            if len(row) < 4:
                continue
            rtype, rdate, _ts, status = [c.strip() for c in row[:4]]
            if rtype.lower() != report_type.lower():
                continue
            if status.lower() not in ok_set:
                continue
            try:
                out.add(dt.date.fromisoformat(rdate))
            except Exception:
                continue
    return out


def load_sent_recipients(
    report_type: str,
    report_date: dt.date,
    ok_status: set[str] | tuple[str, ...] = ("ok",),
) -> set[str]:
    if not isinstance(report_date, dt.date) or not SENT_LOG_DETAIL_PATH.exists():
        return set()
    ok_set = {s.lower() for s in ok_status}
    out: set[str] = set()
    with SENT_LOG_DETAIL_PATH.open("r", encoding="utf-8", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row or row[0].lower() == "reporttype":
                continue
            if len(row) < 5:
                continue
            rtype, rdate, recipient, _ts, status = [c.strip() for c in row[:5]]
            if rtype.lower() != report_type.lower():
                continue
            if status.lower() not in ok_set:
                continue
            try:
                row_date = dt.date.fromisoformat(rdate)
            except Exception:
                continue
            if row_date != report_date:
                continue
            if recipient:
                out.add(recipient.strip().lower())
    return out


def resolve_pending_work_days(
    available_dates: Iterable[dt.date],
    holidays: Set[dt.date],
    fallback_picker: Callable[[list[dt.date], Set[dt.date]], dt.date | None],
    report_type: str,
    max_days: int = 7,
) -> list[dt.date]:
    dates = [d for d in available_dates if d]
    if REPORT_DATE_OVERRIDE:
        forced = resolve_report_date(dates, holidays, fallback_picker)
        return [forced] if forced else []

    def _working_days_only(dates: Iterable[dt.date], holidays: Set[dt.date]) -> list[dt.date]:
        out = []; seen = set()
        for d in dates:
            if not d or not isinstance(d, dt.date): continue
            if d.weekday() == 6 or d in holidays: continue
            if d in seen: continue
            seen.add(d); out.append(d)
        return sorted(out)

    working = _working_days_only(dates, holidays)
    sent = load_sent_dates(report_type)
    pending = [d for d in working if d not in sent]
    if pending:
        pending_sorted = sorted(pending)
        if max_days and len(pending_sorted) > max_days:
            pending_sorted = pending_sorted[-max_days:]
        return pending_sorted

    fallback = fallback_picker(working, holidays) if fallback_picker else None
    if fallback and fallback not in sent:
        return [fallback]
    return []
