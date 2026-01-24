# -*- coding: utf-8 -*-
"""
Shared reporting configuration.

Set REPORT_DATE_OVERRIDE to a YYYY-MM-DD string when you want all email
senders to target a specific day instead of auto-selecting the latest
working day from summary_all_days.xlsx. Leave it as None to keep the
existing behaviour.
"""

from __future__ import annotations
from pathlib import Path
import datetime as dt
import csv
from typing import Callable, Iterable, Set

# Example: REPORT_DATE_OVERRIDE = "2025-11-30"
REPORT_DATE_OVERRIDE = None
SENT_LOG_PATH = Path("sent_log.csv")
SENT_LOG_DETAIL_PATH = Path("sent_log_detail.csv")
LOCAL_DB_PATH = Path(r"C:\Users\gsakhare\AppData\Local\KCReports\kc_reports.db")

# Optional: only consider report dates on/after this date when selecting pending work.
# Example: TRACKING_START_DATE = "2026-01-21"
TRACKING_START_DATE = "2026-01-21"


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
            raise ValueError(
                f"Forced report date {forced.isoformat()} not present "
                "in summary_all_days.xlsx. Refresh the summary first."
            )
        if forced.weekday() == 6:
            print(f"[WARN] Forced report date {forced} is a Sunday; continuing per override.")
        if forced in holidays:
            print(f"[WARN] Forced report date {forced} is listed as a holiday; continuing per override.")
        return forced

    return fallback_picker(dates, holidays)


def _working_days_only(dates: Iterable[dt.date], holidays: Set[dt.date]) -> list[dt.date]:
    out = []
    seen = set()
    for d in dates:
        if not d or not isinstance(d, dt.date):
            continue
        if d.weekday() == 6 or d in holidays:  # skip Sundays/holidays
            continue
        if d in seen:
            continue
        seen.add(d)
        out.append(d)
    return sorted(out)


def _ensure_sent_log_dir():
    SENT_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)

def _parse_date_or_none(raw: str | None) -> dt.date | None:
    if not raw:
        return None
    try:
        return dt.date.fromisoformat(str(raw).strip())
    except Exception:
        return None

def _ensure_csv_header(path: Path, header: list[str]):
    _ensure_sent_log_dir()
    if not path.exists() or path.stat().st_size == 0:
        with path.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(header)
        return

    # If the file exists but doesn't start with the header, prepend it once.
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
        for idx, row in enumerate(reader):
            if not row:
                continue
            if row[0].lower() == "reporttype":
                # Handle malformed header with data appended: "Statusstaff_mailer,..."
                if len(row) >= 7 and row[3].lower().startswith("status") and row[3].lower() != "status":
                    rtype = row[3][len("status"):]
                    rdate = row[4] if len(row) > 4 else ""
                    status = row[6] if len(row) > 6 else ""
                    if rtype and status.lower() in ok_set:
                        try:
                            out.add(dt.date.fromisoformat(rdate))
                        except Exception:
                            pass
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
            if not row:
                continue
            if row[0].lower() == "reporttype":
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
    start_dt = _parse_date_or_none(TRACKING_START_DATE)
    if start_dt:
        dates = [d for d in dates if d >= start_dt]

    # Respect override (even if already sent)
    if REPORT_DATE_OVERRIDE:
        forced = resolve_report_date(dates, holidays, fallback_picker)
        return [forced] if forced else []

    working = _working_days_only(dates, holidays)
    sent = load_sent_dates(report_type)
    pending = [d for d in working if d not in sent]
    if pending:
        pending_sorted = sorted(pending)
        if max_days and len(pending_sorted) > max_days:
            pending_sorted = pending_sorted[-max_days:]
        return pending_sorted

    # No pending: fall back to the latest working day if not already logged
    fallback = fallback_picker(working, holidays) if fallback_picker else None
    if fallback and fallback not in sent:
        return [fallback]
    return []
