# -*- coding: utf-8 -*-
"""
Simple orchestrator for the daily KC UK reporting flow (v2 - Database).
Enhanced with JSON progress reporting for UI.
"""

from __future__ import annotations
import traceback
import sys
import datetime as dt
import csv
import os
import time
import json
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path

import daily_wd_data_extractor as extractor
import wd_summariser
import exec_email
import mg_tl_email
import staff_email
import settings_validator

from reporting_config import SETTINGS_PATH, EXEC_OUT_DIR, MGR_OUT_DIR, STAFF_OUT_DIR

# Check for JSON progress flag
JSON_PROGRESS = "--json-progress" in sys.argv

def report_progress(msg: str, step_index: int, total_steps: int):
    if JSON_PROGRESS:
        progress = int((step_index / total_steps) * 100)
        print(json.dumps({"step": step_index, "total": total_steps, "progress": progress, "msg": msg}), flush=True)
    else:
        print(f"\n=== {msg} ===", flush=True)

def run_settings_validation():
    if not settings_validator.validate_settings(SETTINGS_PATH):
        raise ValueError("Settings validation failed. Please fix entries in wd_settings.xlsx.")

def run_disk_cleanup(days_to_keep: int = 14):
    """Deletes preview files older than X days to save space."""
    print(f"\n=== Disk Cleanup (older than {days_to_keep} days) ===")
    folders = [EXEC_OUT_DIR, MGR_OUT_DIR, STAFF_OUT_DIR]
    now = time.time()
    deleted_count = 0
    for folder in folders:
        if not folder.exists(): continue
        print(f"[CLEANUP] Scanning: {folder.name}")
        for file in folder.iterdir():
            if file.is_file():
                age_days = (now - file.stat().st_mtime) / (24 * 3600)
                if age_days > days_to_keep:
                    try:
                        file.unlink()
                        deleted_count += 1
                    except: pass
    print(f"[CLEANUP] Deleted {deleted_count} old preview files.")

STEPS = [
    ("Validating Settings", run_settings_validation),
    ("Updating Database", extractor.fetch_and_upsert_recent_data),
    ("Summarising Data", wd_summariser.main),
    ("Executive Reports", exec_email.main),
    ("Manager Reports", mg_tl_email.main),
    ("Staff Mailers", staff_email.main),
    ("Disk Cleanup", run_disk_cleanup),
]


def run_pipeline(stop_on_error: bool = True):
    log_dir = Path("logs"); log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"daily_flow_{dt.datetime.now().date().isoformat()}.log"

    class Tee:
        def __init__(self, *streams): self.streams = streams
        def write(self, data):
            for s in self.streams:
                s.write(data); s.flush()
        def flush(self):
            for s in self.streams: s.flush()

    with log_path.open("a", encoding="utf-8") as f:
        # Don't Tee if JSON_PROGRESS to avoid corrupting JSON stream
        if not JSON_PROGRESS:
            sys.stdout = Tee(sys.stdout, f)
            sys.stderr = Tee(sys.stderr, f)
        
        print(f"[RUN] Daily flow started at {dt.datetime.now().isoformat()}", flush=True)
        total = len(STEPS)
        for i, (label, fn) in enumerate(STEPS, 1):
            report_progress(label, i, total)
            try:
                fn()
            except Exception as exc:
                print(f"[ERROR] Step '{label}' failed: {exc}", flush=True)
                if not JSON_PROGRESS: traceback.print_exc()
                if stop_on_error: raise
        
        print(f"\n[DONE] Daily flow finished at {dt.datetime.now().isoformat()}", flush=True)
        _write_send_summary(log_dir)


def _write_send_summary(log_dir: Path):
    detail_path = Path("sent_log_detail.csv")
    if not detail_path.exists(): return
    rows = []
    with detail_path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row or row[0].lower() == "reporttype": continue
            if len(row) >= 5: rows.append(row[:5])
    if not rows: return
    try: max_date = max(r[1] for r in rows if r[1])
    except: return
    summary = {}
    for rtype, rdate, _recipient, _ts, status in rows:
        if rdate == max_date:
            summary[(rtype, status)] = summary.get((rtype, status), 0) + 1
    out_path = log_dir / f"send_summary_{max_date}.csv"
    with out_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f); w.writerow(["ReportType", "ReportDate", "Status", "Count"])
        for (rtype, status), count in sorted(summary.items()):
            w.writerow([rtype, max_date, status, count])


if __name__ == "__main__":
    run_pipeline()
