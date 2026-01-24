# -*- coding: utf-8 -*-
"""
Simple orchestrator for the daily KC UK reporting flow (v2 - Database).

Steps:
1) (One-time) Initializes the local SQLite database from historical CSVs.
2) Fetches recent records from the source DB and upserts them into the local DB.
3) Rebuild the wd_summariser outputs from the local DB.
4) Send Executive, Manager, and Staff digests.
"""

from __future__ import annotations
import traceback
import sys
import datetime as dt
import csv
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path

import daily_wd_data_extractor as extractor
import wd_summariser
import exec_email
import mg_tl_email
import staff_eamil

STEPS = [
    ("Initialize/Backfill Database (if needed)", extractor.initialize_database),
    ("Fetch & Upsert Recent Data", extractor.fetch_and_upsert_recent_data),
    ("WD Summariser", wd_summariser.main),
    ("Executive Digest", exec_email.main),
    ("Manager Digest", mg_tl_email.main),
    ("Staff Mailers", staff_eamil.main),
]


def run_pipeline(stop_on_error: bool = True):
    log_dir = Path("logs")
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"daily_flow_{dt.datetime.now().date().isoformat()}.log"

    class Tee:
        def __init__(self, *streams):
            self.streams = streams
        def write(self, data):
            for s in self.streams:
                s.write(data)
                s.flush()
        def flush(self):
            for s in self.streams:
                s.flush()

    with log_path.open("a", encoding="utf-8") as f:
        tee_out = Tee(sys.stdout, f)
        tee_err = Tee(sys.stderr, f)
        with redirect_stdout(tee_out), redirect_stderr(tee_err):
            print(f"[RUN] Daily flow started at {dt.datetime.now().isoformat()}")
            for label, fn in STEPS:
                print(f"\n=== {label} ===")
                try:
                    fn()
                except Exception as exc:
                    print(f"[ERROR] Step '{label}' failed: {exc}")
                    traceback.print_exc()
                    if stop_on_error:
                        raise
            print(f"[DONE] Daily flow finished at {dt.datetime.now().isoformat()}")
            _write_send_summary(log_dir)


def _write_send_summary(log_dir: Path):
    detail_path = Path("sent_log_detail.csv")
    if not detail_path.exists():
        return
    rows = []
    with detail_path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row or row[0].lower() == "reporttype":
                continue
            if len(row) < 5:
                continue
            rows.append(row[:5])
    if not rows:
        return
    # Use most recent report date in detail log
    try:
        max_date = max(r[1] for r in rows if r[1])
    except Exception:
        return
    summary = {}
    for rtype, rdate, _recipient, _ts, status in rows:
        if rdate != max_date:
            continue
        key = (rtype, status)
        summary[key] = summary.get(key, 0) + 1
    out_path = log_dir / f"send_summary_{max_date}.csv"
    with out_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ReportType", "ReportDate", "Status", "Count"])
        for (rtype, status), count in sorted(summary.items()):
            w.writerow([rtype, max_date, status, count])


if __name__ == "__main__":
    run_pipeline()
