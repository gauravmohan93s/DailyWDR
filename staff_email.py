# -*- coding: utf-8 -*-
"""
KC Daily — Individual/Staff Mailer (v3.3 - Refactored)
Consolidates logic via data_loader.py
"""

from __future__ import annotations
from pathlib import Path
import datetime as dt
import pandas as pd
import numpy as np
import re
import traceback
import time

from reporting_config import (
    append_sent_log,
    append_sent_detail,
    load_sent_recipients,
    resolve_pending_work_days,
)
from data_loader import load_settings, load_summary

# =========================== CONFIG =========================== #
SUMMARY_PATH  = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\raw\summary_all_days.xlsx")
SUMMARY_SHEET = "summary_daily_all"
SETTINGS_PATH = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\setting\wd_settings.xlsx")
OUT_DIR       = SUMMARY_PATH.parent / "staff_email_previews"

DRY_RUN    = False
SEND_VIA   = "outlook"
SUBJECT_PREFIX = "KC Daily Performance"
MAX_CATCHUP_DAYS = 7

SIG_HTML = """<br><br><p style='font-family:Calibri,Arial;font-size:14px;'><strong>Thanks & Regards,</strong><br><strong>Gaurav Mohan Sakhare</strong></p>"""
WKHTMLTOPDF_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

# UI
BRAND_PRIMARY = "#0B1220"; BRAND_ACCENT = "#2563EB"; SURFACE_BG = "#F5F7FB"; CARD_BG = "#FFFFFF"; TEXT_DARK = "#0F172A"; TEXT_MUTED = "#6B7280"; DIVIDER = "#E5E7EB"

# ========= HELPERS ========= #
def is_sunday(d: dt.date) -> bool: return d.weekday() == 6
def latest_working_day(dates: list[dt.date], holidays: set[dt.date]) -> dt.date | None:
    for d in sorted(set(dates), reverse=True):
        if (not is_sunday(d)) and (d not in holidays): return d
    return None

def wrap_card(inner: str) -> str: return f"<table role='presentation' width='100%' style='background:{CARD_BG};border:1px solid {DIVIDER};border-radius:12px;'><tr><td style='padding:14px;'>{inner}</td></tr></table>"

# ============== SECTIONS ============== #
def build_staff_html(row: pd.Series, work_day: dt.date, role_kra_codes: dict, labels_all: dict) -> str:
    role = str(row.get("Role", "")); codes = role_kra_codes.get(role, []); items = []
    for c in codes:
        val = int(row.get(c, 0)); mtd = int(row.get(f"MTD_{c}", 0)); lbl = labels_all.get(c, c)
        items.append(f"<tr><td style='padding:8px;border-bottom:1px solid {DIVIDER};'>{lbl}</td><td style='padding:8px;border-bottom:1px solid {DIVIDER};text-align:right;'><b>{val}</b></td><td style='padding:8px;border-bottom:1px solid {DIVIDER};text-align:right;'>{mtd}</td></tr>")
    table = f"<table role='presentation' width='100%' style='border-collapse:collapse;'><thead><tr style='color:{TEXT_MUTED};font-size:12px;'><th align='left'>Measure</th><th align='right'>Today</th><th align='right'>MTD</th></tr></thead><tbody>{''.join(items)}</tbody></table>"
    badges = str(row.get("Badges_Today", ""))
    badge_html = f"<div style='margin-top:12px;'>{badges}</div>" if badges else ""
    return f"<!doctype html><html><body style='margin:0;background:{SURFACE_BG};padding:16px;'><div style='max-width:600px;margin:0 auto;'><div style='background:{BRAND_PRIMARY};color:#FFF;padding:16px;border-radius:12px;margin-bottom:16px;'><strong>KC Daily — {row.get('EmployeeName')} — {work_day.isoformat()}</strong></div>{wrap_card(f'<strong>KRA Score: {round(row.get('KRA_Score',0),1)}</strong>' + table + badge_html)}</div></body></html>"

# ============== MAIL & PDF ============== #
def html_to_pdf(html: str, pdf_path: Path) -> bool:
    try:
        import pdfkit; config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)
        pdfkit.from_string(html, str(pdf_path), configuration=config, options={"quiet":""})
        return True
    except: return False

def send_outlook(to_addr: str, subject: str, html: str, attachments: list[Path]):
    import pythoncom, win32_utils; pythoncom.CoInitialize()
    try:
        ol = win32_utils.safe_dispatch("Outlook.Application"); mail = ol.CreateItem(0)
        mail.Subject = subject; mail.HTMLBody = html + SIG_HTML; mail.To = to_addr
        for p in attachments: mail.Attachments.Add(Source=str(p))
        if DRY_RUN: mail.Display(False)
        else: mail.Send()
    finally: pythoncom.CoUninitialize()

from concurrent.futures import ThreadPoolExecutor

# ============== MAIN ============== #
def process_staff_member(r: pd.Series, work_day: dt.date, role_kra_codes: dict, labels_all: dict, sr: set):
    email = str(r.get("EmployeeEmail", "")).lower()
    if not email or email in sr: return
    try:
        html = build_staff_html(r, work_day, role_kra_codes, labels_all)
        pp = OUT_DIR / f"Staff_Report_{r.get('EmployeeName')}_{work_day.isoformat()}.pdf"
        html_to_pdf(html, pp)
        send_outlook(email, f"{SUBJECT_PREFIX} | {work_day.isoformat()} | {r.get('EmployeeName')}", html, [pp] if pp.exists() else [])
        append_sent_detail('staff_mailer', work_day, email, 'dry_run' if DRY_RUN else 'ok')
    except Exception as e:
        print(f"[ERROR] Parallel fail for {email}: {e}")
        append_sent_detail('staff_mailer', work_day, email, 'fail')

def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True); cfg = load_settings(SETTINGS_PATH); role_kra_codes = cfg["role_kra_codes"]; labels_all = cfg["label_map"]; holidays = cfg["holidays"]
    df = load_summary(SUMMARY_PATH, SUMMARY_SHEET)
    if df.empty: return
    work_days = resolve_pending_work_days(df['ReportDate'].dropna().tolist(), holidays, latest_working_day, 'staff_mailer', MAX_CATCHUP_DAYS)
    for work_day in work_days:
        try:
            df_day = df[df['ReportDate']==work_day].copy(); sr = set(load_sent_recipients('staff_mailer', work_day))
            print(f"[STAFF] Sending reports for {work_day} (Parallel)...")
            with ThreadPoolExecutor(max_workers=4) as executor:
                for _, r in df_day.iterrows():
                    executor.submit(process_staff_member, r, work_day, role_kra_codes, labels_all, sr)
            append_sent_log('staff_mailer', work_day, 'dry_run' if DRY_RUN else 'ok')
        except: traceback.print_exc()

if __name__ == '__main__':
    main()
