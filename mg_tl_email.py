# -*- coding: utf-8 -*-
"""
KC Manager Daily — Digest (v3.3 - Refactored)
Consolidates logic via data_loader.py
"""

from __future__ import annotations
from pathlib import Path
import datetime as dt
import calendar
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
OUT_DIR       = SUMMARY_PATH.parent / "manager_digest_previews"

DRY_RUN    = False
SEND_VIA   = "outlook"
FROM_DISPLAY  = "gsakhare@kcoverseas.com"
SUBJECT_PREFIX = "KC Manager Daily"
MAX_CATCHUP_DAYS = 7
RETRY_ON_ERROR = 2

SIG_HTML = """<br><br><p style='font-family:Calibri,Arial;font-size:14px;'><strong>Thanks & Regards,</strong><br><strong>Gaurav Mohan Sakhare</strong></p>"""
SIG_IMAGE_PATH = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\Conversion Report\signature_banner.jpg")
WKHTMLTOPDF_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

# UI
BRAND_PRIMARY = "#0B1220"; BRAND_ACCENT = "#2563EB"; SURFACE_BG = "#F5F7FB"; CARD_BG = "#FFFFFF"; TEXT_DARK = "#0F172A"; TEXT_MUTED = "#6B7280"; DIVIDER = "#E5E7EB"; TABLE_HEAD_BG = "#EEF2FF"

# ========= HELPERS ========= #
def is_sunday(d: dt.date) -> bool: return d.weekday() == 6
def latest_working_day(dates: list[dt.date], holidays: set[dt.date]) -> dt.date | None:
    for d in sorted(set(dates), reverse=True):
        if (not is_sunday(d)) and (d not in holidays): return d
    return None

def working_days_in_month(reference_day: dt.date, holidays: set[dt.date]) -> int:
    first_day = reference_day.replace(day=1); _, last_day_num = calendar.monthrange(reference_day.year, reference_day.month)
    days = [first_day + dt.timedelta(days=i) for i in range(last_day_num)]
    return max(sum(1 for d in days if not is_sunday(d) and d not in holidays), 1)

def wrap_card(inner: str) -> str: return f"<table role='presentation' width='100%' style='background:{CARD_BG};border:1px solid {DIVIDER};border-radius:12px;'><tr><td style='padding:14px;'>{inner}</td></tr></table>"
def th(t, align="left"): return f"<th style='padding:8px 10px;background:{TABLE_HEAD_BG};border-bottom:1px solid {DIVIDER};color:{TEXT_MUTED};font:700 12px Segoe UI;text-align:{align};'>{t}</th>"
def td(t, right=False, bold=False): return f"<td style='padding:8px 10px;border-bottom:1px solid {DIVIDER};text-align:{'right' if right else 'left'};font-weight:{'700' if bold else '400'};color:{TEXT_DARK};'>{t}</td>"

def cell_t_mtd_pct(t, m, target):
    pct = 0 if target<=0 else min(100, int(round((m/target)*100)))
    return f"<div style='font:700 12px Segoe UI;'>{t} / {m}</div><div style='height:4px;background:{DIVIDER};border-radius:4px;margin-top:4px;'><div style='height:4px;width:{pct}%;background:{BRAND_ACCENT};border-radius:4px;'></div></div>"

# ============== SECTIONS ============== #
def build_manager_digest_html(manager: str, work_day: dt.date, df_day: pd.DataFrame, df_all: pd.DataFrame, role_kra_codes: dict, labels_all: dict, month_targets: dict) -> str:
    sub_day = df_day[df_day["Manager"] == manager].copy(); roles = sorted(sub_day["Role"].unique()); chunks = []
    for role in roles:
        codes = role_kra_codes.get(role, []); members = sub_day[sub_day["Role"] == role].sort_values("KRA_Score", ascending=False)
        if not codes or members.empty: continue
        head = [th("Member"), th("KRA", "right")] + [th(labels_all.get(c,c), "right") for c in codes]
        rows = []
        for _, r in members.iterrows():
            row = [td(r.get("EmployeeName") or r.get("EmployeeEmail")), td(str(round(r.get("KRA_Score",0),1)), True, True)]
            for c in codes: row.append(td(cell_t_mtd_pct(int(r.get(c,0)), int(r.get(f"MTD_{c}",0)), month_targets.get(role,{}).get(c,1)), True))
            rows.append("<tr>" + "".join(row) + "</tr>")
        chunks.append(wrap_card(f"<div style='font:800 14px Segoe UI;margin-bottom:8px;'>Role: {role}</div><table role='presentation' width='100%' style='border-collapse:collapse;'><thead><tr>{''.join(head)}</tr></thead><tbody>{''.join(rows)}</tbody></table>"))
    body = "\n".join(chunks) or "<div style='color:{TEXT_MUTED}'>No activity.</div>"
    return f"<!doctype html><html><body style='margin:0;background:{SURFACE_BG};padding:16px;'><div style='max-width:800px;margin:0 auto;'><div style='background:{BRAND_PRIMARY};color:#FFF;padding:16px;border-radius:12px;margin-bottom:16px;'><strong>KC Manager Daily — {manager} — {work_day.isoformat()}</strong></div>{body}</div></body></html>"

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

# ============== MAIN ============== #
def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True); cfg = load_settings(SETTINGS_PATH); team = cfg["team"]; role_kra_codes = cfg["role_kra_codes"]; labels_all = cfg["label_map"]; holidays = cfg["holidays"]; role_meas = cfg["role_meas"]
    df = load_summary(SUMMARY_PATH, SUMMARY_SHEET)
    if df.empty: return
    work_days = resolve_pending_work_days(df['ReportDate'].dropna().tolist(), holidays, latest_working_day, 'manager_digest', MAX_CATCHUP_DAYS)
    for work_day in work_days:
        try:
            wd = working_days_in_month(work_day, holidays); mt = {str(role): {str(r.MeasureCode): int(round(max(0.0, float(r.get("DailyTarget") or 0.0)) * wd)) for _, r in part.iterrows()} for role, part in role_meas[role_meas["IncludeInKRA"]==True].groupby("Role")}
            df_day = df[df['ReportDate']==work_day].copy(); managers = sorted(df_day["Manager"].unique()); sr = load_sent_recipients('manager_digest', work_day)
            for mg in managers:
                if not mg or str(mg).lower() == "nan": continue
                mg_email = team[team["EmployeeName"] == mg]["EmployeeEmail"].iloc[0] if not team[team["EmployeeName"] == mg].empty else None
                if not mg_email or mg_email.lower() in sr: continue
                html = build_manager_digest_html(mg, work_day, df_day, df, role_kra_codes, labels_all, mt)
                pp = OUT_DIR / f"Manager_Digest_{mg}_{work_day.isoformat()}.pdf"; html_to_pdf(html, pp)
                send_outlook(mg_email, f"{SUBJECT_PREFIX} | {work_day.isoformat()} | {mg}", html, [pp] if pp.exists() else [])
                append_sent_detail('manager_digest', work_day, mg_email.lower(), 'dry_run' if DRY_RUN else 'ok')
            append_sent_log('manager_digest', work_day, 'dry_run' if DRY_RUN else 'ok')
        except: traceback.print_exc()

if __name__ == '__main__':
    main()
