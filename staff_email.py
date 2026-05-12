# -*- coding: utf-8 -*-
"""
KC Daily — Individual Mailer (v3.9 FINAL — Refactored & Optimized)
- Phase A (Parallel): Generate all PDFs and Excel files.
- Phase B (Sequential): Send emails via Outlook (bypasses concurrency issues).
- SQL-Level Filtering: Fetches only required rows from DB (saves RAM/Time).
"""

from __future__ import annotations
from pathlib import Path
import datetime as dt
import re, hashlib, hmac, base64, urllib.parse as _url
import pandas as pd
import numpy as np
import traceback
import time, gc
from concurrent.futures import ThreadPoolExecutor

from reporting_config import (
    append_sent_log,
    append_sent_detail,
    load_sent_recipients,
    resolve_pending_work_days,
    LOCAL_DB_PATH,
    SETTINGS_PATH,
    SUMMARY_PATH,
    KC_LOGO_PATH,
    SIG_IMAGE_PATH,
    STAFF_OUT_DIR as OUT_DIR,
    WKHTMLTOPDF_PATH,
    DRY_RUN_GLOBAL as DRY_RUN
)
from data_loader import load_settings, load_summary, apply_config_filters

# =========================== CONFIG =========================== #
SUMMARY_SHEET = "summary_daily_all"

# --- Email transport ---
ATTACH_PDF       = True
ATTACH_CERT_PDF  = True
SAVE_DRAFTS_WHEN_DRY_RUN = True
BATCH_SIZE       = 100
RETRY_ON_ERROR   = 2
MAX_CATCHUP_DAYS = 7

FROM_DISPLAY     = "gsakhare@kcoverseas.com"
SUBJECT_PREFIX   = "KC Daily Performance"

SIG_HTML = """ 
<br><br>
<p style="font-family:Calibri, Arial, sans-serif; font-size:14px; line-height:20px; margin:0;">
  <strong>Thanks & Regards,</strong><br>
  <strong>Gaurav Mohan Sakhare</strong> | Senior Lead (Operations)<br>
  KC Overseas Education Pvt Ltd<br>
  Off: +91 712 2222061/62/63
</p>
<p style="margin:8px 0 0 0;">
  <img src="cid:kc_sig" alt="KC Signature" style="max-width:140px;height:auto;border:0;display:block;">
</p>
"""

PDF_ENGINE = "wkhtmltopdf"
IST = dt.timezone(dt.timedelta(hours=5, minutes=30))

# UI
BRAND_PRIMARY = "#0B1220"; BRAND_ACCENT = "#2563EB"; SURFACE_BG = "#F5F7FB"; CARD_BG = "#FFFFFF"; TEXT_DARK = "#111827"; TEXT_MUTED = "#6B7280"; ACCENT_SUCCESS = "#16A34A"; ACCENT_WARN = "#B45309"; ACCENT_DANGER = "#DC2626"; DIVIDER = "#E5E7EB"; TABLE_HEAD_BG = "#F3F4F6"; PILL_BG = "#EEF2FF"; PILL_TEXT = "#3730A3"

ALLOWED_BADGE_TITLES = ["Top KRA", "Top Total Comments", "Top Total Offer Follow Ups", "Top Total Pending from Partner F/U", "Top Total Status Changed", "Top Submission"]

# ==================== HELPERS ==================== #
def _total_workdays_in_month(d: dt.date, holidays: set[dt.date]) -> int:
    import calendar
    start = dt.date(d.year, d.month, 1); end = dt.date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])
    cur = start; cnt = 0
    while cur <= end:
        if cur.weekday() != 6 and cur not in holidays: cnt += 1
        cur += dt.timedelta(days=1)
    return cnt

def _remaining_workdays_in_month(d: dt.date, holidays: set[dt.date]) -> int:
    import calendar
    last_day = dt.date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])
    cur = d + dt.timedelta(days=1); cnt = 0
    while cur <= last_day:
        if cur.weekday() != 6 and cur not in holidays: cnt += 1
        cur += dt.timedelta(days=1)
    return cnt

def _daily_targets_for_role(role: str, role_meas: pd.DataFrame) -> dict[str, float]:
    rm = role_meas[(role_meas["Role"] == role) & (role_meas["IncludeInKRA"] == True)].copy()
    rm["DailyTarget"] = pd.to_numeric(rm.get("DailyTarget", 0), errors="coerce").fillna(0.0)
    return {str(r["MeasureCode"]): float(r["DailyTarget"]) for _, r in rm.iterrows()}

def _safe_int(x):
    try: return int(round(float(x)))
    except: return 0

def _pill(text: str, bg: str=PILL_BG, fg: str=PILL_TEXT):
    return f'<span style="display:inline-block;padding:5px 12px;border-radius:999px;background:{bg};color:{fg};font:600 12px Inter,Segoe UI,Arial;margin-right:8px;">{text}</span>'

def is_sunday(d: dt.date) -> bool: return d.weekday() == 6
def latest_working_day(dates: list[dt.date], holidays: set[dt.date]) -> dt.date | None:
    for d in sorted(set(dates), reverse=True):
        if (not is_sunday(d)) and (d not in holidays): return d
    return None

# ==================== DATABASE OPTIMIZATION ==================== #
def load_member_actions_sql(email: str, day: dt.date) -> pd.DataFrame:
    """SQL-filtered fetch for specific employee on specific day."""
    from sqlalchemy import create_engine
    engine = create_engine(f"sqlite:///{LOCAL_DB_PATH}")
    # We use a safer approach for SQLite strings:
    df = pd.read_sql("SELECT * FROM actions WHERE EmployeeEmail = ?", engine, params=[email])
    if df.empty: return df
    
    df["ActionDate"] = pd.to_datetime(df["ActionDate"], errors='coerce').dt.tz_localize('UTC')
    df["ActionDateIST"] = df["ActionDate"].dt.tz_convert("Asia/Kolkata")
    df["ActionDateIST_Date"] = df["ActionDateIST"].dt.date
    
    sub = df[df["ActionDateIST_Date"] == day].copy()
    if sub.empty: return sub

    sub["ActionDateIST_Str"] = sub["ActionDateIST"].dt.strftime("%Y-%m-%d %H:%M:%S")
    def _fill_action_taken(row):
        s1 = str(row.get("SubType1","") or "").strip()
        if s1: return s1
        v2 = str(row.get("ActionType","") or "").strip()
        if v2: return v2
        return str(row.get("Action Group","") or "").strip()
    sub["ActionTaken"] = sub.apply(_fill_action_taken, axis=1)
    return sub

# ==================== COMPONENT SECTIONS ==================== #
def combined_kra_section(role: str, row: pd.Series, role_meas: pd.DataFrame, labels_all: dict) -> str:
    rm = role_meas[(role_meas["Role"] == role) & (role_meas["IncludeInKRA"])].copy()
    if rm.empty: return f"<div style='font:800 12px Inter;color:{TEXT_MUTED};text-transform:uppercase;'>KRA Components</div><div style='margin-top:8px;padding:12px;border:1px dashed {DIVIDER};border-radius:10px;background:#FAFAFB;color:{TEXT_MUTED}'>None.</div>"
    rm["_row_order"] = np.arange(len(rm)); rm["__has"] = rm["DisplayOrder"].notna().astype(int)
    rm = rm.sort_values(by=["__has","DisplayOrder","_row_order"], ascending=[False, True, True])
    total_w = float(rm["Weight"].fillna(0).sum()); rm["WeightNorm"] = (rm["Weight"].fillna(0) / total_w) if total_w > 0 else 0.0
    rows = []
    for _, r in rm.iterrows():
        code = str(r["MeasureCode"]); lbl = labels_all.get(code, code); wt_pct = int(round(float(r["WeightNorm"])*100.0))
        tgt = r.get("DailyTarget"); tgt_txt = f"{_safe_int(tgt)}" if pd.notna(tgt) and float(tgt)>0 else "-"
        rows.append(f"<tr><td style='padding:10px 12px;border-bottom:1px solid {DIVIDER};'>{lbl}</td><td style='padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};'>{wt_pct}%</td><td style='padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};'>{tgt_txt}</td><td style='padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;font-weight:700;'>{int(row.get(code, 0))}</td><td style='padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};'>{row.get(f'{code}_PctOfDay', 0.0)}%</td><td style='padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};'>#{_safe_int(row.get(f'Rank_{code}', 0))}</td></tr>")
    return f"<div style='font:800 12px Inter;color:{TEXT_MUTED};text-transform:uppercase;'>KRA Components & Today's Actuals</div><table width='100%' cellpadding='0' cellspacing='0' style='margin-top:10px;border-collapse:collapse;border-radius:10px;overflow:hidden;background:{CARD_BG};border:1px solid {DIVIDER};'><thead><tr style='background:{TABLE_HEAD_BG};color:{TEXT_MUTED};text-align:left;'><th style='padding:10px 12px;font:800 12px Inter;'>Measure</th><th style='padding:10px 12px;font:800 12px Inter;text-align:right;'>Weight</th><th style='padding:10px 12px;font:800 12px Inter;text-align:right;'>Target</th><th style='padding:10px 12px;font:800 12px Inter;text-align:right;'>Today</th><th style='padding:10px 12px;font:800 12px Inter;text-align:right;'>% of Day</th><th style='padding:10px 12px;font:800 12px Inter;text-align:right;'>Rank</th></tr></thead><tbody>{''.join(rows)}</tbody></table>"

def rolling_mtd_table(row: pd.Series, labels_kra_for_member: dict, role: str, role_meas: pd.DataFrame, holidays: set[dt.date]) -> str:
    d = row["ReportDate"]; total_workdays = _total_workdays_in_month(d, holidays); remaining_days = _remaining_workdays_in_month(d, holidays); daily_targets = _daily_targets_for_role(role, role_meas)
    mrows = []
    for code, label in labels_kra_for_member.items():
        cur = float(row.get(f"MTD_{code}", 0.0)); month_tgt = float(daily_targets.get(code, 0.0)) * total_workdays
        pct = 0 if month_tgt <= 0 else min(100, int(round((cur / month_tgt) * 100))); gap = max(0, int(round(month_tgt - cur))); need = int(np.ceil(gap / remaining_days)) if remaining_days > 0 else 0
        mrows.append(f"<tr><td style='padding:10px 12px;border-bottom:1px solid {DIVIDER};'>{label}</td><td style='padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;font-weight:700;'>{_safe_int(cur)}</td><td style='padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};'>{_safe_int(month_tgt)}</td><td style='padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};'>{pct}%</td><td style='padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};'>{gap}</td><td style='padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};'>{need}</td></tr>")
    return f"<div style='font:800 12px Inter;color:{TEXT_MUTED};text-transform:uppercase;'>Rolling MTD (Month-to-Date)</div><table width='100%' cellpadding='0' cellspacing='0' style='margin-top:10px;border-collapse:collapse;border-radius:10px;overflow:hidden;background:{CARD_BG};border:1px solid {DIVIDER};'><thead><tr style='background:{TABLE_HEAD_BG};color:{TEXT_MUTED};text-align:left;'><th style='padding:10px 12px;font:800 12px Inter;'>Measure</th><th style='padding:10px 12px;font:800 12px Inter;text-align:right;'>MTD</th><th style='padding:10px 12px;font:800 12px Inter;text-align:right;'>Month Target</th><th style='padding:10px 12px;font:800 12px Inter;text-align:right;'>%Complete</th><th style='padding:10px 12px;font:800 12px Inter;text-align:right;'>Gap</th><th style='padding:10px 12px;font:800 12px Inter;text-align:right;'>Needed</th></tr></thead><tbody>{''.join(mrows)}</tbody></table>"

def build_role_leaderboard(df_day: pd.DataFrame, role: str, labels_kra_for_role: dict) -> str:
    rdf = df_day[df_day["Role"] == role].copy()
    if rdf.empty: return ""
    codes = list(labels_kra_for_role.keys()); top = rdf.sort_values("KRA_Score", ascending=False).head(3); rows = []
    for _, r in top.iterrows():
        nm = r["EmployeeName"] or r["EmployeeEmail"]; parts = []
        for code in codes: parts.append(f"{labels_kra_for_role.get(code, code)}: <b>{int(r.get(code, 0))}</b>")
        rows.append(f"<tr><td style='padding:8px 10px;border-bottom:1px solid {DIVIDER};'><b>{nm}</b><div style='color:{TEXT_MUTED};font:600 12px Inter;'>{r.get('Region','')}</div></td><td style='padding:8px 10px;border-bottom:1px solid {DIVIDER};text-align:right;color:{BRAND_ACCENT};font-weight:800;'>{round(float(r.get('KRA_Score', 0.0)), 1)}</td><td style='padding:8px 10px;border-bottom:1px solid {DIVIDER};'>{' • '.join(parts)}</td></tr>")
    return f"<div style='background:{CARD_BG};border:1px solid {DIVIDER};border-radius:12px;padding:16px;'><div style='font:800 15px Inter;margin-bottom:6px;'>Role Newsletter — {role}</div><table width='100%' style='border-collapse:collapse;'><thead><tr style='background:{TABLE_HEAD_BG};color:{TEXT_MUTED};'><th align='left' style='padding:8px 10px;'>Member</th><th align='right' style='padding:8px 10px;'>KRA</th><th align='left' style='padding:8px 10px;'>Components Today</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div>"

def build_member_html(row: pd.Series, df_day: pd.DataFrame, labels_all: dict, labels_kra_for_member: dict, role_meas: pd.DataFrame, role_leaderboard_html: str, badges_today_html: str, badges_mtd_html: str, holidays: set[dt.date]) -> str:
    kra = row.get("KRA_Score", 0.0); kra_color = ACCENT_SUCCESS if float(kra) >= 100 else (ACCENT_WARN if float(kra) >= 60 else ACCENT_DANGER)
    badges = str(row.get("Badges_Today","")).strip(); badge_pills = " ".join([_pill(b.strip()) for b in badges.split(",") if b.strip()]) or _pill("No badges today", bg="#F3F4F6", fg="#6B7280")
    top_score = f"<div style='background:{CARD_BG};border:1px solid {DIVIDER};border-radius:12px;padding:18px;'><div style='font:800 12px Inter;color:{TEXT_MUTED};text-transform:uppercase;'>KRA Score</div><div style='font:900 42px Inter;color:{kra_color};margin:8px 0;'>{round(float(kra),1)}</div><div style='color:{TEXT_MUTED};font:600 12px Inter;'>Role Rank: #{_safe_int(row.get('KRA_Rank_Role',0))} • Overall Rank: #{_safe_int(row.get('KRA_Rank_Overall',0))}</div><div style='margin-top:12px;'>{_pill('Hit Target' if int(row.get('Hit_Target_Today',0)) else 'Below Target', ACCENT_SUCCESS if int(row.get('Hit_Target_Today',0)) else ACCENT_WARN, '#FFFFFF')} {_pill(f'Streak: {int(row.get('Daily_Streak',0))}', BRAND_ACCENT, '#FFFFFF')}</div></div>"
    return f"<!doctype html><html><body style='margin:0;background:{SURFACE_BG};padding:18px 0;font-family:Inter,Segoe UI,Arial;'><table role='presentation' width='100%' style='max-width:920px;margin:0 auto;'><tr><td style='padding:0 18px 14px 18px;'><div style='background:{CARD_BG};border:1px solid {DIVIDER};border-radius:12px;padding:18px;'><div style='font:800 20px Inter;color:{BRAND_PRIMARY};'>KC Daily — {row['ReportDate']}</div><div style='margin-top:6px;font:600 13px Inter;color:{TEXT_MUTED}'>{row.get('EmployeeName') or row.get('EmployeeEmail')} • {row.get('Role','')} • {row.get('Region','')}</div><div style='margin-top:12px;'>{badge_pills}</div></div></td></tr><tr><td style='padding:0 18px 0 18px;'>{top_score}</td></tr><tr><td style='padding:0 18px;'><table width='100%'><tr><td width='50%' style='padding-right:8px;vertical-align:top;'>{combined_kra_section(row.get('Role',''), row, role_meas, labels_all)}</td><td width='50%' style='padding-left:8px;vertical-align:top;'>{rolling_mtd_table(row, labels_kra_for_member, row.get('Role',''), role_meas, holidays)}</td></tr></table></td></tr><tr><td style='padding:18px 18px 8px 18px;'>{role_leaderboard_html}</td></tr><tr><td style='padding:0 18px;'><table width='100%'><tr><td width='50%' style='padding-right:8px;vertical-align:top;'>{badges_today_html}</td><td width='50%' style='padding-left:8px;vertical-align:top;'>{badges_mtd_html}</td></tr></table></td></tr><tr><td style='padding:18px;'><div style='height:1px;background:{DIVIDER};margin-bottom:12px;'></div><p>Thanks & Regards,<br><b>Gaurav Mohan Sakhare</b></p></td></tr></table></body></html>"

# ==================== PDF/EXCEL EXPORTERS ==================== #
def html_to_pdf(html: str, pdf_path: Path) -> bool:
    try:
        import pdfkit; config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)
        pdfkit.from_string(html, str(pdf_path), configuration=config, options={"quiet":"","enable-local-file-access": ""})
        return True
    except: return False

def export_actions_excel(email: str, day: dt.date, xlsx_path: Path) -> bool:
    sub = load_member_actions_sql(email, day)
    if sub.empty: return False
    cols = [("ActionDateIST_Str", "Action (IST)"), ("ActionType", "Action Type"), ("Action Group", "Action Group"), ("ActionTaken", "Action Taken"), ("AcknowledgementNumber", "Ack Number"), ("StudentId", "Student ID"), ("University", "University"), ("PartnerName", "Partner Name"), ("Assignee", "Assignee"), ("Region", "Region")]
    use = sub[[c[0] for c in cols]].rename(columns=dict(cols))
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as w: use.to_excel(w, index=False, sheet_name="Actions")
    return True

# ==================== PARALLEL PHASE A: GENERATION ==================== #
def generate_files_for_member(r: pd.Series, df_day: pd.DataFrame, labels_all: dict, role_labels: dict, role_meas: pd.DataFrame, role_blocks: dict, bt_html: str, bm_html: str, holidays: set[dt.date]):
    email = str(r["EmployeeEmail"]).lower(); name = r.get("EmployeeName") or email; work_day = r["ReportDate"]
    safe = re.sub(r'[^a-zA-Z0-9]+', '_', str(name))[:60]
    paths = {"html": None, "pdf": None, "xlsx": None}
    try:
        html = build_member_html(r, df_day, labels_all, role_labels.get(r['Role'],{}), role_meas, role_blocks.get(r['Role'],''), bt_html, bm_html, holidays)
        html_path = OUT_DIR / f"{work_day.isoformat()}_{safe}.html"; html_path.write_text(html, encoding='utf-8')
        paths["html"] = html_path
        pdf_path = OUT_DIR / f"{work_day.isoformat()}_{safe}_Report.pdf"
        if html_to_pdf(html, pdf_path): paths["pdf"] = pdf_path
        xlsx_path = OUT_DIR / f"{work_day.isoformat()}_{safe}_Actions.xlsx"
        if export_actions_excel(email, work_day, xlsx_path): paths["xlsx"] = xlsx_path
        return email, paths
    except: traceback.print_exc(); return email, None

# ==================== MAIN ==================== #
def send_outlook(to_addr: str, subject: str, html_path: Path, attachments: list[Path]):
    import pythoncom, win32_utils; pythoncom.CoInitialize()
    try:
        ol = win32_utils.safe_dispatch("Outlook.Application"); mail = ol.CreateItem(0)
        mail.Subject = subject; mail.HTMLBody = html_path.read_text(encoding='utf-8') + SIG_HTML; mail.To = to_addr
        if FROM_DISPLAY:
            try: mail.SentOnBehalfOfName = FROM_DISPLAY
            except: pass
        for p in attachments:
            if p and p.exists(): mail.Attachments.Add(Source=str(p))
        if DRY_RUN: mail.Save() if SAVE_DRAFTS_WHEN_DRY_RUN else mail.Display(False)
        else: mail.Send()
    finally: pythoncom.CoUninitialize()

def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True); cfg = load_settings(SETTINGS_PATH); team = cfg["team"]; labels_all = cfg["label_map"]; role_meas = cfg["role_meas"]; holidays = cfg["holidays"]
    global HOLIDAYS; HOLIDAYS = holidays

    df_all_dates = load_summary(SUMMARY_PATH, SUMMARY_SHEET)
    if df_all_dates.empty: print('No summary rows.'); return

    work_days = resolve_pending_work_days(df_all_dates['ReportDate'].dropna().tolist(), holidays, latest_working_day, 'staff_mailer', MAX_CATCHUP_DAYS)
    roster = team[["EmployeeEmail","EmployeeName","Role","Region","Include","SendTo"]].copy()

    for work_day in work_days:
        try:
            df_day = load_summary(SUMMARY_PATH, SUMMARY_SHEET, report_date=work_day)
            if df_day.empty: print(f'[WARN] No summary data for {work_day}.'); continue

            # APPLY UI FILTERS
            df_day = apply_config_filters(df_day)
            if df_day.empty: print(f'[INFO] No employees match filters for {work_day}. skipping.'); continue

            sr = load_sent_recipients('staff_mailer', work_day)
            merged = roster.merge(df_day, on='EmployeeEmail', how='inner', suffixes=('', '_sum'))
            merged = merged[pd.to_numeric(merged.get('KRA_Score',0), errors='coerce').fillna(0.0)>0].copy()
            if merged.empty: continue
            
            role_labels = {}; role_blocks = {}
            for role in merged['Role'].unique():
                rms = role_meas[(role_meas['Role']==role) & (role_meas['IncludeInKRA'])].copy()
                codes = [str(x) for x in rms['MeasureCode'].dropna().tolist()]; rlabels = {c: labels_all.get(c, c) for c in codes}
                role_labels[role] = rlabels; role_blocks[role] = build_role_leaderboard(merged, role, rlabels)
            
            from exec_email import build_badges_today_block, build_monthly_badges_block
            bt_html = build_badges_today_block(merged, labels_all); bm_html = build_monthly_badges_block(df_all_dates, work_day, labels_all)

            print(f"[STAFF] Phase A: Generating files for {len(merged)} members (Parallel)...")
            ready_files = {}
            with ThreadPoolExecutor(max_workers=4) as exc:
                futures = [exc.submit(generate_files_for_member, r, merged, labels_all, role_labels, role_meas, role_blocks, bt_html, bm_html, holidays) for _, r in merged.iterrows()]
                for f in futures:
                    email, p = f.result()
                    if p: ready_files[email] = p
            
            print(f"[STAFF] Phase B: Sending emails via Outlook (Sequential)...")
            for _, r in merged.iterrows():
                email = str(r['EmployeeEmail']).lower()
                if email in sr or email not in ready_files: continue
                p = ready_files[email]; att = [p[k] for k in ["pdf","xlsx"] if p[k]]
                subj = f"{SUBJECT_PREFIX} | {work_day.isoformat()} | {r.get('EmployeeName')}"
                send_outlook(r.get('SendTo') or email, subj, p["html"], att)
                append_sent_detail('staff_mailer', work_day, email, 'dry_run' if DRY_RUN else 'ok')
            
            append_sent_log('staff_mailer', work_day, 'dry_run' if DRY_RUN else 'ok')
        except: traceback.print_exc()

if __name__ == '__main__':
    main()
