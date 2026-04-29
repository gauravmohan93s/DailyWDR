# -*- coding: utf-8 -*-
"""
KC Executive Daily — ORG Roll-up (v3.3 - Refactored)
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
OUT_DIR       = SUMMARY_PATH.parent / "executive_digest_previews"

DRY_RUN    = False
SEND_VIA   = "outlook"
FROM_DISPLAY  = "gsakhare@kcoverseas.com"
SUBJECT_PREFIX = "KC Executive Daily"
MAX_CATCHUP_DAYS = 7
RETRY_ON_ERROR = 2

EXEC_RECIPIENTS = [
    "vbisen@kcoverseas.com",
    "sshroff@kcoverseas.com",
    "ketan@kcoverseas.com"
]

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
SIG_IMAGE_PATH = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\Conversion Report\signature_banner.jpg")

WKHTMLTOPDF_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
PDF_ENGINE = "wkhtmltopdf"
IST = dt.timezone(dt.timedelta(hours=5, minutes=30))

# =================== BRAND / UI =================== #
BRAND_PRIMARY    = "#0B1220"
BRAND_ACCENT     = "#2563EB"
SURFACE_BG       = "#F5F7FB"
CARD_BG          = "#FFFFFF"
TEXT_DARK        = "#0F172A"
TEXT_MUTED       = "#6B7280"
ACCENT_SUCCESS   = "#16A34A"
ACCENT_WARN      = "#B45309"
ACCENT_INFO      = "#1D4ED8"
DIVIDER          = "#E5E7EB"
TABLE_HEAD_BG    = "#EEF2FF"
PILL_BG          = "#EDE9FE"
PILL_TEXT        = "#3730A3"
TINT_GOOD        = "#ECFDF5"
TINT_BAD         = "#FEF2F2"

ALLOWED_BADGE_TITLES = [
    "Top KRA",
    "Top Submission",
    "Top Total Comments",
    "Top Total Offer Follow Ups",
    "Top Total Pending from Partner F/U",
    "Top Total Status Changed",
]

# ========= HELPERS ========= #
def is_sunday(d: dt.date) -> bool: return d.weekday() == 6
def latest_working_day(dates: list[dt.date], holidays: set[dt.date]) -> dt.date | None:
    for d in sorted(set(dates), reverse=True):
        if (not is_sunday(d)) and (d not in holidays): return d
    return None

def working_days_in_month(reference_day: dt.date, holidays: set[dt.date]) -> int:
    first_day = reference_day.replace(day=1)
    _, last_day_num = calendar.monthrange(reference_day.year, reference_day.month)
    days = [first_day + dt.timedelta(days=i) for i in range(last_day_num)]
    cnt = 0
    for d in days:
        if is_sunday(d) or d in holidays: continue
        cnt += 1
    return max(cnt, 1)

def _tidy_area(s: str) -> str:
    s = str(s or "").strip()
    s = re.sub(r'[\[\]\"]', "", s)
    s = re.sub(r'\s*,\s*', ", ", s)
    return s.strip()

def region_key(row) -> str:
    rg = _tidy_area(row.get("Region", ""))
    sr = _tidy_area(row.get("SubRegion", ""))
    return f"{rg} — {sr}" if sr else rg

def wrap_card(inner_html: str) -> str:
    return f"""<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:{CARD_BG};border:1px solid {DIVIDER};border-radius:12px;"><tr><td style="padding:14px;">{inner_html}</td></tr></table>"""

def hlabel(text: str) -> str:
    return f"""<div style="font:800 14px/1.2 Segoe UI,Arial;color:{TEXT_DARK};margin-bottom:8px;">{text}</div>"""

def th(txt: str, w: str = "", align: str="left") -> str:
    wstyle = f"width:{w};" if w else ""
    align_css = "text-align:right;" if align=="right" else "text-align:left;"
    return f"""<th style="padding:8px 10px;background:{TABLE_HEAD_BG};border-bottom:1px solid {DIVIDER};color:{TEXT_MUTED};font:700 12px Segoe UI,Arial;{wstyle}{align_css}">{txt}</th>"""

def td(txt: str, right: bool=False, bold: bool=False, muted: bool=False, bg: str="", nowrap: bool=False) -> str:
    ta = "text-align:right;" if right else "text-align:left;"
    fw = "font-weight:700;" if bold else ""
    col = f"color:{TEXT_MUTED};" if muted else f"color:{TEXT_DARK};"
    bgc = f"background:{bg};" if bg else ""
    nw = "white-space:nowrap;" if nowrap else ""
    return f"""<td style="padding:8px 10px;border-bottom:1px solid {DIVIDER};{ta}{fw}{col}{bgc}{nw};line-height:1.2;">{txt}</td>"""

def pill(txt: str, bg: str=PILL_BG, fg: str=PILL_TEXT) -> str:
    return f"""<span style="display:inline-block;padding:5px 10px;border-radius:999px;background:{bg};color:{fg};font:600 12px Segoe UI,Arial;">{txt}</span>"""

def fmt_i(x) -> int:
    try: return int(round(float(x)))
    except: return 0

_ABBR_MAP = {
    "Submission": "Subm", "Assessed": "Assess", "Total Comments": "Cmt(T)", "Unique Comments": "Cmt(U)",
    "Total Offer Follow Ups": "OFU(T)", "Unique Offer Follow Up": "OFU(U)",
    "Total Pending from Partner F/U": "PFU(T)", "Unique Pending from Partner F/U": "PFU(U)",
    "Total Status Changed": "Stat(T)", "Unique Status Changed": "Stat(U)", "Month-to-Date": "MTD",
}
def abbrev_label(lbl: str) -> str:
    s = (lbl or "").strip()
    for k, v in _ABBR_MAP.items(): s = re.sub(k, v, s, flags=re.I)
    return s

def cell_t_mtd_pct(t_val: int, m_val: int, month_target: int) -> str:
    pct = 0 if month_target<=0 else min(100, max(0, int(round((m_val / month_target) * 100))))
    return f"""
      <div style="font:700 12px Segoe UI,Arial;color:{TEXT_DARK};">{t_val} / {m_val}</div>
      <div style="height:6px;background:{DIVIDER};border-radius:8px;margin-top:4px;">
        <div style="height:6px;width:{pct}%;border-radius:8px;background:{ACCENT_INFO};"></div>
      </div>
      <div style="font:600 11px Segoe UI,Arial;color:{TEXT_MUTED};margin-top:2px;">{pct}%</div>
    """

def has_any_kra_today(row, role_kra_codes: dict) -> bool:
    role = str(row.get("Role","") or "")
    codes = role_kra_codes.get(role, [])
    return any(fmt_i(row.get(code, 0)) > 0 for code in codes)

# ============== SECTIONS ============== #
def build_availability_matrix_org(df_day_all: pd.DataFrame, team_df: pd.DataFrame, role_kra_codes: dict) -> str:
    part = df_day_all.copy()
    part["Team"] = part.apply(region_key, axis=1)
    part["Available"] = part.apply(lambda r: has_any_kra_today(r, role_kra_codes), axis=1)
    roles_all = sorted(set(part["Role"].dropna().astype(str)))
    roster_sub = team_df.copy(); roster_sub["Team"] = roster_sub.apply(region_key, axis=1)
    teams = sorted(set(roster_sub["Team"].dropna().astype(str)))
    count_map = {}
    for t in teams:
        for role in roles_all:
            total = int(((roster_sub["Team"]==t) & (roster_sub["Role"]==role)).sum())
            avail = int(((part["Team"]==t) & (part["Role"]==role) & (part["Available"]==True)).sum())
            count_map[(t, role)] = (avail, total)
    roles = [r for r in roles_all if sum(count_map[(t, r)][1] for t in teams) > 0]
    headers = [th("Team (Region — SubRegion)")] + [th(abbrev_label(r), align="right") for r in roles] + [th("Row Total", align="right")]
    body = []
    for t in teams:
        row_cells = [td(f"<b>{t}</b>")]
        row_av = row_to = 0
        for role in roles:
            av, tot = count_map[(t, role)]; row_av += av; row_to += tot
            row_cells.append(td(f"<b>{av}</b> <span style='color:{TEXT_MUTED}'>| {tot}</span>", right=True))
        row_cells.append(td(f"<b>{row_av}</b> <span style='color:{TEXT_MUTED}'>| {row_to}</span>", right=True, bold=True))
        body.append("<tr>" + "".join(row_cells) + "</tr>")
    gt_html = ""
    if len(teams) > 1:
        col_av = {r: sum(count_map[(t, r)][0] for t in teams) for r in roles}
        col_to = {r: sum(count_map[(t, r)][1] for t in teams) for r in roles}
        cells = [td("<b>Grand Total</b>", bold=True)] + [td(f"<b>{col_av[r]}</b> <span style='color:{TEXT_MUTED}'>| {col_to[r]}</span>", right=True) for r in roles]
        cells.append(td(f"<b>{sum(col_av.values())}</b> <span style='color:{TEXT_MUTED}'>| {sum(col_to.values())}</span>", right=True, bold=True))
        gt_html = "<tr>" + "".join(cells) + "</tr>"
    hint = f"<div style='margin-top:6px;color:{TEXT_MUTED};font:600 12px Segoe UI,Arial;'>Cell shows <b>Available</b> | Total. Available = any KRA measure &gt; 0 yesterday.</div>"
    return wrap_card(f"{hlabel('YESTERDAY SUMMARY — TEAM AVAILABILITY MATRIX (ORG)')} {hint}<table role='presentation' width='100%' cellspacing='0' cellpadding='0' style='border-collapse:collapse;border:1px solid {DIVIDER};border-radius:10px;overflow:hidden;'><thead><tr>{''.join(headers)}</tr></thead><tbody>{''.join(body)}{gt_html}</tbody></table>")

def build_action_summary_team_org(df_day_all: pd.DataFrame, role_kra_codes: dict, labels_all: dict, month_targets: dict) -> str:
    df = df_day_all.copy(); df["Team"] = df.apply(region_key, axis=1)
    df = df[df.apply(lambda r: has_any_kra_today(r, role_kra_codes), axis=1)].copy()
    codes_union = []
    for role in sorted(set(df["Role"].dropna().astype(str))):
        for c in role_kra_codes.get(role, []):
            if c not in codes_union: codes_union.append(c)
    if not codes_union: return wrap_card(hlabel("ACTION SUMMARY — TEAM (ORG)") + f"<div style='color:{TEXT_MUTED}'>No activity.</div>")
    head = [th("Team")] + [th(f"{abbrev_label(labels_all.get(c,c))}<br><small>T/MTD</small>", align="right") for c in codes_union] + [th("Total", align="right")]
    teams = sorted(set(df["Team"].dropna().astype(str))); body = []
    ctt = {c:0 for c in codes_union}; ctm = {c:0 for c in codes_union}; cte = {c:0 for c in codes_union}
    gt = gm = ge = 0
    for t in teams:
        sub = df[df["Team"]==t]; rc = [td(f"<b>{t}</b>")]; rt = rm = re = 0
        for c in codes_union:
            vt = int(sub.get(c,0).sum()); vm = int(sub.get(f"MTD_{c}",0).sum()); ve = sum(month_targets.get(str(r.Role),{}).get(c,0) for _,r in sub.iterrows())
            rt+=vt; rm+=vm; re+=ve; ctt[c]+=vt; ctm[c]+=vm; cte[c]+=ve
            rc.append(td(cell_t_mtd_pct(vt, vm, max(ve,1)), right=True))
        gt+=rt; gm+=rm; ge+=re; rc.append(td(cell_t_mtd_pct(rt, rm, max(re,1)), right=True))
        body.append("<tr>" + "".join(rc) + "</tr>")
    if len(teams)>1:
        tc = [td("<b>Grand Total</b>", bold=True)] + [td(cell_t_mtd_pct(ctt[c], ctm[c], max(cte[c],1)), right=True) for c in codes_union] + [td(cell_t_mtd_pct(gt, gm, max(ge,1)), right=True)]
        body.append("<tr>" + "".join(tc) + "</tr>")
    return wrap_card(f"{hlabel('ACTION SUMMARY — TEAM (ORG)')}<table role='presentation' width='100%' cellspacing='0' cellpadding='0' style='border-collapse:collapse;border:1px solid {DIVIDER};border-radius:10px;overflow:hidden;'><thead><tr>{''.join(head)}</tr></thead><tbody>{''.join(body)}</tbody></table>")

def build_rolewise_work_org(df_day_all: pd.DataFrame, role_kra_codes: dict, labels_all: dict, month_targets: dict) -> str:
    df = df_day_all.copy(); df["Team"] = df.apply(region_key, axis=1); chunks = []
    roles_all = [r for r in sorted(set(df["Role"].dropna().astype(str))) if role_kra_codes.get(r, [])]
    for role in roles_all:
        codes = role_kra_codes.get(role, []); sub = df[df["Role"]==role].copy()
        sub = sub[sub.apply(lambda r: any(fmt_i(r.get(c,0))>0 for c in codes), axis=1)].copy()
        if sub.empty: continue
        sub["KRA_Score"] = pd.to_numeric(sub.get("KRA_Score", 0), errors="coerce").fillna(0.0)
        sub = sub.sort_values("KRA_Score", ascending=False)
        head = [th("Member"), th("KRA", align="right")] + [th(f"{abbrev_label(labels_all.get(c,c))}<br><small>T/MTD</small>", align="right") for c in codes] + [th("Total", align="right")]
        rows = []; col_t = {c:0 for c in codes}; col_m = {c:0 for c in codes}; col_e = {c:0 for c in codes}
        srt = srm = sre = 0; kra_vals = []
        for _, r in sub.iterrows():
            nm = f"{r.get('EmployeeName') or r.get('EmployeeEmail')}<div style='font:600 11px Segoe UI;color:{TEXT_MUTED};'>{r['Team']}</div>"
            kra = float(r.get("KRA_Score", 0)); kra_vals.append(kra); row = [td(nm), td(str(round(kra,1)), right=True, bold=True)]
            rt = rm = rexp = 0
            for c in codes:
                tv = fmt_i(r.get(c, 0)); mv = fmt_i(r.get(f"MTD_{c}", 0)); mt = int(month_targets.get(role, {}).get(c, 0))
                rt+=tv; rm+=mv; rexp+=mt; col_t[c]+=tv; col_m[c]+=mv; col_e[c]+=mt
                row.append(td(cell_t_mtd_pct(tv, mv, mt), right=True))
            srt+=rt; srm+=rm; sre+=rexp; row.append(td(cell_t_mtd_pct(rt, rm, max(rexp,1)), right=True))
            rows.append("<tr>" + "".join(row) + "</tr>")
        if len(rows)>1:
            tot = [td("<b>Total</b>", bold=True), td(f"<b>{round(np.mean(kra_vals),1)}</b>", right=True, bold=True)] + [td(cell_t_mtd_pct(col_t[c], col_m[c], max(col_e[c],1)), right=True) for c in codes] + [td(cell_t_mtd_pct(srt, srm, max(sre,1)), right=True)]
            rows.append("<tr>" + "".join(tot) + "</tr>")
        chunks.append(wrap_card(f"<div style='font:800 15px Segoe UI;color:{TEXT_DARK};margin-bottom:8px;'>Role-wise Work — {role}</div><table role='presentation' width='100%' cellspacing='0' cellpadding='0' style='border-collapse:collapse;border:1px solid {DIVIDER};border-radius:10px;overflow:hidden;'><thead><tr>{''.join(head)}</tr></thead><tbody>{''.join(rows)}</tbody></table>"))
    return "".join(chunks) if chunks else wrap_card(hlabel("ROLE-WISE WORK — ASSIGNEE (ORG)") + f"<div style='color:{TEXT_MUTED}'>No activity.</div>")

def _team_avg_kra_yesterday(df_day_all: pd.DataFrame) -> pd.DataFrame:
    temp = df_day_all.copy(); temp["Team"] = temp.apply(region_key, axis=1)
    temp["KRA_Score"] = pd.to_numeric(temp.get("KRA_Score", 0), errors="coerce").fillna(0.0)
    temp = temp[temp["KRA_Score"]>0]
    return temp.groupby("Team", dropna=True)["KRA_Score"].mean(numeric_only=True).reset_index() if not temp.empty else pd.DataFrame(columns=["Team","KRA_Score"])

def _team_avg_kra_mtd_aw(df_all: pd.DataFrame, work_day: dt.date) -> pd.DataFrame:
    m0 = work_day.replace(day=1); dfm = df_all[(df_all["ReportDate"] >= m0) & (df_all["ReportDate"] <= work_day)].copy()
    if dfm.empty: return pd.DataFrame(columns=["Team","MTD_KRA_AW"])
    dfm["Team"] = dfm.apply(region_key, axis=1); dfm["KRA_Score"] = pd.to_numeric(dfm.get("KRA_Score", 0), errors="coerce").fillna(0.0)
    day = dfm[dfm["KRA_Score"]>0].groupby(["Team", "ReportDate"], dropna=True)["KRA_Score"].agg(avg="mean", cnt="count").reset_index()
    if day.empty: return pd.DataFrame(columns=["Team","MTD_KRA_AW"])
    agg = day.groupby("Team", dropna=True).apply(lambda x: pd.Series({"MTD_KRA_AW": (x["avg"]*x["cnt"]).sum()/x["cnt"].sum()})).reset_index()
    return agg

def build_top_bottom_teams(df_day_all: pd.DataFrame, df_all: pd.DataFrame, work_day: dt.date) -> str:
    y = _team_avg_kra_yesterday(df_day_all); m = _team_avg_kra_mtd_aw(df_all, work_day)
    def tbl(sub, lbl, col, bad=False):
        if sub.empty: return wrap_card(hlabel(lbl) + f"<div style='color:{TEXT_MUTED}'>No data.</div>")
        avg = sub[col].mean(); top = sub[sub[col]>=avg].sort_values(col, ascending=False).head(3); bot = sub[sub[col]<avg].sort_values(col, ascending=True).head(3)
        def rtbl(s, b):
            rows = [f"<tr>{td(f'#{i} {r.Team}', bg=TINT_BAD if b else TINT_GOOD)}{td(str(round(float(r[col]),1)), right=True, bold=True, bg=TINT_BAD if b else TINT_GOOD)}</tr>" for i,(_,r) in enumerate(s.iterrows(), 1)]
            return f"<table role='presentation' width='100%' cellspacing='0' cellpadding='0' style='border-collapse:collapse;border:1px solid {DIVIDER};border-radius:10px;overflow:hidden;'><thead><tr>{th('Team')}{th('Avg', align='right')}</tr></thead><tbody>{''.join(rows)}</tbody></table>"
        return f"<table role='presentation' width='100%'><tr><td style='width:50%;padding-right:6px;vertical-align:top;'>{wrap_card(hlabel('Top 3 '+lbl)+rtbl(top,False))}</td><td style='width:50%;padding-left:6px;vertical-align:top;'>{wrap_card(hlabel('Bottom 3 '+lbl)+rtbl(bot,True))}</td></tr></table>"
    return hlabel("TOP & BOTTOM — TEAMS") + tbl(y, "Yesterday", "KRA_Score") + tbl(m, "MTD (AW)", "MTD_KRA_AW")

def build_org_leaderboard_daily(df_day_all: pd.DataFrame) -> str:
    temp = df_day_all.copy(); temp["Team"] = temp.apply(region_key, axis=1)
    grp = temp[pd.to_numeric(temp.get("KRA_Score", 0), errors="coerce").fillna(0.0)>0].groupby("Team", dropna=True)["KRA_Score"].mean(numeric_only=True).reset_index().sort_values("KRA_Score", ascending=False).head(5)
    rows = [f"<tr>{td(f'#{i} {r.Team}')} {td(str(round(float(r.KRA_Score),1)), right=True, bold=True)}</tr>" for i, (_, r) in enumerate(grp.iterrows(), 1)]
    return wrap_card(hlabel("TEAM LEADERBOARD — AVG KRA (Yesterday, ORG)") + f"<table role='presentation' width='100%' cellspacing='0' cellpadding='0' style='border:1px solid {DIVIDER};border-radius:10px;overflow:hidden;border-collapse:collapse;'><thead><tr>{th('Team')}{th('Avg KRA', align='right')}</tr></thead><tbody>{''.join(rows)}</tbody></table>")

def build_org_leaderboard_mtd(df_all: pd.DataFrame, work_day: dt.date) -> str:
    m = _team_avg_kra_mtd_aw(df_all, work_day)
    if m.empty: return wrap_card(hlabel("TEAM LEADERBOARD — AVG KRA (MTD, ORG)") + f"<div style='color:{TEXT_MUTED}'>No data.</div>")
    m = m.sort_values("MTD_KRA_AW", ascending=False).head(5)
    rows = [f"<tr>{td(f'#{i} {r.Team}')} {td(str(round(float(r.MTD_KRA_AW),1)), right=True, bold=True)}</tr>" for i, (_, r) in enumerate(m.iterrows(), 1)]
    return wrap_card(hlabel("TEAM LEADERBOARD — AVG KRA (MTD, ORG)") + f"<table role='presentation' width='100%' cellspacing='0' cellpadding='0' style='border:1px solid {DIVIDER};border-radius:10px;overflow:hidden;border-collapse:collapse;'><thead><tr>{th('Team')}{th('Avg KRA (MTD, AW)', align='right')}</tr></thead><tbody>{''.join(rows)}</tbody></table>")

_BADGE_KEYS = {"Top Submission": {"subm"}, "Top Total Comments": {"comment"}, "Top Total Offer Follow Ups": {"offer"}, "Top Total Pending from Partner F/U": {"pending"}, "Top Total Status Changed": {"status"}}
def build_org_team_badges_today(df_day_all: pd.DataFrame, labels_all: dict) -> str:
    temp = df_day_all.copy(); temp["Team"] = temp.apply(region_key, axis=1); items = []
    grp_kra = temp[pd.to_numeric(temp.get("KRA_Score", 0), errors="coerce").fillna(0.0)>0].groupby("Team", dropna=True)["KRA_Score"].mean(numeric_only=True).reset_index().sort_values("KRA_Score", ascending=False)
    if not grp_kra.empty: items.append(f"<div style='margin:6px 0;'>{pill('Top KRA', bg='#DBEAFE', fg='#1E3A8A')} — Avg KRA: {round(float(grp_kra.iloc[0].KRA_Score),1)}<br><b>Team:</b> {grp_kra.iloc[0].Team}</div>")
    for title in [t for t in ALLOWED_BADGE_TITLES if t != "Top KRA"]:
        keys = _BADGE_KEYS.get(title, set()); best = None; score = -1
        for col in [c for c in temp.columns if not str(c).startswith("MTD_")]:
            lab = labels_all.get(col, col); pts = sum(1 for k in keys if k in (col+lab).lower())
            if pts > score: score, best = pts, col
        if best and score>0:
            agg = temp.groupby("Team", dropna=True)[best].sum(numeric_only=True).reset_index().sort_values(best, ascending=False)
            if not agg.empty and fmt_i(agg.iloc[0][best])>0: items.append(f"<div style='margin:6px 0;'>{pill(title, bg='#FEF3C7', fg='#92400E')} — {abbrev_label(labels_all.get(best,best))}: {fmt_i(agg.iloc[0][best])}<br><b>Team:</b> {agg.iloc[0].Team}</div>")
    return wrap_card(hlabel("TEAM BADGES (TODAY) — ORG") + ("".join(items) or "<div style='color:{TEXT_MUTED}'>No badges.</div>"))

def build_org_team_badges_mtd(df_day_snapshot: pd.DataFrame, df_all_days: pd.DataFrame, labels_all: dict) -> str:
    temp_all = df_all_days.copy(); temp_all["Team"] = temp_all.apply(region_key, axis=1); items = []
    m = _team_avg_kra_mtd_aw(df_all_days, df_day_snapshot.iloc[0]["ReportDate"]).sort_values("MTD_KRA_AW", ascending=False)
    if not m.empty and float(m.iloc[0]["MTD_KRA_AW"]) > 0: items.append(f"<div style='margin:6px 0;'>{pill('Top KRA (MTD, AW)', bg='#DBEAFE', fg='#1E3A8A')} — Avg KRA: {round(float(m.iloc[0]['MTD_KRA_AW']),1)}<br><b>Team:</b> {m.iloc[0].Team}</div>")
    snap = df_day_snapshot.copy(); snap["Team"] = snap.apply(region_key, axis=1); mtd_cols = [c for c in snap.columns if str(c).startswith("MTD_")]
    for title, keys in _BADGE_KEYS.items():
        best = None; score = -1
        for col in mtd_cols:
            lab = labels_all.get(col[4:], col[4:]); pts = sum(1 for k in keys if k in (col+lab).lower())
            if pts > score: score, best = pts, col
        if best and score>0:
            agg = snap.groupby("Team", dropna=True)[best].sum(numeric_only=True).reset_index().sort_values(best, ascending=False)
            if not agg.empty and fmt_i(agg.iloc[0][best])>0: items.append(f"<div style='margin:6px 0;'>{pill(title+' (MTD)', bg='#E0F2FE', fg='#075985')} — {abbrev_label(labels_all.get(best[4:],best[4:]))}: {fmt_i(agg.iloc[0][best])}<br><b>Team:</b> {agg.iloc[0].Team}</div>")
    return wrap_card(hlabel("TEAM BADGES (MTD) — ORG") + ("".join(items) or "<div style='color:{TEXT_MUTED}'>No data.</div>"))

# ============== PDF & MAIL ============== #
def html_to_pdf(html: str, pdf_path: Path) -> tuple[bool, str]:
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        import pdfkit; config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)
        options = {"enable-local-file-access": None,"quiet":"","page-size":"A4","margin-top":"8mm","margin-bottom":"8mm","margin-left":"8mm","margin-right":"8mm"}
        pdfkit.from_string(html, str(pdf_path), configuration=config, options=options)
        return True, "ok"
    except Exception as e: return False, str(e)

def send_outlook_html_with_attachments(to_addr: str, subject: str, html: str, attachments: list[Path], DRY_RUN=True, FROM_DISPLAY="", sig_img_path: Path | None = None, sig_html: str = ""):
    import pythoncom, win32_utils; pythoncom.CoInitialize()
    try:
        ol = win32_utils.safe_dispatch("Outlook.Application"); ns = ol.GetNamespace("MAPI"); mail = ol.CreateItem(0)
        mail.Subject = subject; mail.HTMLBody = html + (sig_html or ""); mail.To = to_addr
        if FROM_DISPLAY:
            try: mail.SentOnBehalfOfName = FROM_DISPLAY
            except: pass
        for p in attachments:
            if p and Path(p).exists(): mail.Attachments.Add(Source=str(p))
        if DRY_RUN: mail.Display(False)
        else:
            for attempt in range(3):
                try: mail.Send(); break
                except: time.sleep(3)
    finally: pythoncom.CoUninitialize()

# ============== MAIN ============== #
def build_month_targets(role_meas: pd.DataFrame, holidays: set[dt.date], work_day: dt.date) -> dict:
    wd = working_days_in_month(work_day, holidays); out = {}
    for role, part in role_meas[role_meas["IncludeInKRA"]==True].groupby("Role"):
        out[str(role)] = {str(r.MeasureCode): int(round(max(0.0, float(r.get("DailyTarget") or 0.0)) * wd)) for _, r in part.iterrows()}
    return out

def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True); cfg = load_settings(SETTINGS_PATH); team = cfg["team"]; role_kra_codes = cfg["role_kra_codes"]; labels_all = cfg["label_map"]; holidays = cfg["holidays"]; role_meas = cfg["role_meas"]
    df = load_summary(SUMMARY_PATH, SUMMARY_SHEET)
    if df.empty: return
    work_days = resolve_pending_work_days(df['ReportDate'].dropna().tolist(), holidays, latest_working_day, 'exec_digest', MAX_CATCHUP_DAYS)
    roster = team[["EmployeeEmail","EmployeeName","Role","Region","SubRegion","Manager","SendTo"]].copy()
    for work_day in work_days:
        try:
            mt = build_month_targets(role_meas, holidays, work_day); dda = df[df['ReportDate']==work_day].copy(); sr = load_sent_recipients('exec_digest', work_day)
            base = roster.merge(dda, on='EmployeeEmail', how='left', suffixes=('_t',''))
            for c in ['EmployeeName','Role','Region','SubRegion','Manager']: base[c] = base[c].combine_first(base.get(f"{c}_t")).fillna('')
            body = "\n".join([build_availability_matrix_org(base, team, role_kra_codes), build_action_summary_team_org(base, role_kra_codes, labels_all, mt), build_rolewise_work_org(base, role_kra_codes, labels_all, mt), build_top_bottom_teams(base, df, work_day), f"<table role='presentation' width='100%'><tr><td style='width:50%;padding-right:6px;vertical-align:top;'>{build_org_leaderboard_daily(base)}</td><td style='width:50%;padding-left:6px;vertical-align:top;'>{build_org_team_badges_today(base, labels_all)}</td></tr></table>", f"<table role='presentation' width='100%'><tr><td style='width:50%;padding-right:6px;vertical-align:top;'>{build_org_leaderboard_mtd(df, work_day)}</td><td style='width:50%;padding-left:6px;vertical-align:top;'>{build_org_team_badges_mtd(base, df, labels_all)}</td></tr></table>"])
            html = f"<!doctype html><html><body style='margin:0;background:{SURFACE_BG};padding:16px 0;font-family:Segoe UI,Arial;'><table role='presentation' width='100%' style='max-width:980px;margin:0 auto;'><tr><td style='padding:0 16px 14px 16px;'><table role='presentation' width='100%' style='background:{BRAND_PRIMARY};border-radius:14px;'><tr><td style='padding:16px 18px;'><div style='font:800 20px Segoe UI;color:#FFFFFF;'>KC Executive Daily — {work_day.isoformat()} — ORG</div></td></tr></table></td></tr><tr><td style='padding:0 16px 10px 16px;'>{body}</td></tr></table></body></html>"
            hp = OUT_DIR / f"Exec_Digest_{work_day.isoformat()}.html"; hp.write_text(html, encoding='utf-8')
            pp = OUT_DIR / f"Exec_Digest_{work_day.isoformat()}.pdf"; okp, _ = html_to_pdf(html, pp)
            att = [pp] if okp else []
            for ex in EXEC_RECIPIENTS:
                if ex.lower() in sr: continue
                send_outlook_html_with_attachments(ex, f"{SUBJECT_PREFIX} | {work_day.isoformat()}", html, att, DRY_RUN, FROM_DISPLAY, SIG_IMAGE_PATH, SIG_HTML)
                append_sent_detail('exec_digest', work_day, ex.lower(), 'dry_run' if DRY_RUN else 'ok')
            append_sent_log('exec_digest', work_day, 'dry_run' if DRY_RUN else 'ok')
        except Exception as e:
            append_sent_log('exec_digest', work_day, 'fail')
            traceback.print_exc()

if __name__ == '__main__':
    main()
