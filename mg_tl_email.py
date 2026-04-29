# -*- coding: utf-8 -*-
"""
KC Manager Daily — TL/MG Roll-up (v3.1)
- Action Summary — Team is now a strict aggregation of the same member rows shown in Role-wise Work
  (same 'available yesterday' filter, same per-member month-targets -> consistent %)
- From address set to 'gsakhare@kcoverseas.com' via Outlook (SentOnBehalfOfName)
- Signature block appended with optional inline CID image 'kc_sig'
- Added monthly org view:
  • TEAM LEADERBOARD — Avg KRA (MTD, ORG)  -> uses per-day KRA across the month up to work_day
  • TEAM BADGES (MTD) — ORG (Team-wise)    -> based on summed MTD_* measures at work_day
- All prior layout/colors unchanged; uniform two-line metric cells kept (T/MTD + % of Month with bar)
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

# =========================== CONFIG =========================== #
SUMMARY_PATH  = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\raw\summary_all_days.xlsx")
SUMMARY_SHEET = "summary_daily_all"
SETTINGS_PATH = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\setting\wd_settings.xlsx")
OUT_DIR       = SUMMARY_PATH.parent / "manager_digest_previews"

# Email transport
DRY_RUN    = False
SEND_VIA   = "outlook"  # 'outlook' or 'smtp'
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT   = 587
SMTP_USERNAME = ""
SMTP_PASSWORD = ""
MAX_CATCHUP_DAYS = 7  # safety cap when catching up on missed days
RETRY_ON_ERROR = 2

# From / Sender (Outlook)
FROM_DISPLAY  = "gsakhare@kcoverseas.com"
SUBJECT_PREFIX = "KC Manager Daily"

# Signature (HTML with optional CID image 'kc_sig')
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

# PDF (optional)
WKHTMLTOPDF_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
PDF_ENGINE = "wkhtmltopdf"
_PDF_WARNED = False

# Holidays logic
IST = dt.timezone(dt.timedelta(hours=5, minutes=30))

# =================== BRAND / UI =================== #
BRAND_PRIMARY    = "#0B1220"
BRAND_ACCENT     = "#2563EB"
BRAND_SECONDARY  = "#F59E0B"
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

# ====================== LOADERS ====================== #
def load_settings(settings_path: Path):
    xls = pd.read_excel(settings_path, sheet_name=None)

    team = xls.get("Team")
    if team is None:
        raise ValueError("wd_settings.xlsx missing 'Team' sheet.")
    for c in ["EmployeeEmail","EmployeeName","Role","Region","SubRegion","Manager","Include","SendTo"]:
        if c not in team.columns: team[c] = ""
    team["EmployeeEmail"] = team["EmployeeEmail"].astype(str).str.strip().str.lower()
    # DEDUPLICATE TEAM ROSTER: Ensure each email only appears once (prevents repetitive rows in reports)
    team = team.drop_duplicates(subset=["EmployeeEmail"], keep="first")
    
    for c in ["Region","SubRegion","Role","Manager","EmployeeName","SendTo"]:
        team[c] = team[c].astype(str).str.strip()

    measures = xls.get("Measures")
    if measures is None:
        raise ValueError("wd_settings.xlsx missing 'Measures' sheet.")
    measures["MeasureCode"]  = measures["MeasureCode"].astype(str).str.strip()
    measures["MeasureLabel"] = measures["MeasureLabel"].astype(str).str.strip()
    labels_all = {r["MeasureCode"]: (r["MeasureLabel"] or r["MeasureCode"]) for _, r in measures.iterrows()}

    role_meas = xls.get("RoleMeasures")
    if role_meas is None:
        raise ValueError("wd_settings.xlsx missing 'RoleMeasures' sheet.")
    for c in ["Role","MeasureCode","Weight","DailyTarget","MinFloor","Cap","IncludeInKRA","DisplayOrder"]:
        if c not in role_meas.columns: role_meas[c] = ""
    role_meas["Role"] = role_meas["Role"].astype(str).str.strip()
    role_meas["MeasureCode"] = role_meas["MeasureCode"].astype(str).str.strip()
    role_meas["IncludeInKRA"] = role_meas["IncludeInKRA"].astype(str).str.upper().map({"Y":True,"YES":True}).fillna(False)
    role_meas["DailyTarget"]  = pd.to_numeric(role_meas.get("DailyTarget", 0), errors="coerce")
    if "DisplayOrder" in role_meas.columns:
        role_meas["DisplayOrder"] = pd.to_numeric(role_meas["DisplayOrder"], errors="coerce")
    else:
        role_meas["DisplayOrder"] = np.nan

    role_kra_codes = {}
    for role in sorted(role_meas["Role"].dropna().unique()):
        sub = role_meas[(role_meas["Role"]==role) & (role_meas["IncludeInKRA"]==True)].copy()
        if sub.empty:
            role_kra_codes[role] = []
            continue
        sub["_row"] = np.arange(len(sub))
        sub["__has"] = sub["DisplayOrder"].notna().astype(int)
        sub = sub.sort_values(["__has","DisplayOrder","_row"], ascending=[False,True,True])
        role_kra_codes[role] = [str(x) for x in sub["MeasureCode"].dropna().tolist()]

    holidays = set()
    hol = xls.get("Holidays")
    if hol is not None:
        for col in ["Date","HolidayDate","Holiday","Dt"]:
            if col in hol.columns:
                h = pd.to_datetime(hol[col], errors="coerce").dt.date.dropna()
                holidays = set(h.tolist())
                break

    return team, role_kra_codes, labels_all, holidays, role_meas

def load_summary(path: Path, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet)
    df["ReportDate"] = pd.to_datetime(df["ReportDate"]).dt.date
    df["EmployeeEmail"] = df["EmployeeEmail"].astype(str).str.strip().str.lower()
    for c in ["EmployeeName","Role","Region","SubRegion","Manager"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).str.strip()
            
    # DEDUPLICATE: Ensure each employee has only one record per day to prevent inflated totals
    df = df.drop_duplicates(subset=["ReportDate", "EmployeeEmail"], keep="first")
    return df

# ============== WORKING DAY HELPERS & MONTH TARGETS ============== #
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
        if is_sunday(d): 
            continue
        if d in holidays:
            continue
        cnt += 1
    return max(cnt, 1)

def build_month_targets(role_meas: pd.DataFrame, holidays: set[dt.date], work_day: dt.date) -> dict[str, dict[str, int]]:
    wd = working_days_in_month(work_day, holidays)
    out: dict[str, dict[str, int]] = {}
    sub = role_meas[role_meas["IncludeInKRA"]==True].copy()
    for role in sorted(sub["Role"].dropna().unique()):
        out[role] = {}
        part = sub[sub["Role"]==role].copy()
        for _, r in part.iterrows():
            code = str(r["MeasureCode"])
            daily = float(r.get("DailyTarget") or 0.0)
            out[role][code] = int(round(max(0.0, daily) * wd))
    return out

# ============== SMALL HTML HELPERS ============== #
BR_SMALL = "font:600 11px Segoe UI,Arial;color:#6B7280;"
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

def small(txt: str) -> str:
    return f"""<span style="{BR_SMALL}">{txt}</span>"""

def fmt_i(x) -> int:
    try: return int(round(float(x)))
    except: return 0

def region_key(row) -> str:
    rg = str(row.get("Region","") or "").strip()
    sr = str(row.get("SubRegion","") or "").strip()
    return f"{rg} — {sr}" if sr else rg

def has_any_kra_today(row, role_kra_codes: dict) -> bool:
    role = str(row.get("Role","") or "")
    codes = role_kra_codes.get(role, [])
    if not codes: return False
    for code in codes:
        if fmt_i(row.get(code, 0)) > 0:
            return True
    return False

def collect_manager_teams(team_df: pd.DataFrame, manager_email: str) -> list[str]:
    my = team_df[team_df["Manager"].str.lower()==str(manager_email).lower()]
    keys = sorted(set([region_key(r) for _, r in my.iterrows() if (r.get("Region","") or r.get("SubRegion",""))]))
    return keys

# ============== ABBREVIATIONS (improved) ============== #
_ABBR_MAP = {
    "Submission": "Subm",
    "Assessed": "Assess",
    "Total Comments": "Cmt(T)",
    "Unique Comments": "Cmt(U)",
    "Total Offer Follow Ups": "OFU(T)",
    "Unique Offer Follow Up": "OFU(U)",
    "Total Pending from Partner F/U": "PFU(T)",
    "Unique Pending from Partner F/U": "PFU(U)",
    "Total Status Changed": "Stat(T)",
    "Unique Status Changed": "Stat(U)",
    "Month-to-Date": "MTD",
}
def abbrev_label(lbl: str) -> str:
    s = (lbl or "").strip()
    for k, v in _ABBR_MAP.items():
        s = re.sub(k, v, s, flags=re.I)
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

# ============== SECTION 1: TEAM AVAILABILITY MATRIX (unchanged logic) ============== #
def build_availability_matrix(df_mgr_day: pd.DataFrame, team_df: pd.DataFrame, role_kra_codes: dict) -> str:
    part = df_mgr_day.copy()
    part["Team"] = part.apply(region_key, axis=1)
    part["Available"] = part.apply(lambda r: has_any_kra_today(r, role_kra_codes), axis=1)

    roles_all = sorted(set(part["Role"].dropna().astype(str)))
    roster_sub = team_df[team_df["Manager"].str.lower()==part.iloc[0]["Manager"].lower()].copy() if not part.empty else team_df.iloc[0:0]
    roster_sub["Team"] = roster_sub.apply(region_key, axis=1)

    teams = sorted(set(roster_sub["Team"].dropna().astype(str)))
    count_map = {}
    for t in teams:
        for role in roles_all:
            total = int(((roster_sub["Team"]==t) & (roster_sub["Role"]==role)).sum())
            avail = int(((part["Team"]==t) & (part["Role"]==role) & (part["Available"]==True)).sum())
            count_map[(t, role)] = (avail, total)

    roles = [r for r in roles_all if sum(count_map[(t, r)][1] for t in teams) > 0]

    col_tot_av = {role: sum(count_map[(t, role)][0] for t in teams) for role in roles}
    col_tot_to = {role: sum(count_map[(t, role)][1] for t in teams) for role in roles}

    headers = th("Team (Region — SubRegion)") + "".join([th(role, align="right") for role in roles]) + th("Row Total", align="right")
    body = []
    for t in teams:
        cells_av = [td(f"<b>{t}</b>")]
        row_av = 0
        for role in roles:
            av, tot = count_map[(t, role)]
            row_av += av
            cells_av.append(td(str(av), right=True, bold=True))
        cells_av.append(td(str(row_av), right=True, bold=True))

        cells_to = [td("Total", muted=True)]
        row_to = 0
        for role in roles:
            av, tot = count_map[(t, role)]
            row_to += tot
            cells_to.append(td(small(str(tot)), right=True, muted=True))
        cells_to.append(td(small(str(row_to)), right=True, muted=True))

        body.append("<tr>" + "".join(cells_av) + "</tr>")
        body.append("<tr>" + "".join(cells_to) + "</tr>")

    gt_rows = ""
    if len(teams) > 1:
        gt_av = [td("<b>Grand Total — Available</b>")]
        av_sum = 0
        for role in roles:
            v = col_tot_av[role]; av_sum += v
            gt_av.append(td(str(v), right=True, bold=True))
        gt_av.append(td(str(av_sum), right=True, bold=True))

        gt_to = [td("<b>Grand Total — Total</b>")]
        to_sum = 0
        for role in roles:
            v = col_tot_to[role]; to_sum += v
            gt_to.append(td(str(v), right=True))
        gt_to.append(td(str(to_sum), right=True))

        gt_rows = f"<tr>{''.join(gt_av)}</tr><tr>{''.join(gt_to)}</tr>"

    hint = f"<div style='margin-top:6px;color:{TEXT_MUTED};font:600 12px Segoe UI,Arial;'>Hint: <b>Available</b> means any KRA measure &gt; 0 yesterday.</div>"

    table = f"""
      {hlabel("YESTERDAY SUMMARY — TEAM AVAILABILITY MATRIX")}{hint}
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid {DIVIDER};border-radius:10px;overflow:hidden;">
        <thead>
          <tr>{headers}</tr>
        </thead>
        <tbody>
          {''.join(body)}
          {gt_rows}
        </tbody>
      </table>
    """
    return wrap_card(table)

# ============== SECTION 1B: ACTION SUMMARY — TEAM (CONSISTENT AGG FROM MEMBER ROWS) ============== #
def build_action_summary_team(df_mgr_day: pd.DataFrame, role_kra_codes: dict, labels_all: dict,
                              month_targets: dict[str, dict[str, int]]) -> str:
    df = df_mgr_day.copy()
    df["Team"] = df.apply(region_key, axis=1)

    # Only members "available yesterday" (any KRA>0) — same as assignee tables
    def is_active_row(r):
        role = str(r.get("Role","") or "")
        codes = role_kra_codes.get(role, [])
        return any(fmt_i(r.get(c,0))>0 for c in codes)
    df["__active"] = df.apply(is_active_row, axis=1)
    df = df[df["__active"]==True].copy()

    teams = sorted(set(df["Team"].dropna().astype(str)))

    # union of KRA codes present under this manager (from the active set)
    codes_union = []
    roles_present = sorted(set(df["Role"].dropna().astype(str)))
    for role in roles_present:
        for c in role_kra_codes.get(role, []):
            if c not in codes_union:
                codes_union.append(c)
    if not codes_union:
        return wrap_card(hlabel("ACTION SUMMARY — TEAM") + f"<div style='color:{TEXT_MUTED}'>No KRA measures yesterday.</div>")

    head_cells = [th("Team (Region — SubRegion)")]
    for code in codes_union:
        lab = abbrev_label(labels_all.get(code, code))
        head_cells.append(th(f"{lab}<br><span style='font-weight:600'>T / MTD</span><br><span style='font-weight:600'>% Month</span>", align="right"))
    head_cells.append(th("Row Total<br><span style='font-weight:600'>T / MTD</span><br><span style='font-weight:600'>% Month</span>", align="right"))

    body = []
    # aggregate strictly from member rows:
    #   T_team = sum(member T)
    #   M_team = sum(member MTD)
    #   E_team = sum(member month-targets by role
    col_tot_today = {c: 0 for c in codes_union}
    col_tot_mtd   = {c: 0 for c in codes_union}
    col_tot_exp   = {c: 0 for c in codes_union}
    grand_t = grand_m = grand_e = 0

    for t in teams:
        sub = df[df["Team"]==t].copy()

        row_cells = [td(f"<b>{t}</b>")]
        row_t = row_m = row_e = 0

        # sum per-code across members (explicitly member-based, no recalculation)
        for c in codes_union:
            v_t = int(pd.to_numeric(sub.get(c, 0), errors="coerce").fillna(0).sum()) if c in sub.columns else 0
            v_m = int(pd.to_numeric(sub.get(f"MTD_{c}", 0), errors="coerce").fillna(0).sum()) if f"MTD_{c}" in sub.columns else 0
            # E_team = sum(member month target for that measure)
            v_e = 0
            for _, r in sub.iterrows():
                role = str(r.get("Role","") or "")
                v_e += int(month_targets.get(role, {}).get(c, 0))

            row_t += v_t; row_m += v_m; row_e += v_e
            col_tot_today[c] += v_t; col_tot_mtd[c] += v_m; col_tot_exp[c] += v_e
            row_cells.append(td(cell_t_mtd_pct(v_t, v_m, max(v_e,1)), right=True))

        grand_t += row_t; grand_m += row_m; grand_e += row_e
        row_cells.append(td(cell_t_mtd_pct(row_t, row_m, max(row_e,1)), right=True))
        body.append("<tr>" + "".join(row_cells) + "</tr>")

    if len(teams) > 1:
        tot_cells = [td("<b>Grand Total</b>", bold=True)]
        for c in codes_union:
            tot_cells.append(td(cell_t_mtd_pct(col_tot_today[c], col_tot_mtd[c], max(col_tot_exp[c],1)), right=True))
        tot_cells.append(td(cell_t_mtd_pct(grand_t, grand_m, max(grand_e,1)), right=True))
        body.append("<tr>" + "".join(tot_cells) + "</tr>")

    tbl = f"""
      {hlabel("ACTION SUMMARY — TEAM")}
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid {DIVIDER};border-radius:10px;overflow:hidden;">
        <thead><tr>{''.join(head_cells)}</tr></thead>
        <tbody>{''.join(body)}</tbody>
      </table>
    """
    return wrap_card(tbl)

# ============== SECTION 2: ROLE-WISE WORK — ASSIGNEE (unchanged visuals) ============== #
def build_rolewise_work(df_mgr_day: pd.DataFrame, role_kra_codes: dict, labels_all: dict,
                        month_targets: dict[str, dict[str, int]]) -> str:
    df = df_mgr_day.copy()
    df["Team"] = df.apply(region_key, axis=1)
    chunks = []

    roles_all = [r for r in sorted(set(df["Role"].dropna().astype(str))) if role_kra_codes.get(r, [])]
    if not roles_all:
        return wrap_card(hlabel("Role-wise Work — Assignee") + f"<div style='color:{TEXT_MUTED}'>No activity yesterday.</div>")

    for role in roles_all:
        codes = role_kra_codes.get(role, [])
        df_role = df[df["Role"]==role].copy()

        role_title = f"""<div style="font:800 15px Segoe UI,Arial;color:{TEXT_DARK};margin-bottom:8px;">Role-wise Work — {role}</div>"""
        role_chunk = [role_title]

        teams = sorted(set(df_role["Team"].dropna().astype(str)))
        for tix, team in enumerate(teams):
            df_team_role = df_role[df_role["Team"]==team].copy()
            df_team_role["__active"] = df_team_role.apply(lambda r: any(fmt_i(r.get(c,0))>0 for c in codes), axis=1)
            df_team_role = df_team_role[df_team_role["__active"]==True].copy()
            if df_team_role.empty:
                continue

            if tix > 0:
                role_chunk.append(f"<div style='height:1px;background:{DIVIDER};margin:8px 0;'></div>")

            team_hdr = f"<div style='font:800 12px Segoe UI,Arial;color:{TEXT_MUTED};margin:2px 0 4px 0;'>{team}</div>"

            head = [th("Member"), th("KRA", align="right")]
            for c in codes:
                lab = abbrev_label(labels_all.get(c, c))
                head.append(th(f"{lab}<br><span style='font-weight:600'>T / MTD</span><br><span style='font-weight:600'>% Month</span>", align="right"))
            head.append(th("Row Total<br><span style='font-weight:600'>T / MTD</span><br><span style='font-weight:600'>% Month</span>", align="right"))

            rows_html = []
            col_t = {c:0 for c in codes}; col_m = {c:0 for c in codes}; col_e = {c:0 for c in codes}
            sum_row_t = 0; sum_row_m = 0; sum_row_e = 0
            kravalues = []

            for _, r in df_team_role.iterrows():
                name = r.get("EmployeeName") or r.get("EmployeeEmail")
                kra = float(pd.to_numeric(r.get("KRA_Score", 0), errors="coerce") or 0)
                kravalues.append(kra)
                row = [td(name)]
                row.append(td(str(round(kra,1)), right=True, bold=True))
                rt = 0; rm = 0; rexp = 0
                for c in codes:
                    t_val = fmt_i(r.get(c, 0))
                    m_val = fmt_i(r.get(f"MTD_{c}", 0))
                    mtarget = int(month_targets.get(role, {}).get(c, 0))
                    rt += t_val; rm += m_val; rexp += mtarget
                    col_t[c] += t_val; col_m[c] += m_val; col_e[c] += mtarget
                    row.append(td(cell_t_mtd_pct(t_val, m_val, mtarget), right=True))
                sum_row_t += rt; sum_row_m += rm; sum_row_e += rexp
                row.append(td(cell_t_mtd_pct(rt, rm, max(rexp,1)), right=True))
                rows_html.append("<tr>" + "".join(row) + "</tr>")

            if len(rows_html) > 1:
                tot = [td("<b>Total</b>", bold=True), td(f"<b>{round(np.mean(kravalues),1)}</b>", right=True, bold=True)]
                for c in codes:
                    tot.append(td(cell_t_mtd_pct(col_t[c], col_m[c], max(col_e[c],1)), right=True))
                tot.append(td(cell_t_mtd_pct(sum_row_t, sum_row_m, max(sum_row_e,1)), right=True))
                rows_html.append("<tr>" + "".join(tot) + "</tr>")

            table = f"""
              {team_hdr}
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:6px;border-collapse:collapse;border:1px solid {DIVIDER};border-radius:10px;overflow:hidden;">
                <thead><tr>{''.join(head)}</tr></thead>
                <tbody>{''.join(rows_html)}</tbody>
              </table>
            """
            role_chunk.append(table)

        if len(role_chunk) > 1:
            chunks.append(wrap_card("".join(role_chunk)))

    return "".join(chunks) if chunks else wrap_card(hlabel("Role-wise Work — Assignee") + f"<div style='color:{TEXT_MUTED}'>No activity yesterday.</div>")

# ============== SECTION 3: TOP / BOTTOM 3 (same threshold logic) ============== #
def build_top_bottom(df_mgr_day: pd.DataFrame) -> str:
    df = df_mgr_day.copy()
    df["Team"] = df.apply(region_key, axis=1)
    df["KRA_Score"] = pd.to_numeric(df.get("KRA_Score", 0), errors="coerce").fillna(0.0)
    df = df[df["KRA_Score"]>0].copy()
    if df.empty:
        return wrap_card(hlabel("Top 3 & Bottom 3") + f"<div style='color:{TEXT_MUTED}'>No qualifying members.</div>")
    avg = float(df["KRA_Score"].mean())

    top = df[df["KRA_Score"] >= avg].sort_values("KRA_Score", ascending=False).head(3)
    bot = df[df["KRA_Score"] <  avg].sort_values("KRA_Score", ascending=True).head(3)

    def table_for(sub: pd.DataFrame, bad: bool=False):
        if sub.empty:
            return f"<div style='color:{TEXT_MUTED}'>No members.</div>"
        rows=[]
        for _, r in sub.iterrows():
            tint = TINT_BAD if bad else TINT_GOOD
            rows.append(
                "<tr>" +
                td(r.get('EmployeeName') or r.get('EmployeeEmail'), bg=tint) +
                td(r['Team'], bg=tint) +
                td(str(round(float(r['KRA_Score']),1)), right=True, bold=True, bg=tint) +
                "</tr>"
            )
        head = f"<tr>{th('Member')}{th('Team')}{th('KRA', align='right')}</tr>"
        return f"""<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid {DIVIDER};border-radius:10px;overflow:hidden;"><thead>{head}</thead><tbody>{''.join(rows)}</tbody></table>"""

    left  = hlabel(f"Top 3 by KRA — Team (≥ Avg {round(avg,1)})") + table_for(top, bad=False)
    right = hlabel(f"Bottom 3 by KRA — Team (< Avg {round(avg,1)})") + table_for(bot, bad=True)

    html = f"""
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
        <tr>
          <td style="width:50%;padding-right:6px;vertical-align:top;">{wrap_card(left)}</td>
          <td style="width:50%;padding-left:6px;vertical-align:top;">{wrap_card(right)}</td>
        </tr>
      </table>
    """
    return html

# ============== SECTION 4: TEAM BEST vs ORG BEST (unchanged) ============== #
def build_team_vs_org_best(df_mgr_day: pd.DataFrame, df_day_all: pd.DataFrame, role_kra_codes: dict, labels_all: dict) -> str:
    roles = sorted(set(df_mgr_day["Role"].dropna().astype(str)))
    chunks = []
    for role in roles:
        codes = role_kra_codes.get(role, [])
        if not codes:
            continue
        sub_team = df_mgr_day[df_mgr_day["Role"]==role].copy()
        sub_team = sub_team[pd.to_numeric(sub_team.get("KRA_Score",0), errors="coerce").fillna(0.0)>0]
        if sub_team.empty:
            continue
        team_best = sub_team.sort_values("KRA_Score", ascending=False).iloc[0]

        sub_org = df_day_all[df_day_all["Role"]==role].copy()
        sub_org = sub_org[pd.to_numeric(sub_org.get("KRA_Score",0), errors="coerce").fillna(0.0)>0]
        if sub_org.empty:
            continue
        org_best = sub_org.sort_values("KRA_Score", ascending=False).iloc[0]

        def comp_line(row):
            parts=[]
            for c in codes:
                lab = abbrev_label(labels_all.get(c, c))
                parts.append(f"{lab}: {fmt_i(row.get(c,0))}")
            return " · ".join(parts)

        body = f"""
          <div style="margin-bottom:6px;font:600 13px Segoe UI,Arial;color:{TEXT_DARK}">
            Team: <b>{team_best.get('EmployeeName') or team_best.get('EmployeeEmail')}</b> ({round(float(team_best.get('KRA_Score',0)),1)})
            &nbsp;&nbsp; vs &nbsp;&nbsp;
            Org: <b>{org_best.get('EmployeeName') or org_best.get('EmployeeEmail')}</b> ({round(float(org_best.get('KRA_Score',0)),1)})
          </div>
          <div style="color:{TEXT_MUTED};font:600 12px Segoe UI,Arial;">
            {comp_line(team_best)} &nbsp; | &nbsp; {comp_line(org_best)}
          </div>
        """
        chunks.append(wrap_card(hlabel(f"Team Best vs Org Best — {role}") + body))

    return "".join(chunks) if chunks else wrap_card(hlabel("Team Best vs Org Best") + f"<div style='color:{TEXT_MUTED}'>No qualifying roles.</div>")

# ============== SECTION 5A: ORG LEADERBOARD (Yesterday) ============== #
def build_org_leaderboard_daily(df_day_all: pd.DataFrame) -> str:
    temp = df_day_all.copy()
    temp["Team"] = temp.apply(region_key, axis=1)
    grp = temp.groupby("Team", dropna=True)["KRA_Score"].mean(numeric_only=True).reset_index()
    grp["KRA_Score"] = pd.to_numeric(grp["KRA_Score"], errors="coerce").fillna(0.0)
    grp = grp.sort_values("KRA_Score", ascending=False).head(5)
    rows=[]
    for i, (_, r) in enumerate(grp.iterrows(), start=1):
        rows.append(f"<tr>{td(f'#{i} {r.Team}')}{td(str(round(float(r.KRA_Score),1)), right=True, bold=True)}</tr>")
    head = f"<tr>{th('Team (Region — SubRegion)')}{th('Avg KRA', align='right')}</tr>"
    table = f"""<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border:1px solid {DIVIDER};border-radius:10px;overflow:hidden;border-collapse:collapse;"><thead>{head}</thead><tbody>{''.join(rows)}</tbody></table>"""
    return wrap_card(hlabel("TEAM LEADERBOARD — AVG KRA (Yesterday, ORG)") + table)

# ============== SECTION 5B: ORG LEADERBOARD (MTD Avg KRA) ============== #
def build_org_leaderboard_mtd(df_all: pd.DataFrame, work_day: dt.date) -> str:
    """
    TEAM LEADERBOARD — AVG KRA (MTD, ORG) availability-weighted.
    A member is 'available' on a day if KRA_Score > 0.
    Team-day contributes with weight = active_count that day.
    """
    m0 = work_day.replace(day=1)
    dfm = df_all[(df_all["ReportDate"] >= m0) & (df_all["ReportDate"] <= work_day)].copy()
    if dfm.empty:
        return wrap_card(hlabel("TEAM LEADERBOARD — AVG KRA (MTD, ORG)") + f"<div style='color:{TEXT_MUTED}'>No data.</div>")

    dfm["Team"] = dfm.apply(region_key, axis=1)
    dfm["KRA_Score"] = pd.to_numeric(dfm.get("KRA_Score", 0), errors="coerce").fillna(0.0)
    dfm["__active"] = dfm["KRA_Score"] > 0

    # Per Team–Day: active_count & day_kra_avg (active only)
    day = (
        dfm[dfm["__active"]]
        .groupby(["Team", "ReportDate"], dropna=True)["KRA_Score"]
        .agg(day_kra_avg="mean", active_count="count")
        .reset_index()
    )
    if day.empty:
        return wrap_card(hlabel("TEAM LEADERBOARD — AVG KRA (MTD, ORG)") + f"<div style='color:{TEXT_MUTED}'>No available members this month.</div>")

    # Availability-weighted mean across the month
    day["w_sum"] = day["day_kra_avg"] * day["active_count"]
    agg = (
        day.groupby("Team", dropna=True)[["w_sum", "active_count"]]
        .sum()
        .reset_index()
    )
    agg["MTD_KRA_AW"] = np.where(agg["active_count"] > 0, agg["w_sum"] / agg["active_count"], 0.0)
    agg = agg.sort_values("MTD_KRA_AW", ascending=False).head(5)

    rows = []
    for i, (_, r) in enumerate(agg.iterrows(), start=1):
        rows.append(f"<tr>{td(f'#{i} {r.Team}')} {td(str(round(float(r.MTD_KRA_AW),1)), right=True, bold=True)}</tr>")

    head = f"<tr>{th('Team (Region — SubRegion)')}{th('Avg KRA (MTD, AW)', align='right')}</tr>"
    table = f"""<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border:1px solid {DIVIDER};border-radius:10px;overflow:hidden;border-collapse:collapse;"><thead>{head}</thead><tbody>{''.join(rows)}</tbody></table>"""
    return wrap_card(hlabel("TEAM LEADERBOARD — AVG KRA (MTD, ORG)") + table)

# ============== BADGES HELPERS (title->measure) ============== #
_BADGE_KEYS = {
    "Top Submission": {"subm", "submission"},
    "Top Total Comments": {"comment", "cmt"},
    "Top Total Offer Follow Ups": {"offer", "follow"},
    "Top Total Pending from Partner F/U": {"pending", "partner", "f/u", "fu"},
    "Top Total Status Changed": {"status", "change"},
}
def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+"," ",str(s or "").lower()).strip()

def resolve_measure_for_badge(labels_all: dict, columns: list[str], title: str) -> str | None:
    keys = _BADGE_KEYS.get(title, set())
    if not keys: return None
    best = None; score = -1
    for col in columns:
        if col.startswith("MTD_"): continue
        base = col
        lab  = labels_all.get(base, base)
        tokens = _norm(base).split() + _norm(lab).split()
        points = sum(1 for k in keys if any(k in t for t in tokens))
        if points > score:
            score = points; best = col
    return best if score>0 else None

# ============== SECTION 6A: TEAM BADGES (TODAY) — ORG (Team-wise) ============== #
def build_org_team_badges_today(df_day_all: pd.DataFrame, labels_all: dict) -> str:
    temp = df_day_all.copy()
    temp["Team"] = temp.apply(region_key, axis=1)

    measure_today_cols = [c for c in temp.columns if (not str(c).startswith("MTD_")) and c not in ["KRA_Score","ReportDate","Role","Region","SubRegion","EmployeeEmail","EmployeeName","Manager","SendTo"]]

    grp_kra = temp.groupby("Team", dropna=True)["KRA_Score"].mean(numeric_only=True).reset_index()
    grp_kra = grp_kra.sort_values("KRA_Score", ascending=False)
    items=[]
    if not grp_kra.empty:
        items.append(f"<div style='margin:6px 0;'>{pill('Top KRA', bg='#DBEAFE', fg='#1E3A8A')} {small(f'— Avg KRA: {round(float(grp_kra.iloc[0].KRA_Score),1)}')}<br><b>Team:</b> {grp_kra.iloc[0].Team}</div>")

    for title in [t for t in ALLOWED_BADGE_TITLES if t != "Top KRA"]:
        col = resolve_measure_for_badge(labels_all, measure_today_cols, title)
        if not col or col not in temp.columns:
            continue
        agg = temp.groupby("Team", dropna=True)[col].sum(numeric_only=True).reset_index()
        agg = agg.sort_values(col, ascending=False)
        if agg.empty or fmt_i(agg.iloc[0][col])<=0:
            continue
        label_short = abbrev_label(labels_all.get(col, col))
        items.append(
            f"<div style='margin:6px 0;'>{pill(title, bg='#FEF3C7', fg='#92400E')} {small(f'— {label_short}: {fmt_i(agg.iloc[0][col])}')}<br><b>Team:</b> {agg.iloc[0].Team}</div>"
        )

    body = "".join(items) if items else f"<div style='color:{TEXT_MUTED}'>No team badges today.</div>"
    return wrap_card(hlabel("TEAM BADGES (TODAY) — ORG (Team-wise)") + body)

# ============== SECTION 6B: TEAM BADGES (MTD) — ORG (Team-wise) ============== #
def build_org_team_badges_mtd(df_all: pd.DataFrame, labels_all: dict) -> str:
    """
    TEAM BADGES (MTD) — ORG (Team-wise)
    - Top KRA (MTD, AW): availability-weighted MTD KRA as in leaderboard_mtd
    - Other badges: based on summed MTD_* measures (as before)
    """
    temp = df_all.copy()
    temp["Team"] = temp.apply(region_key, axis=1)
    temp["KRA_Score"] = pd.to_numeric(temp.get("KRA_Score", 0), errors="coerce").fillna(0.0)
    temp["__active"] = temp["KRA_Score"] > 0

    # ---- Top KRA (MTD, AW)
    day = (
        temp[temp["__active"]]
        .groupby(["Team", "ReportDate"], dropna=True)["KRA_Score"]
        .agg(day_kra_avg="mean", active_count="count")
        .reset_index()
    )
    items = []
    if not day.empty:
        day["w_sum"] = day["day_kra_avg"] * day["active_count"]
        agg = day.groupby("Team", dropna=True)[["w_sum","active_count"]].sum().reset_index()
        agg["MTD_KRA_AW"] = np.where(agg["active_count"]>0, agg["w_sum"]/agg["active_count"], 0.0)
        agg = agg.sort_values("MTD_KRA_AW", ascending=False)
        if not agg.empty and float(agg.iloc[0]["MTD_KRA_AW"]) > 0:
            items.append(
                f"<div style='margin:6px 0;'>{pill('Top KRA (MTD, AW)', bg='#DBEAFE', fg='#1E3A8A')} "
                f"{small(f'— Avg KRA: {round(float(agg.iloc[0].MTD_KRA_AW),1)}')}<br><b>Team:</b> {agg.iloc[0].Team}</div>"
            )

    # ---- Other MTD badges (sum of MTD_* measures)
    mtd_cols = [c for c in temp.columns if str(c).startswith("MTD_")]
    for title in [t for t in ALLOWED_BADGE_TITLES if t != "Top KRA"]:
        keys = _BADGE_KEYS.get(title, set())
        best = None; score = -1
        for col in mtd_cols:
            base = col[4:]
            lab  = labels_all.get(base, base)
            tokens = _norm(base).split() + _norm(lab).split()
            pts = sum(1 for k in keys if any(k in t for t in tokens))
            if pts > score:
                score = pts; best = col
        if not best or score <= 0:
            continue
        agg2 = temp.groupby("Team", dropna=True)[best].sum(numeric_only=True).reset_index()
        agg2 = agg2.sort_values(best, ascending=False)
        if agg2.empty or fmt_i(agg2.iloc[0][best]) <= 0:
            continue
        label_short = abbrev_label(labels_all.get(best[4:], best[4:]))
        items.append(
            f"<div style='margin:6px 0;'>{pill(title+' (MTD)', bg='#E0F2FE', fg='#075985')} "
            f"{small(f'— {label_short}: {fmt_i(agg2.iloc[0][best])}')}<br><b>Team:</b> {agg2.iloc[0].Team}</div>"
        )

    body = "".join(items) if items else f"<div style='color:{TEXT_MUTED}'>No team badges (MTD) yet.</div>"
    return wrap_card(hlabel("TEAM BADGES (MTD) — ORG (Team-wise)") + body)

# ============== ABBREVIATIONS FOOTER ============== #
def build_abbrev_footer() -> str:
    left = """
      <div style="padding:4px 0;"><b>KRA</b> — Key Result Area</div>
      <div style="padding:4px 0;"><b>MTD</b> — Month-to-Date</div>
      <div style="padding:4px 0;"><b>T / MTD</b> — Today / Month-to-Date</div>
      <div style="padding:4px 0;"><b>% Month</b> — % of full-month target</div>
      <div style="padding:4px 0;"><b>Subm</b> — Submission</div>
      <div style="padding:4px 0;"><b>Assess</b> — Assessed</div>
    """
    right = """
      <div style="padding:4px 0;"><b>Cmt(T)</b> — Comments (Total)</div>
      <div style="padding:4px 0;"><b>OFU(T)</b> — Offer Follow-ups (Total)</div>
      <div style="padding:4px 0;"><b>PFU(T)</b> — Pending from Partner F/U (Total)</div>
      <div style="padding:4px 0;"><b>Stat(T)</b> — Status Changed (Total)</div>
      <div style="padding:4px 0;"><b>Team</b> — Region — SubRegion group</div>
      <div style="padding:4px 0;"><b>Available</b> — Any KRA measure &gt; 0 yesterday</div>
    """
    two_col = f"""
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
        <tr>
          <td style="width:50%;padding-right:6px;vertical-align:top;">{left}</td>
          <td style="width:50%;padding-left:6px;vertical-align:top;">{right}</td>
        </tr>
      </table>
    """
    return wrap_card(hlabel("ABBREVIATIONS") + two_col)

# ============== PDF helper (optional) ============== #
def html_to_pdf(html: str, pdf_path: Path) -> tuple[bool, str]:
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    global _PDF_WARNED
    try:
        if not html or not str(html).strip():
            print("[WARN] HTML content is empty; skipping PDF generation.")
            return False, "empty-html"
        if PDF_ENGINE.lower() == "wkhtmltopdf":
            if not WKHTMLTOPDF_PATH or not Path(WKHTMLTOPDF_PATH).exists():
                if not _PDF_WARNED:
                    print("[WARN] wkhtmltopdf not found; skipping PDF generation.")
                    _PDF_WARNED = True
                return False, "wkhtmltopdf-missing"
            import pdfkit
            config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)
            options = {"enable-local-file-access": None,"quiet":"","page-size":"A4","margin-top":"8mm","margin-bottom":"8mm","margin-left":"8mm","margin-right":"8mm"}
            pdfkit.from_string(html, str(pdf_path), configuration=config, options=options)
            if pdf_path.exists() and pdf_path.stat().st_size > 0:
                return True, "pdfkit"
            else:
                print(f"[WARN] PDF generation failed - output file empty or missing: {pdf_path}")
                return False, "pdfkit-empty-output"
    except Exception as e:
        print(f"[ERROR] PDF generation error: {e}")
        return False, f"pdfkit-error: {str(e)[:100]}"
    return False, "pdf-disabled"

def send_outlook_html_with_attachments(to_addr: str, subject: str, html: str, attachments: list[Path], DRY_RUN=True, FROM_DISPLAY="",
                                       sig_img_path: Path | None = None, sig_html: str = ""):
    import pythoncom
    import win32com.client as win32
    import win32_utils  # Import the safe dispatch helper
    pythoncom.CoInitialize()
    mail = ns = ol = None
    try:
        # Use safe_dispatch to handle gen_py corruption
        ol = win32_utils.safe_dispatch("Outlook.Application")
        ns = ol.GetNamespace("MAPI")
        mail = ol.CreateItem(0)
        mail.Subject = subject
        # append signature HTML to body
        mail.HTMLBody = html + (sig_html or "")
        mail.To = to_addr
        if FROM_DISPLAY:
            try: mail.SentOnBehalfOfName = FROM_DISPLAY
            except Exception: pass
        # Inline CID image for signature (optional)
        if sig_img_path and sig_img_path.exists():
            try:
                _attach_cid_image_outlook(mail, sig_img_path, "kc_sig")
            except Exception:
                pass
        for p in attachments:
            if p and Path(p).exists(): mail.Attachments.Add(Source=str(p))
        if DRY_RUN: 
            mail.Display(False)
        else:
            for attempt in range(1, RETRY_ON_ERROR + 2):
                try:
                    mail.Send()
                    break
                except Exception:
                    if attempt > RETRY_ON_ERROR:
                        raise
                    time.sleep(3)
    finally:
        try: pythoncom.CoUninitialize()
        except Exception: pass

def send_smtp_html_with_attachments(to_addr: str, subject: str, html: str, attachments: list[Path], DRY_RUN=True,
                                    SMTP_SERVER="", SMTP_PORT=587, SMTP_USERNAME="", SMTP_PASSWORD="", sig_img_path: Path | None = None, sig_html: str = ""):
    import smtplib, mimetypes
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    msg = MIMEMultipart('related')
    sender = SMTP_USERNAME or f"no-reply@{SMTP_SERVER}"
    msg['From'] = sender; msg['To'] = to_addr; msg['Subject'] = subject

    alt = MIMEMultipart('alternative')
    alt.attach(MIMEText(html + (sig_html or ""), 'html', 'utf-8'))
    msg.attach(alt)

    if sig_img_path and sig_img_path.exists():
        with open(sig_img_path, 'rb') as f:
            part = MIMEBase('image', 'png')
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-ID', '<kc_sig>')
        part.add_header('Content-Disposition', 'inline', filename=sig_img_path.name)
        msg.attach(part)

    for p in attachments:
        if not p or not Path(p).exists(): continue
        ctype, _ = mimetypes.guess_type(str(p))
        if not ctype: ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        with open(p, 'rb') as f:
            part = MIMEBase(maintype, subtype); part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=Path(p).name)
        msg.attach(part)

    if DRY_RUN: return
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        if SMTP_USERNAME: server.login(SMTP_USERNAME, SMTP_PASSWORD)
        server.sendmail(sender, [to_addr], msg.as_string())

# ============== PAGE SHELL ============== #
def page_html(title_line: str, body_html: str) -> str:
    return f"""<!doctype html>
<html>
<head><meta charset="utf-8"><meta name="x-apple-disable-message-reformatting">
<title>{title_line}</title></head>
<body style="margin:0;background:{SURFACE_BG};padding:16px 0;font-family:Segoe UI, Arial, sans-serif;">
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="max-width:980px;margin:0 auto;">
    <tr><td style="padding:0 16px 14px 16px;">
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:{BRAND_PRIMARY};border-radius:14px;">
        <tr>
          <td style="padding:16px 18px;">
            <div style="font:800 20px/1.25 Segoe UI,Arial;color:#FFFFFF;">KC Manager Daily — {title_line}</div>
            <div style="margin-top:4px;color:rgba(255,255,255,.9);font:600 13px Segoe UI,Arial;">Daily roll-up for your teams</div>
          </td>
        </tr>
      </table>
    </td></tr>

    <tr><td style="padding:0 16px 10px 16px;">{body_html}</td></tr>
  </table>
</body>
</html>"""

# ============== MAIN DIGEST BUILDER ============== #
def html_digest(manager_email: str,
                work_day: dt.date,
                df_mgr_day: pd.DataFrame,
                team_df: pd.DataFrame,
                role_kra_codes: dict,
                labels_all: dict,
                df_day_all: pd.DataFrame,
                df_all_days: pd.DataFrame,
                month_targets: dict[str, dict[str, int]]) -> str:

    teams_under_mgr = collect_manager_teams(team_df, manager_email)
    label = ", ".join(teams_under_mgr) if teams_under_mgr else manager_email

    sec1 = build_availability_matrix(df_mgr_day, team_df, role_kra_codes)
    sec1b = build_action_summary_team(df_mgr_day, role_kra_codes, labels_all, month_targets)
    sec2 = build_rolewise_work(df_mgr_day, role_kra_codes, labels_all, month_targets)
    sec3 = build_top_bottom(df_mgr_day)
    sec4 = build_team_vs_org_best(df_mgr_day, df_day_all, role_kra_codes, labels_all)

    # Daily ORG (yesterday) side
    org_lb_daily  = build_org_leaderboard_daily(df_day_all)
    org_bad_daily = build_org_team_badges_today(df_day_all, labels_all)

    # MTD ORG side
    org_lb_mtd  = build_org_leaderboard_mtd(df_all_days, work_day)
    org_bad_mtd = build_org_team_badges_mtd(df_day_all, labels_all)

    sec5_daily = f"""
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
        <tr>
          <td style="width:50%;padding-right:6px;vertical-align:top;">{org_lb_daily}</td>
          <td style="width:50%;padding-left:6px;vertical-align:top;">{org_bad_daily}</td>
        </tr>
      </table>
    """

    sec5_mtd = f"""
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
        <tr>
          <td style="width:50%;padding-right:6px;vertical-align:top;">{org_lb_mtd}</td>
          <td style="width:50%;padding-left:6px;vertical-align:top;">{org_bad_mtd}</td>
        </tr>
      </table>
    """

    sec6 = build_abbrev_footer()

    body = "\n".join([sec1, sec1b, sec2, sec3, sec4, sec5_daily, sec5_mtd, sec6])
    title_line = f"{work_day.isoformat()} — {label}"
    return page_html(title_line, body)


# =========================== MAIN ============================ #

def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    team, role_kra_codes, labels_all, holidays, role_meas = load_settings(SETTINGS_PATH)
    df = load_summary(SUMMARY_PATH, SUMMARY_SHEET)
    if df.empty:
        print('No summary rows.'); return

    work_days = resolve_pending_work_days(
        df['ReportDate'].dropna().tolist(), holidays, latest_working_day,
        report_type='manager_digest', max_days=MAX_CATCHUP_DAYS
    )
    if not work_days:
        print('[INFO] No pending manager digests to send.'); return

    roster = team[["EmployeeEmail","EmployeeName","Role","Region","SubRegion","Manager","SendTo"]].copy()

    def send_for_day(work_day: dt.date) -> bool:
        month_targets = build_month_targets(role_meas, holidays, work_day)
        df_day_all = df[df['ReportDate']==work_day].copy()
        sent_recipients = load_sent_recipients('manager_digest', work_day)

        base = roster.merge(df_day_all, on='EmployeeEmail', how='left', suffixes=('_t',''))
        for c in ['EmployeeName','Role','Region','SubRegion','Manager']:
            base[c] = base[c].combine_first(base.get(f"{c}_t")).fillna('')
        for c in [f"{x}_t" for x in ['EmployeeName','Role','Region','SubRegion','Manager']]:
            if c in base.columns:
                base.drop(columns=[c], inplace=True)

        mgrs = sorted(set([m for m in team['Manager'].dropna().astype(str) if '@' in m]))
        if not mgrs:
            print('No manager emails found in Team.Manager.'); return True

        sent = skipped = already_sent = failed = 0
        for mgr in mgrs:
            mgr_norm = str(mgr).strip().lower()
            if mgr_norm in sent_recipients:
                already_sent += 1
                continue
            my_emails = set(team[team['Manager'].str.lower()==mgr.lower()]['EmployeeEmail'].astype(str))
            df_mgr_day = base[base['EmployeeEmail'].isin(my_emails)].copy()
            if df_mgr_day.empty:
                print(f"[SKIP] {mgr}: no rows.")
                skipped += 1
                continue

            html = html_digest(mgr, work_day, df_mgr_day, team, role_kra_codes, labels_all, df_day_all, df, month_targets)

            safe_mgr = re.sub(r'[^a-zA-Z0-9]+', '_', mgr)[:80]
            html_path = OUT_DIR / f"Manager_Digest_{work_day.isoformat()}_{safe_mgr}.html"
            html_path.write_text(html, encoding='utf-8')
            pdf_path = OUT_DIR / f"Manager_Digest_{work_day.isoformat()}_{safe_mgr}.pdf"
            _ok_pdf, pdf_status = html_to_pdf(html, pdf_path)
            
            attachments = [pdf_path] if _ok_pdf else []
            if not _ok_pdf:
                print(f"[WARN] PDF generation skipped ({pdf_status}); email will be sent without PDF attachment")

            teams_under_mgr = collect_manager_teams(team, mgr)
            label = ', '.join(teams_under_mgr) if teams_under_mgr else mgr
            subject = f"{SUBJECT_PREFIX} | {work_day.isoformat()} | {label}"

            try:
                if SEND_VIA.lower()=='outlook':
                    send_outlook_html_with_attachments(
                        mgr, subject, html, attachments,
                        DRY_RUN=DRY_RUN, FROM_DISPLAY=FROM_DISPLAY,
                        sig_img_path=SIG_IMAGE_PATH, sig_html=SIG_HTML
                    )
                else:
                    send_smtp_html_with_attachments(
                        mgr, subject, html, attachments, DRY_RUN=DRY_RUN,
                        SMTP_SERVER=SMTP_SERVER, SMTP_PORT=SMTP_PORT,
                        SMTP_USERNAME=SMTP_USERNAME, SMTP_PASSWORD=SMTP_PASSWORD,
                        sig_img_path=SIG_IMAGE_PATH, sig_html=SIG_HTML
                    )
                sent += 1
                if DRY_RUN:
                    append_sent_detail('manager_digest', work_day, mgr_norm, status='dry_run')
                else:
                    append_sent_detail('manager_digest', work_day, mgr_norm, status='ok')
                print(f"Prepared{' (DRY RUN)' if DRY_RUN else ''} -> {mgr}; files: {[p.name for p in attachments if p]}")
            except Exception as e:
                print(f"[ERROR] Manager <{mgr}> -> {e}")
                traceback.print_exc()
                failed += 1
                append_sent_detail('manager_digest', work_day, mgr_norm, status='fail', meta=str(e)[:200])
        print(f"[INFO] Managers sent: {sent}. Skipped: {skipped}. Already sent: {already_sent}. Failed: {failed}.")
        return failed == 0

    for work_day in work_days:
        try:
            ok = send_for_day(work_day)
            status = 'dry_run' if DRY_RUN else ('ok' if ok else 'fail')
            append_sent_log('manager_digest', work_day, status=status)
        except Exception as exc:
            append_sent_log('manager_digest', work_day, status='fail')
            print(f"[ERROR] Failed manager digest for {work_day}: {exc}")
            traceback.print_exc()


if __name__ == '__main__':
    main()
