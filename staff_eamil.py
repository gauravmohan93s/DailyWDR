# -*- coding: utf-8 -*-
"""
KC Daily — Individual Mailer (v3.9 FINAL — merged with requested tweaks)
- Latest working day (excludes Sundays + Holidays)
- KRA SCORE full width
- Below: two tables side-by-side (50/50) — Outlook-safe table grid
    • LEFT: KRA Components & Today's Actuals
    • RIGHT: Rolling MTD (Month-to-Date, KRA Measures) with MTD, Month Target, %Complete, Gap, Needed/Day
- Role Newsletter: compact Top-3 with component-wise Today (MTD)
- Badges Today & Badges (Month-to-Date) side-by-side; only the approved set
- PDF mirrors the layout
- Certificate: LANDSCAPE Style B (gold frame, navy ribbon, KC logo bottom-left), concise text
- Attachment: Excel (.xlsx) with strict column set; "Action Taken" filled via SubType1 -> ActionTaken -> ActionType -> Action Group
- Skips members whose role has no KRA components OR KRA_Score == 0
- Outlook-only sending; From: gsakhare@kcoverseas.com; inline signature image via CID
"""

from __future__ import annotations
from pathlib import Path
import datetime as dt
import re, hashlib, hmac, base64, urllib.parse as _url
import pandas as pd
import numpy as np
import traceback
import time, gc

from reporting_config import (
    append_sent_log,
    append_sent_detail,
    load_sent_recipients,
    resolve_pending_work_days,
    LOCAL_DB_PATH,
)

# =========================== CONFIG =========================== #
SUMMARY_PATH  = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\raw\summary_all_days.xlsx")
SUMMARY_SHEET = "summary_daily_all"
SETTINGS_PATH = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\setting\wd_settings.xlsx")
OUT_DIR       = SUMMARY_PATH.parent / "daily_email_previews"

# Certificate verification (kept; not used in layout text)
VERIFY_BASE_URL = ""  # leave empty if unused
VERIFY_SECRET   = "replace-with-a-strong-random-secret"

# KC logo (for PDF certificate)
LOGO_PATH     = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\Conversion Report\KC-LOGO-1-1024x225.png")

# --- Email transport ---
SEND_VIA         = "outlook"   # 'outlook' only
DRY_RUN          = False       # send for real
ATTACH_PDF       = True
ATTACH_CERT_PDF  = True

PREVIEW_LIMIT                 = 0
SAVE_DRAFTS_WHEN_DRY_RUN      = True
BATCH_SIZE                    = 100
SLEEP_BETWEEN_BATCHES_SEC     = 1
RETRY_ON_ERROR                = 2
MAX_CATCHUP_DAYS              = 7  # safety cap when catching up on missed days

# SMTP kept for completeness; unused when SEND_VIA='outlook'
SMTP_SERVER      = "smtp.office365.com"
SMTP_PORT        = 587
SMTP_USERNAME    = ""
SMTP_PASSWORD    = ""

FROM_DISPLAY     = "gsakhare@kcoverseas.com"  # Send-as (must have permission in Outlook profile)
CC_MANAGER       = False                      # per request: NO manager CC
SUBJECT_PREFIX   = "KC Daily"
DEFAULT_INCLUDE_IF_IN_SUMMARY = True

# Signature (HTML + image path; inline via CID=kc_sig)
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
SIG_IMG_PATH = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\Conversion Report\signature_banner.jpg")

WKHTMLTOPDF_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
PDF_ENGINE = "wkhtmltopdf"
_PDF_WARNED = False
IST = dt.timezone(dt.timedelta(hours=5, minutes=30))

# =================== BRAND / UI =================== #
BRAND_PRIMARY    = "#0B1220"
BRAND_ACCENT     = "#2563EB"
SURFACE_BG       = "#F5F7FB"
CARD_BG          = "#FFFFFF"
TEXT_DARK        = "#111827"
TEXT_MUTED       = "#6B7280"
ACCENT_SUCCESS   = "#16A34A"
ACCENT_WARN      = "#B45309"
ACCENT_DANGER    = "#DC2626"
DIVIDER          = "#E5E7EB"
TABLE_HEAD_BG    = "#F3F4F6"
PILL_BG          = "#EEF2FF"
PILL_TEXT        = "#3730A3"

# Allowed badges (both Daily & MTD)
ALLOWED_BADGE_TITLES = [
    "Top KRA",
    "Top Total Comments",
    "Top Total Offer Follow Ups",
    "Top Total Pending from Partner F/U",
    "Top Total Status Changed",
    "Top Submission",
]

# Optional mapping for MTD badges → preferred measure code (the part AFTER 'MTD_')
BADGE_MTD_CODE_HINTS = {
    # Fill to force mapping if needed, e.g.:
    # "Top Total Comments": "Total_Comments",
}

# ---------------- Remaining Workdays Helper ---------------- #
HOLIDAYS: set[dt.date] = set()

def _remaining_workdays_in_month(d: dt.date, holidays: set[dt.date]) -> int:
    """Remaining working days AFTER 'd' in the same month (excludes Sundays + holidays)."""
    import calendar
    last_day = dt.date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])
    cur = d + dt.timedelta(days=1)
    cnt = 0
    while cur <= last_day:
        if cur.weekday() != 6 and cur not in holidays:  # 6=Sunday
            cnt += 1
        cur += dt.timedelta(days=1)
    return cnt

# --- Full-month workday helpers (exclude Sundays + Holidays) ---
import calendar

def _total_workdays_in_month(d: dt.date, holidays: set[dt.date]) -> int:
    """Working days in the whole month of d (excludes Sundays + holidays)."""
    start = dt.date(d.year, d.month, 1)
    end   = dt.date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])
    cur = start
    cnt = 0
    while cur <= end:
        if cur.weekday() != 6 and cur not in holidays:  # 6 = Sunday
            cnt += 1
        cur += dt.timedelta(days=1)
    return cnt

def _daily_targets_for_role(role: str, role_meas: pd.DataFrame) -> dict[str, float]:
    """
    Returns {MeasureCode: DailyTarget} for measures that IncludeInKRA for the role.
    Missing/NaN DailyTarget → treated as 0.
    """
    rm = role_meas[(role_meas["Role"] == role) & (role_meas["IncludeInKRA"] == True)].copy()
    rm["DailyTarget"] = pd.to_numeric(rm.get("DailyTarget", 0), errors="coerce").fillna(0.0)
    out = {}
    for _, r in rm.iterrows():
        code = str(r["MeasureCode"])
        out[code] = float(r["DailyTarget"])
    return out

# =========================== LOADERS =========================== #
def load_settings(settings_path: Path):
    xls = pd.read_excel(settings_path, sheet_name=None)

    team = xls.get("Team")
    if team is None:
        raise ValueError("wd_settings.xlsx missing 'Team' sheet.")
    for c in ["EmployeeEmail","EmployeeName","Role","Region","SubRegion","Manager","Include","SendTo"]:
        if c not in team.columns: team[c] = ""
    team["EmployeeEmail"] = team["EmployeeEmail"].astype(str).str.strip().str.lower()
    team["Manager"] = team["Manager"].astype(str).str.strip()
    team["SendTo"]  = team["SendTo"].astype(str).str.strip()
    team["Include"] = (
        team["Include"].astype(str).str.strip().str.lower()
        .map({"yes":True,"y":True,"1":True,"true":True,"t":True})
        .fillna(False)
    )

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
    role_meas["IncludeInKRA"] = role_meas["IncludeInKRA"].astype(str).str.strip().str.upper().map({"Y":True,"YES":True}).fillna(False)
    for c in ["Weight","DailyTarget","MinFloor","Cap","DisplayOrder"]:
        role_meas[c] = pd.to_numeric(role_meas[c], errors="coerce")

    kra_meas_union = sorted(set(role_meas.loc[role_meas["IncludeInKRA"], "MeasureCode"].dropna().astype(str)))
    labels_kra_only_union = {c: labels_all.get(c, c) for c in kra_meas_union}

    holidays = set()
    hol = xls.get("Holidays")
    if hol is not None:
        for col in ["Date","HolidayDate","Holiday","Dt"]:
            if col in hol.columns:
                h = pd.to_datetime(hol[col], errors="coerce").dt.date.dropna()
                holidays = set(h.tolist())
                break

    return team, measures, labels_all, labels_kra_only_union, role_meas, holidays, kra_meas_union

def load_summary(path: Path, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet)
    df["ReportDate"] = pd.to_datetime(df["ReportDate"]).dt.date
    df["EmployeeEmail"] = df["EmployeeEmail"].astype(str).str.strip().str.lower()
    for c in ["EmployeeName","Role","Region","SubRegion","Manager","Badges_Today"]:
        if c in df.columns: df[c] = df[c].fillna("")
    return df

def get_local_db_engine():
    """Returns the engine for the local SQLite database."""
    from sqlalchemy import create_engine
    db_path = Path(LOCAL_DB_PATH)
    if not db_path.exists():
        raise FileNotFoundError(f"Database not found at {db_path}. Please run the extractor script to create it.")
    return create_engine(f"sqlite:///{db_path}")

def load_merged_actions() -> pd.DataFrame:
    """Loads all actions from the local SQLite database."""
    print("[DB] Loading data from kc_reports.db for staff mailer...")
    engine = get_local_db_engine()
    
    # Load data and ensure datetime conversion
    df = pd.read_sql("SELECT * FROM actions", engine, parse_dates=["ActionDate"])
    print(f"[DB] Loaded {len(df)} records for staff mailer.")

    # Re-establish timezone awareness (data is stored as UTC text)
    df["ActionDate"] = df["ActionDate"].dt.tz_localize('UTC')

    # Create IST date columns
    df["ActionDateIST"] = df["ActionDate"].dt.tz_convert("Asia/Kolkata")
    df["ActionDateIST_Date"] = df["ActionDateIST"].dt.date
    
    # Format ActionDateIST for display
    df["ActionDateIST"] = df["ActionDateIST"].dt.strftime("%Y-%m-%d %H:%M:%S")

    # Populate ActionTaken with priority: SubType1 -> ActionType -> Action Group
    def _fill_action_taken(row):
        s1 = str(row.get("SubType1","") or "").strip()
        if s1: return s1
        v2 = str(row.get("ActionType","") or "").strip()
        if v2: return v2
        return str(row.get("Action Group","") or "").strip()
    df["ActionTaken"] = df.apply(_fill_action_taken, axis=1)

    return df

# ==================== WORKING DAY HELPERS ==================== #
def is_sunday(d: dt.date) -> bool: return d.weekday() == 6
def latest_working_day(dates: list[dt.date], holidays: set[dt.date]) -> dt.date | None:
    for d in sorted(set(dates), reverse=True):
        if (not is_sunday(d)) and (d not in holidays): return d
    return None

# ========================= HTML HELPERS ======================== #
def _safe_int(x):
    try: return int(round(float(x)))
    except Exception: return 0

def _pill(text: str, bg: str=PILL_BG, fg: str=PILL_TEXT):
    return f"""<span style="display:inline-block;padding:5px 12px;border-radius:999px;background:{bg};color:{fg};font:600 12px/1 Inter,Segoe UI,Arial;margin-right:8px;">{text}</span>"""

def signature_block():
    # We append SIG_HTML at send-time; keep this simple divider only.
    return f"""
      <div style="margin-top:6px;color:{TEXT_MUTED};font:12px Inter,Segoe UI,Arial;">
        <div style="font-weight:700;color:{TEXT_DARK}">—</div>
      </div>
    """

# ======== Skip Logic ======== #
def role_has_kra(role: str, role_meas: pd.DataFrame) -> bool:
    if not isinstance(role, str): 
        return False
    sub = role_meas[(role_meas["Role"] == role) & (role_meas["IncludeInKRA"] == True)]
    return not sub.empty

# ======== Combined KRA Components + Today (table) ======== #
def combined_kra_section(role: str, row: pd.Series, role_meas: pd.DataFrame, labels_all: dict) -> str:
    rm = role_meas[(role_meas["Role"] == role) & (role_meas["IncludeInKRA"])].copy()
    if rm.empty:
        return f"""
          <div style="font:800 12px/1 Inter;color:{TEXT_MUTED};text-transform:uppercase;letter-spacing:.04em;">KRA Components & Today's Actuals</div>
          <div style="margin-top:8px;padding:12px;border:1px dashed {DIVIDER};border-radius:10px;background:#FAFAFB;color:{TEXT_MUTED}">No KRA components configured for your role yet.</div>
        """
    rm["_row_order"] = np.arange(len(rm))
    rm["__has"] = rm["DisplayOrder"].notna().astype(int)
    rm = rm.sort_values(by=["__has","DisplayOrder","_row_order"], ascending=[False, True, True])

    total_w = float(rm["Weight"].fillna(0).sum())
    rm["WeightNorm"] = (rm["Weight"].fillna(0) / total_w) if total_w > 0 else 0.0

    rows = []
    for _, r in rm.iterrows():
        code = str(r["MeasureCode"])
        lbl  = labels_all.get(code, code)
        wt_pct = int(round(float(r["WeightNorm"])*100.0))
        tgt = r.get("DailyTarget")
        tgt_txt = f"{_safe_int(tgt)}" if pd.notna(tgt) and float(tgt)>0 else "-"
        val_today = int(row.get(code, 0))
        pct_day   = row.get(f"{code}_PctOfDay", 0.0) or 0
        rank      = row.get(f"Rank_{code}", np.nan)
        rank_txt  = f"#{int(rank)}" if not pd.isna(rank) else "-"
        rows.append(f"""
          <tr>
            <td style="padding:10px 12px;border-bottom:1px solid {DIVIDER};">{lbl}</td>
            <td style="padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};">{wt_pct}%</td>
            <td style="padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};">{tgt_txt}</td>
            <td style="padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;font-weight:700;color:{TEXT_DARK};">{val_today}</td>
            <td style="padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};">{pct_day}%</td>
            <td style="padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};">{rank_txt}</td>
          </tr>""")

    return f"""
      <div style="font:800 12px/1 Inter;color:{TEXT_MUTED};text-transform:uppercase;letter-spacing:.04em;">KRA Components & Today's Actuals</div>
      <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:10px;border-collapse:collapse;border-radius:10px;overflow:hidden;background:{CARD_BG};border:1px solid {DIVIDER};">
        <thead>
          <tr style="background:{TABLE_HEAD_BG};color:{TEXT_MUTED};text-align:left;">
            <th style="padding:10px 12px;font:800 12px/1 Inter;">Measure</th>
            <th style="padding:10px 12px;font:800 12px/1 Inter;text-align:right;">Weight</th>
            <th style="padding:10px 12px;font:800 12px/1 Inter;text-align:right;">Target</th>
            <th style="padding:10px 12px;font:800 12px/1 Inter;text-align:right;">Today</th>
            <th style="padding:10px 12px;font:800 12px/1 Inter;text-align:right;">% of Day</th>
            <th style="padding:10px 12px;font:800 12px/1 Inter;text-align:right;">Rank</th>
          </tr>
        </thead>
        <tbody>{''.join(rows)}</tbody>
      </table>
    """

# ======== Rolling MTD (KRA Measures) as table with Gap/Needed ======== #
def rolling_mtd_table(row: pd.Series,
                      labels_kra_for_member: dict,
                      role: str,
                      role_meas: pd.DataFrame,
                      holidays: set[dt.date]) -> str:
    """
    Uses FULL-MONTH target:
      Month Target = DailyTarget (from RoleMeasures) * total working days in the month
      %Complete    = MTD / Month Target
      Gap          = max(0, Month Target - MTD)  (rest of month)
      Needed/Day   = ceil(Gap / remaining working days in month)
    """
    # Date + workday counts
    d = row["ReportDate"]
    total_workdays = _total_workdays_in_month(d, holidays)
    remaining_days = _remaining_workdays_in_month(d, holidays)

    # Daily targets for this role (per measure)
    daily_targets = _daily_targets_for_role(role, role_meas)

    mrows = []
    for code, label in labels_kra_for_member.items():
        cur = float(row.get(f"MTD_{code}", 0.0))  # actual MTD (sum of days)
        daily_tgt = float(daily_targets.get(code, 0.0))
        month_tgt = daily_tgt * total_workdays

        # Build metrics
        pct  = 0 if month_tgt <= 0 else min(100, max(0, int(round((cur / month_tgt) * 100))))
        gap  = max(0, int(round(month_tgt - cur)))
        need = int(np.ceil(gap / remaining_days)) if remaining_days > 0 else 0

        # Row (always show numeric; if month_tgt==0, they’ll read as 0/0/0)
        mrows.append(f"""
          <tr>
            <td style="padding:10px 12px;border-bottom:1px solid {DIVIDER};">{label}</td>
            <td style="padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;font-weight:700;color:{TEXT_DARK};">{_safe_int(cur)}</td>
            <td style="padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};">{_safe_int(month_tgt)}</td>
            <td style="padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};">{pct}%</td>
            <td style="padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};">{gap}</td>
            <td style="padding:10px 12px;border-bottom:1px solid {DIVIDER};text-align:right;color:{TEXT_MUTED};">{need}</td>
          </tr>
        """)

    body = ''.join(mrows) or f"<tr><td colspan='6' style='padding:10px 12px;color:{TEXT_MUTED};'>No KRA targets configured.</td></tr>"
    return f"""
      <div style="font:800 12px/1 Inter;color:{TEXT_MUTED};text-transform:uppercase;letter-spacing:.04em;">Rolling MTD (Month-to-Date, KRA Measures)</div>
      <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:10px;border-collapse:collapse;border-radius:10px;overflow:hidden;background:{CARD_BG};border:1px solid {DIVIDER};">
        <thead>
          <tr style="background:{TABLE_HEAD_BG};color:{TEXT_MUTED};text-align:left;">
            <th style="padding:10px 12px;font:800 12px/1 Inter;">Measure</th>
            <th style="padding:10px 12px;font:800 12px/1 Inter;text-align:right;">MTD</th>
            <th style="padding:10px 12px;font:800 12px/1 Inter;text-align:right;">Month Target</th>
            <th style="padding:10px 12px;font:800 12px/1 Inter;text-align:right;">%Complete</th>
            <th style="padding:10px 12px;font:800 12px/1 Inter;text-align:right;">Gap</th>
            <th style="padding:10px 12px;font:800 12px/1 Inter;text-align:right;">Needed / Day</th>
          </tr>
        </thead>
        <tbody>{body}</tbody>
      </table>
    """

# ===================== Role Newsletter (Top-3) ===================== #
def build_role_leaderboard(df_day: pd.DataFrame, role: str, labels_kra_for_role: dict) -> str:
    rdf = df_day[df_day["Role"] == role].copy()
    if rdf.empty: return ""
    codes = list(labels_kra_for_role.keys())
    top = rdf.copy().sort_values("KRA_Score", ascending=False).head(3)
    rows = []
    for _, r in top.iterrows():
        nm = r["EmployeeName"] or r["EmployeeEmail"]
        region = r.get("Region",""); subreg = r.get("SubRegion","")
        kra = round(float(r.get("KRA_Score", 0.0)), 1)
        parts = []
        for code in codes:
            lbl = labels_kra_for_role.get(code, code)
            t = int(r.get(code, 0)); m = int(r.get(f"MTD_{code}", 0))
            parts.append(f"{lbl}: <b>{t}</b> (<span style='color:{TEXT_MUTED}'>{m}</span>)")
        comp_inline = " · ".join(parts) if parts else "<span style='color:{TEXT_MUTED}'>No KRA components</span>"
        rows.append(f"""
          <tr>
            <td style="padding:8px 10px;border-bottom:1px solid {DIVIDER};"><b>{nm}</b><div style="color:{TEXT_MUTED};font:600 12px/1.2 Inter;">{region}{(' / '+subreg) if subreg else ''}</div></td>
            <td style="padding:8px 10px;border-bottom:1px solid {DIVIDER};text-align:right;color:{BRAND_ACCENT};font-weight:800;">{kra}</td>
            <td style="padding:8px 10px;border-bottom:1px solid {DIVIDER};">{comp_inline}</td>
          </tr>""")
    return f"""
    <div style="background:{CARD_BG};border:1px solid {DIVIDER};border-radius:12px;padding:16px;">
      <div style="font:800 15px/1 Inter;color:{TEXT_DARK};margin-bottom:6px;">Role Newsletter — {role}</div>
      <div style="font:600 12px/1 Inter;color:{TEXT_MUTED};margin-bottom:8px;">Top KRA performers with component-wise Today (MTD)</div>
      <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border-radius:8px;overflow:hidden;">
        <thead>
          <tr style="background:{TABLE_HEAD_BG};color:{TEXT_MUTED};text-align:left;">
            <th style="padding:8px 10px;font:800 12px/1 Inter;">Member</th>
            <th style="padding:8px 10px;font:800 12px/1 Inter;text-align:right;">KRA</th>
            <th style="padding:8px 10px;font:800 12px/1 Inter;">KRA Components — Today (MTD)</th>
          </tr>
        </thead>
        <tbody>{''.join(rows)}</tbody>
      </table>
    </div>
    """

# ================== Badges (Daily & MTD) ================== #
def _norm(s: str) -> str:
    import re
    return re.sub(r"[^a-z0-9]+", " ", str(s or "").lower()).strip()

def _token_set(s: str) -> set[str]:
    return set(_norm(s).split())

_TITLE_KEYWORDS = {
    "Top Total Comments": {"comment", "comments"},
    "Top Total Offer Follow Ups": {"offer", "follow", "ups", "followups", "follow up"},
    "Top Total Pending from Partner F/U": {"pending", "partner", "follow", "fu", "f u"},
    "Top Total Status Changed": {"status", "changed", "change"},
    "Top Submission": {"submission", "submit", "submitted"},
}

def _resolve_mtd_column_for_title(summary_on_day: pd.DataFrame,
                                  labels_all: dict[str, str],
                                  title: str,
                                  code_hints: dict[str, str]) -> str | None:
    mtd_cols = [c for c in summary_on_day.columns if c.startswith("MTD_")]
    if not mtd_cols: return None

    hint = (code_hints or {}).get(title, "").strip().lower()
    if hint:
        for col in mtd_cols:
            base = col[4:]
            if _norm(base) == _norm(hint):
                return col
        for col in mtd_cols:
            base = col[4:]
            if hint in _norm(base) or hint in _norm(labels_all.get(base, base)):
                return col

    keys = _TITLE_KEYWORDS.get(title, set())
    if not keys:
        keys = _token_set(title) - {"top","total","monthly","month","leaders","leader"}

    def score_col(col: str) -> tuple[int, int, str]:
        base = col[4:]
        base_norm = _norm(base)
        lab_norm = _norm(labels_all.get(base, base))
        score = sum(1 for k in keys if k in base_norm or k in lab_norm)
        total_bias = 0 if ("total" in base_norm or "total" in lab_norm) else 1
        return (score, -total_bias, base)

    ranked = sorted(mtd_cols, key=score_col, reverse=True)
    best = ranked[0]
    if score_col(best)[0] <= 0:
        return None
    return best

def _match_best_mtd(summary_on_day: pd.DataFrame, labels_all: dict, title: str, code_hints: dict[str, str]):
    col = _resolve_mtd_column_for_title(summary_on_day, labels_all, title, code_hints)
    if not col: return None
    sub = summary_on_day[[col, "EmployeeName", "EmployeeEmail"]].copy()
    sub[col] = pd.to_numeric(sub[col], errors="coerce").fillna(0)
    sub = sub.sort_values(col, ascending=False)
    if sub.empty or sub.iloc[0][col] <= 0:
        return None
    row = sub.iloc[0]
    who = row["EmployeeName"] or row["EmployeeEmail"]
    base = col[4:]
    label = labels_all.get(base, base)
    return title, who, label, int(row[col])

def winners_daily_badges(df_day: pd.DataFrame, labels_all: dict):
    out = []
    top_kra = df_day[df_day["KRA_Score"]>0].sort_values("KRA_Score", ascending=False)
    if not top_kra.empty:
        best = top_kra.iloc[0]
        out.append(("Top KRA", best["EmployeeName"] or best["EmployeeEmail"], round(float(best["KRA_Score"]),1)))

    allowed = set(t.lower() for t in ALLOWED_BADGE_TITLES if t!="Top KRA")
    for col in [c for c in df_day.columns if c.startswith("Rank_")]:
        code = col.replace("Rank_","")
        if code not in df_day.columns: continue
        label = labels_all.get(code, code)
        title = f"Top {label}"
        if title.lower() not in allowed: 
            continue
        sub = df_day[(df_day[col]==1) & (df_day[code]>0)]
        if sub.empty: continue
        who = sub.iloc[0]["EmployeeName"] or sub.iloc[0]["EmployeeEmail"]
        val = int(sub.iloc[0][code])
        out.append((title, who, val))

    seen, dedup = set(), []
    for t, w, v in out:
        if t.lower() in seen: continue
        seen.add(t.lower()); dedup.append((t,w,v))
    final = []
    for t in ALLOWED_BADGE_TITLES:
        for (tt, w, v) in dedup:
            if tt.lower() == t.lower():
                final.append((tt,w,v)); break
    return final

def build_badges_today_block(df_day: pd.DataFrame, labels_all: dict) -> str:
    winners = winners_daily_badges(df_day, labels_all)
    items = []
    for title, who, val in winners:
        pill = _pill(f"{title} — {val}", bg="#DCFCE7", fg="#065F46")
        items.append(f"<div style='padding:4px 0;'>{pill}<div style='color:{TEXT_MUTED};font:600 12px/1.2 Inter;'>Winner: <b style='color:{TEXT_DARK}'>{who}</b></div></div>")
    body = ''.join(items) or "<div style='color:#6B7280'>No badges were earned today.</div>"
    return f"""
      <div style="background:{CARD_BG};border:1px solid {DIVIDER};border-radius:12px;padding:16px;">
        <div style="font:800 15px/1 Inter;color:{TEXT_DARK};margin-bottom:6px;">Badges — Today</div>
        {body}
      </div>
    """

def build_monthly_badges_block(summary: pd.DataFrame, work_day: dt.date, labels_all: dict) -> str:
    on_day = summary[summary["ReportDate"] == work_day].copy()
    if on_day.empty:
        return f"""
          <div style="background:{CARD_BG};border:1px solid {DIVIDER};border-radius:12px;padding:16px;">
            <div style="font:800 15px/1 Inter;color:{TEXT_DARK};margin-bottom:6px;">Badges — Monthly Leaders (Month-to-Date)</div>
            <div style="color:{TEXT_MUTED}">No data for this date.</div>
          </div>
        """
    items_html, seen_titles = [], set()
    titles = [t for t in ALLOWED_BADGE_TITLES if t != "Top KRA"]
    for t in titles:
        res = _match_best_mtd(on_day, labels_all, t, BADGE_MTD_CODE_HINTS)
        if not res: continue
        title, who, label, val = res
        if title.lower() in seen_titles: continue
        seen_titles.add(title.lower())
        pill = _pill(f"{title} — {label}: {val}", bg="#DBEAFE", fg="#1E40AF")
        items_html.append(f"<div style='padding:4px 0;'>{pill}<div style='color:{TEXT_MUTED};font:600 12px/1.2 Inter;'>Leader: <b style='color:{TEXT_DARK}'>{who}</b></div></div>")
    body = ''.join(items_html) or f"<div style='color:{TEXT_MUTED}'>No qualifying MTD leaders found.</div>"
    return f"""
      <div style="background:{CARD_BG};border:1px solid {DIVIDER};border-radius:12px;padding:16px;">
        <div style="font:800 15px/1 Inter;color:{TEXT_DARK};margin-bottom:6px;">Badges — Monthly Leaders (Month-to-Date)</div>
        {body}
      </div>
    """

# ======================= PDF/EXCEL ====================== #
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

def _build_verification_payload(name: str, role: str, region: str, subregion: str,
                                date: dt.date, kra: float, rank_role: int) -> tuple[str, str]:
    payload = {"n":name,"r":role,"rg":region,"sr":subregion,"d":date.isoformat(),"k":round(float(kra),1),"rr":int(rank_role),"v":1}
    s = f"{payload['n']}|{payload['r']}|{payload['rg']}|{payload['sr']}|{payload['d']}|{payload['k']:.1f}|{payload['rr']}"
    vid = hashlib.sha256(s.encode("utf-8")).hexdigest()[:16].upper()
    import json
    raw = json.dumps(payload, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    b64 = base64.urlsafe_b64encode(raw).rstrip(b"=").decode("ascii")
    sig = hmac.new(VERIFY_SECRET.encode("utf-8"), b64.encode("ascii"), hashlib.sha256).digest()
    sig_b64 = base64.urlsafe_b64encode(sig).rstrip(b"=").decode("ascii")
    token = f"{b64}.{sig_b64}"
    return vid, token

def build_verification_url(name: str, role: str, region: str, subregion: str,
                           date: dt.date, kra: float, rank_role: int) -> tuple[str, str]:
    vid, token = _build_verification_payload(name, role, region, subregion, date, kra, rank_role)
    if VERIFY_BASE_URL:
        q = _url.urlencode({"t": token})
        return vid, f"{VERIFY_BASE_URL}?{q}"
    else:
        return vid, f"KC-CERT:{vid}"

from reportlab.pdfbase.pdfmetrics import stringWidth

def _wrap_lines(text, font_name, font_size, max_width):
    text = str(text or "").strip()
    if not text:
        return []
    words, lines, cur = text.split(), [], ""
    for w in words:
        test = (cur + " " + w).strip()
        if stringWidth(test, font_name, font_size) <= max_width:
            cur = test
        else:
            if cur: lines.append(cur)
            cur = w
    if cur: lines.append(cur)
    return lines

def _draw_paragraph(c, x, y_top, width, line_height, lines, max_lines, font_name, font_size, ellipsis=True):
    c.setFont(font_name, font_size)
    used = 0
    for i, l in enumerate(lines):
        if used == max_lines:
            break
        c.drawString(x, y_top - used*line_height, l)
        used += 1
    if ellipsis and len(lines) > max_lines and used >= 1:
        last = lines[max_lines-1]
        dots = " …"
        while stringWidth(last + dots, font_name, font_size) > width and last:
            last = last[:-1]
        c.drawString(x, y_top - (max_lines-1)*line_height, (last + dots).rstrip())
    return y_top - (max(used, 1) - 1)*line_height

def build_peer_stats_for_role(df_day: pd.DataFrame, role: str, labels_for_member: dict) -> dict:
    sub = df_day[df_day["Role"] == role].copy()
    ps = {"kra_avg": 0.0, "meas_avg": {}}
    if sub.empty:
        return ps
    ps["kra_avg"] = float(pd.to_numeric(sub["KRA_Score"], errors="coerce").mean())
    for code in labels_for_member.keys():
        if code in sub.columns:
            ps["meas_avg"][code] = float(pd.to_numeric(sub[code], errors="coerce").mean())
    return ps

def make_badge_certificate(
    name: str,
    role: str,
    region: str,
    subregion: str,
    date: dt.date,
    kra: float,
    rank_role: int,
    top_breakdown: list[tuple[str, int]],
    badge_title: str,
    pdf_path: Path,
    logo_path: Path = LOGO_PATH,
    labels_for_member: dict | None = None,
    peer_stats: dict | None = None,
) -> tuple[bool, str]:
    """
    KC Daily — Certificate (Portrait, refined navy header, gold accent, no top logo)
    Sections:
      1) Navy header ribbon with title (lighter navy, gold underline)
      2) Badge title (large)
      3) Awardee info
      4) KRA strip (KRA Score + Role Rank)
      5) KRA Components — You vs Team table (single merged, clean lines)
      6) White footer strip with KC logo at bottom-left
    All values are rounded to whole numbers for readability.
    """
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.lib.colors import HexColor, white
        from reportlab.pdfgen import canvas
        from reportlab.pdfbase.pdfmetrics import stringWidth

        # ---------- Palette (lighter navy + bold gold) ----------
        NAVY_HDR   = HexColor("#1C2E59")   # lighter navy for header
        NAVY_TEXT  = HexColor("#0F1B3D")   # deep navy for headings
        GOLD       = HexColor("#D4B245")   # warm gold accent (slightly yellowish)
        PAPER      = HexColor("#FFFFFF")
        TXT_DARK   = HexColor("#111827")
        TXT_MUTED  = HexColor("#6B7280")
        RULE_LIGHT = HexColor("#EBEDF0")   # subtle table rules
        POS        = HexColor("#15803D")
        NEG        = HexColor("#B91C1C")

        labels_for_member = labels_for_member or {}
        peer_stats = peer_stats or {"kra_avg": 0.0, "meas_avg": {}}

        kra_i     = int(round(float(kra)))
        rank_i    = int(rank_role)
        kra_avg_i = int(round(float(peer_stats.get("kra_avg", 0.0))))
        delta_i   = kra_i - kra_avg_i

        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        c = canvas.Canvas(str(pdf_path), pagesize=A4)
        W, H = A4  # portrait

        # ---------- Helpers ----------
        def wrap_lines(text, font, size, maxw):
            text = str(text or "").strip()
            if not text:
                return []
            words, out, cur = text.split(), [], ""
            for w in words:
                t = (cur + " " + w).strip()
                if stringWidth(t, font, size) <= maxw:
                    cur = t
                else:
                    if cur:
                        out.append(cur)
                    cur = w
            if cur:
                out.append(cur)
            return out

        def draw_paragraph(x, y_top, w, font, size, line_h, text, max_lines=2, color=TXT_DARK):
            c.setFillColor(color); c.setFont(font, size)
            lines = wrap_lines(text, font, size, w)
            used = 0
            for ln in lines[:max_lines]:
                c.drawString(x, y_top - used * line_h, ln)
                used += 1
            return y_top - max(used, 1) * line_h

        # ---------- Background ----------
        c.setFillColor(PAPER); c.rect(0, 0, W, H, stroke=0, fill=1)

        # ---------- Header ribbon (lighter navy) ----------
        header_h = 2.1 * cm
        c.setFillColor(NAVY_HDR)
        c.rect(0, H - header_h, W, header_h, stroke=0, fill=1)
        # bold gold underline
        c.setStrokeColor(GOLD); c.setLineWidth(2.2)
        c.line(0.9 * cm, H - header_h, W - 0.9 * cm, H - header_h)

        # Title (no top logo)
        c.setFillColor(white); c.setFont("Helvetica-Bold", 19)
        c.drawString(1.6 * cm, H - header_h / 2 + 0.05 * cm, "KC Daily — Certificate of Achievement")

        # ---------- Badge Title ----------
        c.setFillColor(NAVY_TEXT); c.setFont("Helvetica-Bold", 26)
        c.drawCentredString(W / 2, H - header_h - 1.1 * cm, str(badge_title or "Achievement"))

        # ---------- Layout constants ----------
        margin_x = 1.8 * cm
        col_gap  = 1.0 * cm
        left_w   = (W - 2 * margin_x - col_gap) * 0.56
        right_w  = (W - 2 * margin_x - col_gap) - left_w
        content_top = H - header_h - 2.7 * cm

        # ---------- Awardee info ----------
        xL, y = margin_x, content_top
        y = draw_paragraph(xL, y, left_w, "Helvetica-Bold", 14, 16, f"Awarded to: {name}", max_lines=1)
        y = draw_paragraph(xL, y, left_w, "Helvetica", 11, 13, f"Role: {role}", max_lines=1, color=TXT_MUTED) + 2
        reg_line = f"Region: {region}" + (f"  |  Sub-Region: {subregion}" if subregion else "")
        y = draw_paragraph(xL, y, left_w, "Helvetica", 11, 13, reg_line, max_lines=1, color=TXT_MUTED) + 2
        c.setFillColor(TXT_MUTED); c.setFont("Helvetica", 11)
        c.drawString(xL, y, f"Date: {date.isoformat()}")
        y -= 16

        # space before KRA strip
        y -= 6

        # ---------- KRA strip ----------
        c.setFillColor(TXT_MUTED); c.setFont("Helvetica-Bold", 11)
        c.drawString(xL, y, "KRA Score")
        c.setFillColor(NAVY_TEXT); c.setFont("Helvetica-Bold", 22)
        c.drawString(xL + 2.7 * cm, y, str(kra_i))

        c.setFillColor(TXT_MUTED); c.setFont("Helvetica-Bold", 11)
        c.drawString(xL + 7.0 * cm, y, "Role Rank")
        c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 22)
        c.drawString(xL + 10.0 * cm, y, f"#{rank_i}")

        # gold rule below KRA strip
        y -= 18
        c.setStrokeColor(GOLD); c.setLineWidth(1.4)
        c.line(margin_x, y, W - margin_x, y)

        # add breathing room before table
        y -= 12

        # ---------- KRA Components — You vs Team (single merged table) ----------
        c.setFillColor(TXT_DARK); c.setFont("Helvetica-Bold", 13)
        c.drawString(margin_x, y, "KRA Components — You vs Team")
        y -= 16

        # Team Avg line (concise)
        c.setFillColor(TXT_MUTED); c.setFont("Helvetica", 11)
        if delta_i == 0:
            c.drawString(margin_x, y, f"Team Avg (KRA): {kra_avg_i} — You are the same as the team average")
        else:
            c.drawString(margin_x, y, f"Team Avg (KRA): {kra_avg_i} — You are {abs(delta_i)} {'above' if delta_i > 0 else 'below'} average")
        y -= 14

        # Table geometry
        table_x = margin_x
        table_w = W - 2 * margin_x
        col_component = int(table_w * 0.50)
        col_you = int(table_w * 0.16)
        col_team = int(table_w * 0.17)
        col_delta = int(table_w * 0.17)

        # Header rule (thin & subtle)
        c.setStrokeColor(RULE_LIGHT); c.setLineWidth(1)
        c.line(table_x, y, table_x + table_w, y)
        y -= 4

        # Header labels
        c.setFillColor(TXT_MUTED); c.setFont("Helvetica-Bold", 11)
        c.drawString(table_x + 6, y, "Component")
        c.drawRightString(table_x + col_component + col_you - 6, y, "You Today")
        c.drawRightString(table_x + col_component + col_you + col_team - 6, y, "Team Avg")
        c.drawRightString(table_x + table_w - 6, y, "Delta")
        y -= 8

        # Second header rule
        c.setStrokeColor(RULE_LIGHT); c.line(table_x, y, table_x + table_w, y)
        y -= 6

        # Rows (capped to fit page)
        row_h = 14
        max_rows = int((y - 2.2 * cm) // row_h)
        rows = 0
        rows_data = top_breakdown[:]  # [(label, you_today), ...]
        if not rows_data and labels_for_member:
            for code, lbl in labels_for_member.items():
                rows_data.append((lbl, 0))

        for lbl, val in rows_data:
            if rows >= max_rows:
                break
            # find code to read team avg
            code = None
            for k, v in (labels_for_member or {}).items():
                if v == lbl:
                    code = k
                    break
            team_avg = int(round(float(peer_stats.get("meas_avg", {}).get(code, 0.0)))) if code else 0
            you_i = int(val)
            delta_m = you_i - team_avg

            # component label
            c.setFillColor(TXT_DARK); c.setFont("Helvetica", 11)
            comp = wrap_lines(lbl, "Helvetica", 11, col_component - 10)
            c.drawString(table_x + 6, y, comp[0] if comp else lbl)

            # you/team/delta
            c.setFillColor(TXT_DARK)
            c.drawRightString(table_x + col_component + col_you - 6, y, str(you_i))
            c.setFillColor(TXT_MUTED)
            c.drawRightString(table_x + col_component + col_you + col_team - 6, y, str(team_avg))
            c.setFillColor(POS if delta_m >= 0 else NEG)
            c.drawRightString(table_x + table_w - 6, y, f"{delta_m:+d}")

            # light row divider
            y -= row_h
            c.setStrokeColor(RULE_LIGHT); c.setLineWidth(0.8)
            c.line(table_x, y + 5, table_x + table_w, y + 5)

            rows += 1

        # bottom rule to finish table
        c.setStrokeColor(RULE_LIGHT); c.setLineWidth(1)
        c.line(table_x, y + 5, table_x + table_w, y + 5)

        # ---------- Footer strip (white) ----------
        strip_h = 1.7 * cm
        c.setFillColor(white); c.rect(0, 0, W, strip_h, stroke=0, fill=1)
        c.setStrokeColor(RULE_LIGHT); c.line(0, strip_h, W, strip_h)

        # KC logo bottom-left only
        try:
            if logo_path and Path(logo_path).exists():
                lw, lh = 4.6 * cm, 1.2 * cm
                c.drawImage(
                    str(logo_path),
                    1.0 * cm,
                    0.25 * cm,
                    width=lw,
                    height=lh,
                    mask="auto",
                    preserveAspectRatio=True,
                )
        except Exception:
            pass

        # Footer text right
        c.setFillColor(TXT_MUTED); c.setFont("Helvetica", 9)
        c.drawRightString(W - 1.0 * cm, 0.9 * cm, "KC Overseas Education · Recognizing daily excellence")

        c.showPage(); c.save()
        return True, "ok"

    except Exception as e:
        return False, f"certificate-error: {e}"

def export_actions_excel(merged_df: pd.DataFrame, email: str, day: dt.date, xlsx_path: Path) -> tuple[bool, str]:
    """
    Excel with ONLY these columns (and order):
    Action (IST), Action Type, Action Group, Action Taken, Ack Number, Student ID,
    University, Partner Name, Assignee, Region
    """
    sub = merged_df[(merged_df["EmployeeEmail"]==email) & (merged_df["ActionDateIST_Date"]==day)].copy()

    base_cols = [
        "ActionDateIST","ActionType","Action Group","ActionTaken",
        "AcknowledgementNumber","StudentId","University","PartnerName","Assignee","Region"
    ]
    for c in base_cols:
        if c not in sub.columns:
            sub[c] = ""

    if "ActionDateIST" in sub.columns:
        sub["ActionDateIST"] = pd.to_datetime(sub["ActionDateIST"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M:%S")

    for c in ["ActionType","Action Group","ActionTaken","University","PartnerName","Assignee","Region"]:
        sub[c] = sub[c].astype(str).str.strip()

    # Ensure ActionTaken populated with priority SubType1 -> ActionTaken -> ActionType -> Action Group
    s1 = merged_df.loc[sub.index, "SubType1"].astype(str).str.strip() if "SubType1" in merged_df.columns else ""
    mask_blank = (sub["ActionTaken"].astype(str).str.strip() == "") | (sub["ActionTaken"].astype(str).str.lower().isin(["nan","none","null"]))
    if isinstance(s1, pd.Series):
        sub.loc[mask_blank & (s1 != ""), "ActionTaken"] = s1[mask_blank & (s1 != "")]
    # still blank?
    at = sub["ActionTaken"].astype(str).str.strip()
    mask_blank2 = (at == "") | (at.str.lower().isin(["nan","none","null"]))
    sub.loc[mask_blank2, "ActionTaken"] = sub.loc[mask_blank2, "ActionType"].astype(str).str.strip()
    at2 = sub["ActionTaken"].astype(str).str.strip()
    mask_still = (at2 == "") | (at2.str.lower().isin(["nan","none","null"]))
    sub.loc[mask_still, "ActionTaken"] = sub.loc[mask_still, "Action Group"].astype(str).str.strip()

    cols = [
        ("ActionDateIST", "Action (IST)"),
        ("ActionType", "Action Type"),
        ("Action Group", "Action Group"),
        ("ActionTaken", "Action Taken"),
        ("AcknowledgementNumber", "Ack Number"),
        ("StudentId", "Student ID"),
        ("University", "University"),
        ("PartnerName", "Partner Name"),
        ("Assignee", "Assignee"),
        ("Region", "Region"),
    ]
    use = sub[[c[0] for c in cols]].rename(columns=dict(cols))

    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as w:
        use.to_excel(w, index=False, sheet_name="Actions")
        ws = w.sheets["Actions"]
        for col_idx, col_name in enumerate(use.columns, start=1):
            max_len = max([len(str(col_name))] + [len(str(x)) for x in use[col_name].astype(str).head(300)])
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max(12, max_len + 2), 45)
    return True, f"rows={len(use)}"

# ======================== MAIL TRANSPORTS ======================= #
def _release_com(obj):
    try:
        if obj is not None:
            obj = None
            gc.collect()
    except Exception:
        pass

def send_outlook_html_with_attachments(to_addr: str, cc_addr: str|None, subject: str, html: str, attachments: list[Path],
                                       DRY_RUN=True, FROM_DISPLAY=""):
    import pythoncom, os
    import win32com.client as win32
    import win32_utils  # Import the safe dispatch helper
    pythoncom.CoInitialize()
    mail = ns = ol = None
    try:
        # Use safe_dispatch (which uses Dispatch + clean retry)
        # Replacing EnsureDispatch with safe_dispatch avoids hard-dependency on corrupted cache
        ol = win32_utils.safe_dispatch("Outlook.Application")
        ns = ol.GetNamespace("MAPI")
        mail = ol.CreateItem(0)  # MailItem

        # From (send-as; requires permission)
        if FROM_DISPLAY:
            try:
                mail.SentOnBehalfOfName = FROM_DISPLAY
            except Exception:
                pass

        # Body + signature (with cid:kc_sig)
        html_full = html + SIG_HTML
        mail.HTMLBody = html_full
        mail.Subject = subject
        mail.To = to_addr
        if cc_addr: mail.CC = cc_addr

        # Inline signature image (CID)
        try:
            if SIG_IMG_PATH and Path(SIG_IMG_PATH).exists():
                attach = mail.Attachments.Add(Source=str(SIG_IMG_PATH))
                pa = attach.PropertyAccessor
                # PR_ATTACH_CONTENT_ID
                pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "kc_sig")
                # Mark as inline
                pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001F", "inline")
                try:
                    pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", True)
                except Exception:
                    pass
        except Exception:
            pass

        # File attachments
        for p in attachments:
            if p and Path(p).exists():
                mail.Attachments.Add(Source=str(p))

        if DRY_RUN:
            if SAVE_DRAFTS_WHEN_DRY_RUN: mail.Save()
            else: mail.Display(False)
        else:
            for attempt in range(1, RETRY_ON_ERROR + 2):
                try:
                    mail.Send()
                    break
                except Exception:
                    if attempt > RETRY_ON_ERROR: raise
                    time.sleep(3)
    finally:
        _release_com(mail); _release_com(ns); _release_com(ol)
        try: pythoncom.CoUninitialize()
        except Exception: pass

def send_smtp_html_with_attachments(to_addr: str, cc_addr: str|None, subject: str, html: str, attachments: list[Path], DRY_RUN=True,
                                    SMTP_SERVER="", SMTP_PORT=587, SMTP_USERNAME="", SMTP_PASSWORD=""):
    # Unused when SEND_VIA='outlook' — kept for completeness
    import smtplib, mimetypes
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    msg = MIMEMultipart('mixed')
    sender = SMTP_USERNAME or f"no-reply@{SMTP_SERVER}"
    msg['From'] = sender; msg['To'] = to_addr
    if cc_addr: msg['Cc'] = cc_addr
    msg['Subject'] = subject
    alt = MIMEMultipart('alternative'); alt.attach(MIMEText(html, 'html', 'utf-8')); msg.attach(alt)
    for p in attachments:
        if not p or not Path(p).exists(): continue
        ctype, _ = mimetypes.guess_type(str(p)); 
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
        recipients = [to_addr] + ([cc_addr] if cc_addr else [])
        server.sendmail(sender, recipients, msg.as_string())

# ========================= EMAIL BODY ========================== #
def build_member_html(row: pd.Series, df_day: pd.DataFrame,
                      labels_all: dict, labels_kra_for_member: dict,
                      role_meas: pd.DataFrame,
                      role_leaderboard_html: str,
                      badges_today_html: str,
                      badges_mtd_html: str) -> str:
    date = row["ReportDate"]
    name = row.get("EmployeeName") or row.get("EmployeeEmail")
    role = row.get("Role",""); region = row.get("Region",""); subreg = row.get("SubRegion","")
    badges = str(row.get("Badges_Today","")).strip()
    kra = row.get("KRA_Score", 0.0)
    rrank = row.get("KRA_Rank_Role", 0); orank = row.get("KRA_Rank_Overall", 0)
    hit = int(row.get("Hit_Target_Today", 0)); streak = int(row.get("Daily_Streak", 0))

    badge_html = " ".join([_pill(b.strip()) for b in badges.split(",") if b.strip()]) \
                 if badges else _pill("No badges today", bg="#F3F4F6", fg="#6B7280")

    kra_color = ACCENT_SUCCESS if float(kra) >= 100 else (ACCENT_WARN if float(kra) >= 60 else ACCENT_DANGER)
    top_score = f"""
      <div style="background:{CARD_BG};border:1px solid {DIVIDER};border-radius:12px;padding:18px;">
        <div style="font:800 12px/1 Inter;color:{TEXT_MUTED};text-transform:uppercase;letter-spacing:.04em;">KRA Score</div>
        <div style="font:900 42px/1 Inter;color:{kra_color};margin:8px 0 6px 0;">{round(float(kra),1)}</div>
        <div style="color:{TEXT_MUTED};font:600 12px/1.2 Inter;">Role Rank: <span style="color:{TEXT_DARK}">#{_safe_int(rrank)}</span> · Overall Rank: <span style="color:{TEXT_DARK}">#{_safe_int(orank)}</span></div>
        <div style="margin-top:12px;">{_pill('Hit Target' if hit else 'Below Target', ACCENT_SUCCESS if hit else ACCENT_WARN, '#FFFFFF')} {_pill(f'Streak: {streak}', BRAND_ACCENT, '#FFFFFF')}</div>
      </div>
    """

    combined = combined_kra_section(role, row, role_meas, labels_all)
    mtd_tbl  = rolling_mtd_table(row, labels_kra_for_member, role, role_meas, HOLIDAYS)

    html = f"""
<!doctype html>
<html>
  <head><meta charset="utf-8"><meta name="x-apple-disable-message-reformatting"><title>KC Daily — {date}</title></head>
  <body style="margin:0;background:{SURFACE_BG};padding:18px 0;font-family:Inter, Segoe UI, Arial;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:920px;margin:0 auto;">
      <tr><td style="padding:0 18px 14px 18px;">
        <div style="background:{CARD_BG};border:1px solid {DIVIDER};border-radius:12px;padding:18px;">
          <div style="font:800 20px/1.2 Inter,Segoe UI,Arial;color:{BRAND_PRIMARY};">KC Daily — {date}</div>
          <div style="margin-top:6px;font:600 13px/1.4 Inter,Segoe UI,Arial;color:{TEXT_MUTED}">{name} · {role} · {region}{(' / '+subreg) if subreg else ''}</div>
          <div style="margin-top:12px;">{badge_html}</div>
        </div>
      </td></tr>

      <tr><td style="padding:0 18px 0 18px;">{top_score}</td></tr>

      <!-- Side-by-side tables (50/50) -->
      <tr>
        <td style="padding:0 18px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:18px 0;">
            <tr>
              <td style="width:50%;padding-right:8px;vertical-align:top;">{combined}</td>
              <td style="width:50%;padding-left:8px;vertical-align:top;">{mtd_tbl}</td>
            </tr>
          </table>
        </td>
      </tr>

      <tr><td style="padding:0 18px 8px 18px;">{role_leaderboard_html}</td></tr>

      <!-- Badges row: Today | Monthly -->
      <tr>
        <td style="padding:0 18px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:8px 0;">
            <tr>
              <td style="width:50%;padding-right:8px;vertical-align:top;">{badges_today_html}</td>
              <td style="width:50%;padding-left:8px;vertical-align:top;">{badges_mtd_html}</td>
            </tr>
          </table>
        </td>
      </tr>

      <tr><td style="padding:0 18px;">
        <div style="height:1px;background:{DIVIDER};margin:12px 0 16px 0;"></div>
        {signature_block()}
      </td></tr>
    </table>
  </body>
</html>
    """
    return html


# =========================== MAIN ============================ #

def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    team, measures_df, labels_all, labels_kra_only_union, role_meas, holidays, kra_meas_union = load_settings(SETTINGS_PATH)
    # set global holidays for MTD helper
    global HOLIDAYS
    HOLIDAYS = holidays

    summary = load_summary(SUMMARY_PATH, SUMMARY_SHEET)
    if summary.empty:
        print('No summary rows found.'); return
    merged_all = load_merged_actions()   # load ONCE from database

    all_days = summary['ReportDate'].dropna().tolist()
    work_days = resolve_pending_work_days(
        all_days, holidays, latest_working_day, report_type='staff_mailer', max_days=MAX_CATCHUP_DAYS
    )
    if not work_days:
        print('[INFO] No pending working days to send.'); return

    roster = team[["EmployeeEmail","EmployeeName","Role","Region","SubRegion","Manager","Include","SendTo"]].copy()

    def send_for_day(work_day: dt.date) -> bool:
        df_day = summary[summary['ReportDate'] == work_day].copy()
        sent_recipients = load_sent_recipients('staff_mailer', work_day)

        merged = roster.merge(df_day, on='EmployeeEmail', how='left', suffixes=('__team',''))
        for col in ['EmployeeName','Role','Region','SubRegion','Manager']:
            merged[col] = merged[col].combine_first(merged.get(f"{col}__team")).fillna('')
            if f"{col}__team" in merged.columns:
                merged.drop(columns=[f"{col}__team"], inplace=True)

        merged['__include_norm'] = merged['Include'].astype(bool)
        if DEFAULT_INCLUDE_IF_IN_SUMMARY:
            merged['__include_norm'] = np.where(merged['__include_norm'] | merged['ReportDate'].notna(), True, False)

        total = len(merged)
        incl_count = int(merged['__include_norm'].sum())
        print(f"[INFO] Report date (working day): {work_day} | roster rows: {total} | will include: {incl_count}")

        df_day_filt = merged[merged['__include_norm'] == True].copy()
        df_day_filt['__kra_val'] = pd.to_numeric(df_day_filt.get('KRA_Score', 0), errors='coerce').fillna(0.0)
        df_day_filt = df_day_filt[df_day_filt['__kra_val'] > 0].copy()
        df_day_filt = df_day_filt[df_day_filt['Role'].apply(lambda r: role_has_kra(r, role_meas))].copy()
        df_day_filt.drop(columns=['__kra_val'], errors='ignore', inplace=True)

        if df_day_filt.empty:
            print('[INFO] After skip rules: no recipients to send today.')
            return True

        role_labels = {}
        role_leaderboard_block = {}
        for role in sorted(df_day_filt['Role'].dropna().unique()):
            if not role_has_kra(role, role_meas):
                continue
            rms = role_meas[(role_meas['Role']==role) & (role_meas['IncludeInKRA'])].copy()
            codes = [str(x) for x in rms['MeasureCode'].dropna().tolist()]
            labels_for_role = {c: labels_all.get(c, c) for c in codes}
            role_labels[role] = labels_for_role
            role_leaderboard_block[role] = build_role_leaderboard(df_day_filt, role, labels_for_role)

        badges_today_html = build_badges_today_block(df_day_filt, labels_all)
        badges_mtd_html   = build_monthly_badges_block(summary, work_day, labels_all)

        sent = skipped = already_sent = failed = 0
        for _, r in df_day_filt.iterrows():
            to_addr = r.get('SendTo') or r.get('EmployeeEmail')
            if not to_addr or '@' not in str(to_addr):
                skipped += 1; continue
            to_addr_norm = str(to_addr).strip().lower()
            if to_addr_norm in sent_recipients:
                already_sent += 1
                continue

            name = r.get('EmployeeName') or r.get('EmployeeEmail')
            role = r.get('Role','')
            if not role_has_kra(role, role_meas):
                print(f"[SKIP] {name}: Role has no KRA components.")
                skipped += 1; continue
            if float(r.get('KRA_Score', 0) or 0) <= 0:
                print(f"[SKIP] {name}: KRA_Score == 0 (no activity today).")
                skipped += 1; continue

            labels_for_member = role_labels.get(role, {})
            role_block_html = role_leaderboard_block.get(role, '')

            html = build_member_html(
                row=r, df_day=df_day_filt, labels_all=labels_all, labels_kra_for_member=labels_for_member,
                role_meas=role_meas,
                role_leaderboard_html=role_block_html,
                badges_today_html=badges_today_html,
                badges_mtd_html=badges_mtd_html
            )

            safe = re.sub(r'[^a-zA-Z0-9]+', '_', str(name))[:60]
            html_path = OUT_DIR / f"{work_day.isoformat()}_{safe}.html"
            html_path.write_text(html, encoding='utf-8')

            attachments = []
            if ATTACH_PDF:
                pdf_report = OUT_DIR / f"{work_day.isoformat()}_{safe}_Report.pdf"
                ok_pdf, info = html_to_pdf(html, pdf_report)
                if ok_pdf:
                    attachments.append(pdf_report)
                else:
                    print(f"[WARN] PDF report generation failed for {name} ({info}).")

            badges_str = str(r.get('Badges_Today','')).strip()
            if ATTACH_CERT_PDF and badges_str:
                first_badge = badges_str.split(',')[0].strip() or 'Achievement'
                breakdown = []
                for code in labels_for_member:
                    v = int(r.get(code, 0))
                    if v>0:
                        breakdown.append((labels_all.get(code, code), v))
                breakdown.sort(key=lambda x: -x[1])
                peer_stats = build_peer_stats_for_role(df_day_filt, role, labels_for_member)
                cert_path = OUT_DIR / f"{work_day.isoformat()}_{safe}_Certificate.pdf"
                ok_c, cinf = make_badge_certificate(
                    name=name, role=r.get('Role',''), region=r.get('Region',''), subregion=r.get('SubRegion',''),
                    date=work_day, kra=float(r.get('KRA_Score',0.0)), rank_role=int(r.get('KRA_Rank_Role',0)),
                    top_breakdown=breakdown, badge_title=first_badge, pdf_path=cert_path, logo_path=LOGO_PATH,
                    labels_for_member=labels_for_member, peer_stats=peer_stats
                )
                if ok_c:
                    attachments.append(cert_path)
                else:
                    print(f"[WARN] Certificate failed for {name} ({cinf}).")

            xlsx_path = OUT_DIR / f"{work_day.isoformat()}_{safe}_Actions.xlsx"
            _ok_xl, _ = export_actions_excel(merged_all, r['EmployeeEmail'], work_day, xlsx_path)
            attachments.append(xlsx_path)

            subject = f"{SUBJECT_PREFIX} | {work_day.isoformat()} | {name} | KRA {round(float(r.get('KRA_Score',0)),1)} | Role #{int(r.get('KRA_Rank_Role',0))}"
            cc = None  # CC_MANAGER is False per request

            try:
                if SEND_VIA.lower()=='outlook':
                    send_outlook_html_with_attachments(to_addr, cc, subject, html, attachments,
                                                       DRY_RUN=DRY_RUN, FROM_DISPLAY=FROM_DISPLAY)
                else:
                    send_smtp_html_with_attachments(to_addr, cc, subject, html, attachments, DRY_RUN=DRY_RUN,
                                                    SMTP_SERVER=SMTP_SERVER, SMTP_PORT=SMTP_PORT,
                                                    SMTP_USERNAME=SMTP_USERNAME, SMTP_PASSWORD=SMTP_PASSWORD)
                sent += 1
                if DRY_RUN:
                    append_sent_detail('staff_mailer', work_day, to_addr_norm, status='dry_run', meta=f"employee={r.get('EmployeeEmail','')}")
                else:
                    append_sent_detail('staff_mailer', work_day, to_addr_norm, status='ok', meta=f"employee={r.get('EmployeeEmail','')}")
                if sent % BATCH_SIZE == 0:
                    time.sleep(SLEEP_BETWEEN_BATCHES_SEC); gc.collect()
                print(f"Prepared{' (DRY RUN)' if DRY_RUN else ''} -> {name} <{to_addr}>; files: {[p.name for p in attachments if p]}")
            except Exception as e:
                print(f"[ERROR] {name} <{to_addr}> -> {e}")
                traceback.print_exc()
                failed += 1
                append_sent_detail('staff_mailer', work_day, to_addr_norm, status='fail', meta=str(e)[:200])

        print(f"\\nDone. Members included: {sent}. Skipped: {skipped}. Already sent: {already_sent}. Failed: {failed}. Report date: {work_day}")
        return failed == 0

    for work_day in work_days:
        try:
            ok = send_for_day(work_day)
            status = 'dry_run' if DRY_RUN else ('ok' if ok else 'fail')
            append_sent_log('staff_mailer', work_day, status=status)
        except Exception as exc:
            append_sent_log('staff_mailer', work_day, status='fail')
            print(f"[ERROR] Failed to send for {work_day}: {exc}")
            traceback.print_exc()


if __name__ == '__main__':
    main()

