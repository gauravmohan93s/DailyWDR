# -*- coding: utf-8 -*-
"""
Centralized data and settings loader for Daily Work Done Report.
Ensures consistent data handling, deduplication, and normalization.
"""

from __future__ import annotations
from pathlib import Path
import datetime as dt
import pandas as pd
import numpy as np
import re
import sqlite3
from typing import Set, Dict, Any, Tuple, List

# Token normalization helpers
_token_re = re.compile(r"[^a-z0-9]+")

def _empty_if_placeholder(raw: str) -> str:
    raw = (raw or "").strip().lower()
    return "" if raw in {"nan", "none", "nil"} else raw

def norm_token(s: str) -> str:
    s = _empty_if_placeholder(str(s))
    s = s.replace("_", " ")
    s = _token_re.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def split_norm_list(s: str) -> list[str]:
    s = _empty_if_placeholder(str(s))
    return [norm_token(x) for x in s.split(";") if norm_token(x)]

def parse_saturday_pattern(raw: str) -> set[int]:
    """Parse strings like '1st & 3rd' / '2 & 4' into a set of Saturday ordinals."""
    tokens = re.findall(r"[1-4]", str(raw or ""))
    return {int(t) for t in tokens if t in {"1", "2", "3", "4"}}

def region_key(row) -> str:
    rg = str(row.get("Region","") or "").strip()
    sr = str(row.get("SubRegion","") or "").strip()
    # Team is a combination of Region and Sub Region
    if rg and sr: return f"{rg} — {sr}"
    return rg or sr or "Unknown"

def load_settings(path: Path) -> Dict[str, Any]:
    """Loads settings, preferring SQLite DB over Excel."""
    from reporting_config import LOCAL_DB_PATH
    db_path = Path(LOCAL_DB_PATH)
    if db_path.exists():
        try:
            conn = sqlite3.connect(db_path)
            cur = conn.cursor()
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='roster'")
            if cur.fetchone():
                team      = pd.read_sql("SELECT * FROM roster", conn)
                measures  = pd.read_sql("SELECT * FROM measures_config", conn)
                role_meas = pd.read_sql("SELECT * FROM role_measures", conn)
                
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='holidays'")
                hol_df = pd.read_sql("SELECT * FROM holidays", conn) if cur.fetchone() else pd.DataFrame()
                
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='agg_aliases'")
                aliases = pd.read_sql("SELECT * FROM agg_aliases", conn) if cur.fetchone() else pd.DataFrame()

                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='management_matrix'")
                matrix_df = pd.read_sql("SELECT * FROM management_matrix", conn) if cur.fetchone() else pd.DataFrame()

                conn.close()
                return _process_settings_dfs(team, measures, role_meas, hol_df, aliases, matrix_df)
            conn.close()
        except Exception as e:
            print(f"[DATA_LOADER] DB settings load failed: {e}")

    # Fallback to Excel
    if not path.exists(): raise FileNotFoundError(f"Settings not found: {path}")
    xls = pd.read_excel(path, sheet_name=None)
    return _process_settings_dfs(
        xls.get("Team"), xls.get("Measures"), xls.get("RoleMeasures"),
        xls.get("Holidays", pd.DataFrame()), xls.get("AggAliases", pd.DataFrame()),
        xls.get("ManagementMatrix", pd.DataFrame())
    )

def _process_settings_dfs(team, measures, role_meas, hol_df, aliases, matrix_df) -> Dict[str, Any]:
    # --- 1. TEAM ROSTER ---
    cols = ["EmployeeEmail","EmployeeName","Role","Region","SubRegion","Manager","Include","SendTo","SaturdayOffPattern","WeekOffs"]
    for c in cols:
        if c not in team.columns: team[c] = ""
    team["EmployeeEmail"] = team["EmployeeEmail"].astype(str).str.strip().str.lower()
    team = team.drop_duplicates(subset=["EmployeeEmail"], keep="first")
    team["Include"] = team["Include"].astype(str).str.strip().str.lower().map({"yes":True,"y":True,"1":True,"true":True,"t":True}).fillna(False)
    for c in ["Region","SubRegion","Role","Manager","EmployeeName","SendTo","SaturdayOffPattern","WeekOffs"]:
        team[c] = team[c].astype(str).str.strip()
    
    # --- 2. MANAGEMENT MATRIX SYNC ---
    if not matrix_df.empty:
        # Create a lookup key (Role|Region|SubRegion) -> ManagerEmail
        matrix_df["key"] = matrix_df.apply(lambda r: f"{norm_token(r.get('Role',''))}|{norm_token(r.get('Region',''))}|{norm_token(r.get('SubRegion',''))}", axis=1)
        matrix_map = {r["key"]: str(r.get("ManagerEmail","")).strip() for _, r in matrix_df.iterrows() if str(r.get("ManagerEmail","")).strip()}
        
        def sync_manager(row):
            k = f"{norm_token(row.get('Role',''))}|{norm_token(row.get('Region',''))}|{norm_token(row.get('SubRegion',''))}"
            return matrix_map.get(k, row.get("Manager", ""))
        
        team["Manager"] = team.apply(sync_manager, axis=1)

    # --- 3. MEASURES ---
    m_cols = ["MeasureCode","MeasureLabel","AggType","AggField","ActionTypeIn","AG_In","AG_NotIn","ExcludeAcksFrom","IncludeInKRA","BadgeEligible"]
    for c in m_cols:
        if c not in measures.columns: measures[c] = ""
        measures[c] = measures[c].fillna("")
    measures["IncludeInKRA"] = measures["IncludeInKRA"].astype(str).str.strip().str.upper().map({"Y":True,"YES":True,"1":True,"TRUE":True}).fillna(False)
    measures["BadgeEligible"] = measures["BadgeEligible"].astype(str).str.strip().str.upper().map({"Y":True,"YES":True,"1":True,"N":False,"NO":False,"0":False}).fillna(True)
    measures["ActionTypeIn_normlist"] = measures["ActionTypeIn"].apply(split_norm_list)
    measures["AG_In_normlist"]        = measures["AG_In"].apply(split_norm_list)
    measures["AG_NotIn_normlist"]     = measures["AG_NotIn"].apply(split_norm_list)
    measures["AggField_norm"]         = measures["AggField"].apply(norm_token)
    label_map = {str(r.MeasureCode): (str(r.MeasureLabel).strip() or str(r.MeasureCode)) for _, r in measures.iterrows()}
    badge_ok_map = {str(r.MeasureCode): bool(r.BadgeEligible) for _, r in measures.iterrows()}

    # --- 4. ROLE MEASURES ---
    for c in ["Weight","DailyTarget","MinFloor","Cap","DisplayOrder"]:
        if c not in role_meas.columns: role_meas[c] = 0.0
        role_meas[c] = pd.to_numeric(role_meas[c], errors="coerce").fillna(0.0)
    role_meas["IncludeInKRA"] = role_meas["IncludeInKRA"].astype(str).str.strip().str.upper().map({"Y":True,"YES":True,"1":True,"TRUE":True}).fillna(False)
    role_kra_codes = {}
    for role in sorted(role_meas["Role"].dropna().unique()):
        sub = role_meas[(role_meas["Role"]==role) & (role_meas["IncludeInKRA"]==True)].copy()
        if sub.empty: role_kra_codes[role] = []; continue
        sub["_row"] = np.arange(len(sub)); sub["__has"] = sub["DisplayOrder"].notna().astype(int)
        sub = sub.sort_values(["__has","DisplayOrder","_row"], ascending=[False,True,True])
        role_kra_codes[role] = [str(x) for x in sub["MeasureCode"].dropna().tolist()]

    # --- 5. HOLIDAYS ---
    holidays = set()
    for col in ["Date","HolidayDate","Holiday","Dt"]:
        if col in hol_df.columns:
            h = pd.to_datetime(hol_df[col], errors="coerce").dt.date.dropna()
            holidays = set(h.tolist()); break

    return {
        "team": team, "measures": measures, "role_meas": role_meas, "holidays": holidays,
        "label_map": label_map, "badge_ok_map": badge_ok_map, "role_kra_codes": role_kra_codes,
        "alias_map": {str(r["Name"]).strip(): str(r["Column"]).strip() for _, r in aliases.iterrows() if "Name" in aliases.columns},
        "matrix": matrix_df,
        "kra_meas_union": sorted(list(set(role_meas.loc[role_meas["IncludeInKRA"], "MeasureCode"].dropna().astype(str)))),
    }

def apply_config_filters(df: pd.DataFrame) -> pd.DataFrame:
    from reporting_config import load_app_config
    cfg = load_app_config(); out = df.copy()
    
    def _apply_list_filter(df, col, filter_str):
        if not filter_str: return df
        vals = [s.strip().lower() for s in filter_str.split(",") if s.strip()]
        if not vals: return df
        return df[df[col].astype(str).str.lower().str.contains('|'.join(vals))]

    out = _apply_list_filter(out, "EmployeeName", cfg.get("NAME_FILTER"))
    out = _apply_list_filter(out, "Role", cfg.get("ROLE_FILTER"))
    
    tf = str(cfg.get("TEAM_FILTER", "")).strip().lower()
    if tf:
        vals = [s.strip() for s in tf.split(",") if s.strip()]
        mask = out["Region"].astype(str).str.lower().str.contains('|'.join(vals)) | \
               out["SubRegion"].astype(str).str.lower().str.contains('|'.join(vals))
        out = out[mask]
    return out

def load_summary(path: Path, sheet: str = "summary_daily_all", report_date: dt.date = None) -> pd.DataFrame:
    from reporting_config import LOCAL_DB_PATH
    db_path = Path(LOCAL_DB_PATH)
    if db_path.exists():
        try:
            conn = sqlite3.connect(db_path)
            cur = conn.cursor()
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='daily_summary'")
            if cur.fetchone():
                q = "SELECT * FROM daily_summary"
                if report_date: df = pd.read_sql(f"{q} WHERE ReportDate = ?", conn, params=[str(report_date)])
                else: df = pd.read_sql(q, conn)
                conn.close()
                if not df.empty:
                    df["ReportDate"] = pd.to_datetime(df["ReportDate"]).dt.date
                    df["EmployeeEmail"] = df["EmployeeEmail"].astype(str).str.strip().str.lower()
                    for c in df.columns:
                        if c.startswith(("MTD_","Expected_","Ratio_","KPI_","KRA_Score")):
                            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
                    return df.drop_duplicates(subset=["ReportDate", "EmployeeEmail"])
            conn.close()
        except: pass
    if not path.exists(): return pd.DataFrame()
    df = pd.read_excel(path, sheet_name=sheet)
    df["ReportDate"] = pd.to_datetime(df["ReportDate"]).dt.date
    if report_date: df = df[df["ReportDate"] == report_date]
    df["EmployeeEmail"] = df["EmployeeEmail"].astype(str).str.strip().str.lower()
    return df.drop_duplicates(subset=["ReportDate", "EmployeeEmail"])

def get_filter_options():
    """Returns unique values for UI filters."""
    from reporting_config import LOCAL_DB_PATH
    try:
        conn = sqlite3.connect(LOCAL_DB_PATH)
        roles = sorted(pd.read_sql("SELECT DISTINCT Role FROM roster", conn)["Role"].dropna().tolist())
        regions = pd.read_sql("SELECT DISTINCT Region, SubRegion FROM roster", conn)
        teams = sorted(list(set(regions["Region"].dropna().tolist() + regions["SubRegion"].dropna().tolist())))
        names = sorted(pd.read_sql("SELECT DISTINCT EmployeeName FROM roster", conn)["EmployeeName"].dropna().tolist())
        conn.close()
        return {"roles": roles, "teams": teams, "names": names}
    except: return {"roles": [], "teams": [], "names": []}
