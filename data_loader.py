# -*- coding: utf-8 -*-
"""
Centralized data and settings loader for DailyWD-UK.
Ensures consistent data handling, deduplication, and normalization across all scripts.
"""

from __future__ import annotations
from pathlib import Path
import datetime as dt
import pandas as pd
import numpy as np
import re
from typing import Set, Dict, Any, Tuple

# Token normalization helpers (shared with summariser)
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

def load_settings(path: Path) -> Dict[str, Any]:
    """
    Loads and normalizes all settings from wd_settings.xlsx.
    Returns a dictionary containing DataFrames and processed configuration.
    """
    if not path.exists():
        raise FileNotFoundError(f"Settings file not found at {path}")

    xls = pd.read_excel(path, sheet_name=None)
    
    # 1. TEAM ROSTER
    team = xls.get("Team")
    if team is None:
        raise ValueError("Missing 'Team' sheet in settings.")
    
    # Fill missing columns
    for c in ["EmployeeEmail","EmployeeName","Role","Region","SubRegion","Manager","Include","SendTo","SaturdayOffPattern"]:
        if c not in team.columns: team[c] = ""
    
    # Basic cleanup
    team["EmployeeEmail"] = team["EmployeeEmail"].astype(str).str.strip().str.lower()
    
    # DEDUPLICATE: Ensure each email only appears once
    before_count = len(team)
    team = team.drop_duplicates(subset=["EmployeeEmail"], keep="first")
    if len(team) < before_count:
        print(f"[DATA_LOADER] Deduplicated team roster: {before_count} -> {len(team)} members")
    
    # Boolean and string normalization
    team["Include"] = (
        team["Include"].astype(str).str.strip().str.lower()
        .map({"yes":True,"y":True,"1":True,"true":True,"t":True})
        .fillna(False)
    )
    for c in ["Region","SubRegion","Role","Manager","EmployeeName","SendTo","SaturdayOffPattern"]:
        team[c] = team[c].astype(str).str.strip()
    
    # 2. MEASURES CONFIG
    measures = xls.get("Measures")
    if measures is None:
        raise ValueError("Missing 'Measures' sheet in settings.")
    
    m_cols = ["MeasureCode","MeasureLabel","AggType","AggField","ActionTypeIn","AG_In","AG_NotIn","ExcludeAcksFrom","IncludeInKRA","BadgeEligible"]
    for c in m_cols:
        if c not in measures.columns: measures[c] = ""
        measures[c] = measures[c].fillna("")
    
    measures["MeasureCode"]   = measures["MeasureCode"].astype(str).str.strip()
    measures["MeasureLabel"]  = measures["MeasureLabel"].astype(str).str.strip()
    measures["AggType"]       = measures["AggType"].astype(str).str.strip().str.upper()
    measures["AggField"]      = measures["AggField"].astype(str).str.strip()
    measures["IncludeInKRA"]  = measures["IncludeInKRA"].astype(str).str.strip().str.upper().map({"Y":True,"YES":True}).fillna(False)
    measures["BadgeEligible"] = (
        measures["BadgeEligible"].astype(str).str.strip().str.upper()
        .map({"Y":True,"YES":True,"1":True,"N":False,"NO":False,"0":False})
        .fillna(True)
    )
    
    # Pre-calculate token lists for summariser
    measures["ActionTypeIn_normlist"] = measures["ActionTypeIn"].apply(split_norm_list)
    measures["AG_In_normlist"]        = measures["AG_In"].apply(split_norm_list)
    measures["AG_NotIn_normlist"]     = measures["AG_NotIn"].apply(split_norm_list)
    measures["AggField_norm"]         = measures["AggField"].apply(norm_token)
    
    # Helper maps for display
    label_map = {str(r.MeasureCode): (str(r.MeasureLabel).strip() or str(r.MeasureCode)) for _, r in measures.iterrows()}
    badge_ok_map = {str(r.MeasureCode): bool(r.BadgeEligible) for _, r in measures.iterrows()}

    # 3. ROLE MEASURES
    role_meas = xls.get("RoleMeasures")
    if role_meas is None:
        raise ValueError("Missing 'RoleMeasures' sheet in settings.")
    
    r_cols = ["Role","MeasureCode","Weight","DailyTarget","MinFloor","Cap","IncludeInKRA","DisplayOrder"]
    for c in r_cols:
        if c not in role_meas.columns: role_meas[c] = ""
        if c in ["Weight","DailyTarget","MinFloor","Cap","DisplayOrder"]:
            role_meas[c] = pd.to_numeric(role_meas[c], errors="coerce")
    
    role_meas["Role"] = role_meas["Role"].astype(str).str.strip()
    role_meas["MeasureCode"] = role_meas["MeasureCode"].astype(str).str.strip()
    role_meas["IncludeInKRA"] = role_meas["IncludeInKRA"].astype(str).str.strip().str.upper().map({"Y":True,"YES":True}).fillna(False)
    
    # Role -> List of KRA MeasureCodes (sorted by DisplayOrder)
    role_kra_codes = {}
    roles = sorted(role_meas["Role"].dropna().unique())
    for role in roles:
        sub = role_meas[(role_meas["Role"]==role) & (role_meas["IncludeInKRA"]==True)].copy()
        if sub.empty:
            role_kra_codes[role] = []
            continue
        sub["_row"] = np.arange(len(sub))
        sub["__has"] = sub["DisplayOrder"].notna().astype(int)
        sub = sub.sort_values(["__has","DisplayOrder","_row"], ascending=[False,True,True])
        role_kra_codes[role] = [str(x) for x in sub["MeasureCode"].dropna().tolist()]

    # 4. HOLIDAYS
    holidays = set()
    hol_df = xls.get("Holidays", pd.DataFrame())
    if not hol_df.empty:
        for col in ["Date","HolidayDate","Holiday","Dt"]:
            if col in hol_df.columns:
                h = pd.to_datetime(hol_df[col], errors="coerce").dt.date.dropna()
                holidays = set(h.tolist())
                break

    # 5. AGG ALIASES (Summariser only)
    aliases = xls.get("AggAliases", pd.DataFrame(columns=["Name","Column"]))
    alias_map = {str(r["Name"]).strip(): str(r["Column"]).strip() for _, r in aliases.iterrows() if pd.notna(r.get("Name"))}

    # 6. ROLE TARGETS (Optional)
    role_targets = xls.get("RoleTargets", pd.DataFrame(columns=["Role","KRATargetScore"])).copy()
    if "Role" not in role_targets.columns: role_targets["Role"] = ""
    if "KRATargetScore" not in role_targets.columns: role_targets["KRATargetScore"] = np.nan
    role_targets["Role"] = role_targets["Role"].astype(str).str.strip()
    role_targets["KRATargetScore"] = pd.to_numeric(role_targets["KRATargetScore"], errors="coerce")

    return {
        "team": team,
        "measures": measures,
        "role_meas": role_meas,
        "holidays": holidays,
        "label_map": label_map,
        "badge_ok_map": badge_ok_map,
        "role_kra_codes": role_kra_codes,
        "alias_map": alias_map,
        "role_targets": role_targets,
        "kra_meas_union": sorted(list(set(role_meas.loc[role_meas["IncludeInKRA"], "MeasureCode"].dropna().astype(str)))),
    }

def load_summary(path: Path, sheet: str = "summary_daily_all", report_date: dt.date = None) -> pd.DataFrame:
    """
    Loads and deduplicated daily summary data.
    Prefers SQLite DB if available, falls back to Excel.
    """
    # 1. TRY DATABASE FIRST
    from reporting_config import LOCAL_DB_PATH
    db_path = Path(LOCAL_DB_PATH)
    if db_path.exists():
        try:
            import sqlite3
            conn = sqlite3.connect(db_path)
            # Check if table exists
            cur = conn.cursor()
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='daily_summary'")
            if cur.fetchone():
                if report_date:
                    df = pd.read_sql("SELECT * FROM daily_summary WHERE ReportDate = ?", conn, params=[str(report_date)])
                else:
                    df = pd.read_sql("SELECT * FROM daily_summary", conn)
                conn.close()
                if not df.empty:
                    df["ReportDate"] = pd.to_datetime(df["ReportDate"]).dt.date
                    df["EmployeeEmail"] = df["EmployeeEmail"].astype(str).str.strip().str.lower()
                    # Basic cleanup for numeric columns that might have been stored as text
                    for c in df.columns:
                        if c.startswith(("MTD_","Expected_","Ratio_","KPI_","KRA_Score")):
                            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
                    return df.drop_duplicates(subset=["ReportDate", "EmployeeEmail"], keep="first")
            conn.close()
        except Exception as e:
            print(f"[DATA_LOADER] DB read failed, falling back to Excel: {e}")

    # 2. FALLBACK TO EXCEL (No SQL filtering support for Excel fallback to keep logic simple)
    if not path.exists():
        return pd.DataFrame()

    df = pd.read_excel(path, sheet_name=sheet)
    if df.empty:
        return df

    # Basic cleanup
    df["ReportDate"] = pd.to_datetime(df["ReportDate"]).dt.date
    if report_date:
        df = df[df["ReportDate"] == report_date].copy()
    
    df["EmployeeEmail"] = df["EmployeeEmail"].astype(str).str.strip().str.lower()
    
    # Fill missing string columns
    for c in ["EmployeeName","Role","Region","SubRegion","Manager","Badges_Today"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).str.strip()
            
    # DEDUPLICATE: Ensure each employee has only one record per day to prevent inflated totals
    before = len(df)
    df = df.drop_duplicates(subset=["ReportDate", "EmployeeEmail"], keep="first")
    if len(df) < before:
        print(f"[DATA_LOADER] Cleaned duplicate records in summary: {before} -> {len(df)}")
        
    return df
