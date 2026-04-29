# -*- coding: utf-8 -*-
"""
WD Summariser — All Days + Daily Announcement (single-sheet system, gamified)
- Builds/refreshes single summary sheet: summary_all_days.xlsx / 'summary_daily_all'
- Excludes Sundays + Holidays from 'Expected' (rolling working-day targets)
- Adds ranks, Top3, streaks, badges, uniqueness signals (already in the daily rows)
- Generates a DAILY Markdown announcement (console + .md file; optional webhook post)
- OPEN for new measures: reads dynamically from wd_settings.xlsx
"""

from __future__ import annotations
from pathlib import Path
import datetime as dt
import re
import json
import numpy as np
import pandas as pd

from reporting_config import REPORT_DATE_OVERRIDE, LOCAL_DB_PATH
from data_loader import (
    load_settings, load_summary,
    norm_token, split_norm_list, parse_saturday_pattern
)

# ---------------------- Paths – EDIT THESE ---------------------- #
MERGED_PATH   = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\raw\merged\merged_actions.xlsx")
SETTINGS_PATH = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\setting\wd_settings.xlsx")
OUT_ALL_PATH  = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\raw\summary_all_days.xlsx")
ANNOUNCE_DIR  = OUT_ALL_PATH.parent  # write announcement alongside summary

# ---------------------- Announcement config ---------------------- #
ANNOUNCE_DATE = "yesterday"   # "yesterday" (IST) means "last working day"; or "YYYY-MM-DD"
POST_ANNOUNCEMENT = False     # set True to POST to webhook
ANNOUNCE_WEBHOOK_URL = ""     # Slack/Teams incoming webhook (optional)

# ---------------------- Behaviour toggles ---------------------- #
IGNORE_CAP = True                 # do NOT cap KRA per-measure contribution
UNIQUENESS_HIGH_THRESH = 0.60     # 60% unique/total is "high"
BADGE_MEASURE_TOP_N = 5           # Top-N per measure/day for badges
STREAK_MILESTONES = [3, 5, 10, 20]

# ---------------------- Validation ---------------------- #
def validate_settings(measures: pd.DataFrame, role_meas: pd.DataFrame) -> dict:
    report = {"status":"OK","issues":[]}
    valid_types = {"ROW_COUNT","DISTINCT_COUNT"}
    bad_types = measures[~measures["AggType"].astype(str).str.upper().isin(valid_types)]
    if not bad_types.empty:
        report["issues"].append("Invalid AggType for: " + ", ".join(bad_types["MeasureCode"].astype(str)))
    
    check = measures[measures["AggType"].astype(str).str.upper()=="DISTINCT_COUNT"]
    # AggField_norm is provided by data_loader
    valid_fields = {"ack","student","partner","university"}
    bad_fields = check[~check["AggField_norm"].astype(str).isin(valid_fields)]
    if not bad_fields.empty:
        report["issues"].append("Invalid AggField (DISTINCT_COUNT) for: " + ", ".join(bad_fields["MeasureCode"].astype(str)))
    
    known = set(measures["MeasureCode"].astype(str))
    unknown_refs = role_meas[~role_meas["MeasureCode"].astype(str).isin(known)]
    if not unknown_refs.empty:
        report["issues"].append("RoleMeasures refer to unknown MeasureCode(s): " + ", ".join(sorted(unknown_refs["MeasureCode"].astype(str).unique())))
    
    dup = role_meas.groupby(["Role","MeasureCode"]).size().reset_index(name="count")
    dup = dup[dup["count"]>1]
    if not dup.empty:
        txt = ", ".join([f"{r.Role}/{r.MeasureCode}" for _, r in dup.iterrows()])
        report["issues"].append(f"Duplicate Role+Measure pairs: {txt}")
    
    sums = role_meas.groupby("Role")["Weight"].sum().reset_index()
    not_one = sums[(sums["Weight"] < 0.98) | (sums["Weight"] > 1.02)]
    if not not_one.empty:
        txt = "; ".join([f"{r.Role}={r.Weight:.3f}" for _, r in not_one.iterrows()])
        report["issues"].append(f"Weight sums not ~1.0 for roles: {txt} (normalized in scoring)")
    
    if report["issues"]:
        report["status"] = "WARN"
    return report

# ---------------------- Database Connection ---------------------- #
def get_local_db_engine():
    """Returns the engine for the local SQLite database."""
    from sqlalchemy import create_engine
    db_path = Path(LOCAL_DB_PATH)
    if not db_path.exists():
        raise FileNotFoundError(f"Database not found at {db_path}. Please run the extractor script to create it.")
    return create_engine(f"sqlite:///{db_path}")

# ---------------------- Data Load (from SQLite) ---------------------- #
def load_merged() -> pd.DataFrame:
    """Loads all actions from the local SQLite database."""
    print("[DB] Loading data from kc_reports.db...")
    engine = get_local_db_engine()
    df = pd.read_sql(f"SELECT * FROM actions", engine, parse_dates=["ActionDate"])
    print(f"[DB] Loaded {len(df)} records.")

    df["ActionDate"] = df["ActionDate"].dt.tz_localize('UTC')
    df["ActionDateIST_Date"] = df["ActionDate"].dt.tz_convert("Asia/Kolkata").dt.date
    df["ActionDateIST_Date"] = df["ActionDateIST_Date"].where(pd.notna(df["ActionDate"]), pd.NaT)
    
    df["action_type"]  = df.get("ActionType","").astype(str).map(norm_token)
    df["action_group"] = df.get("Action Group","").astype(str).map(norm_token)
    
    if "PartnerCode" not in df.columns:
        df["PartnerCode"] = ""
    df["PartnerKey"] = np.where(
        df["PartnerCode"].astype(str).str.strip()!="",
        df["PartnerCode"].astype(str).str.strip(),
        df["PartnerName"].astype(str).str.strip().str.lower()
    )
    return df

# ---------------------- Working-days calendar ---------------------- #
def build_working_calendar(min_date: dt.date, max_date: dt.date, holidays: set[dt.date]) -> pd.DataFrame:
    rng = pd.date_range(min_date, max_date, freq="D").date
    cal = pd.DataFrame({"ReportDate": rng})
    ts = pd.to_datetime(cal["ReportDate"])
    cal["IsSunday"]  = ts.dt.dayofweek == 6
    cal["IsSaturday"] = ts.dt.dayofweek == 5
    cal["SaturdayOrdinal"] = np.where(cal["IsSaturday"], ((ts.dt.day - 1) // 7) + 1, 0).astype(int)
    cal["IsHoliday"] = cal["ReportDate"].isin(holidays) if holidays else False
    cal["IsWorkingDay"] = ~(cal["IsSunday"] | cal["IsHoliday"])
    cal["Year"] = pd.to_datetime(cal["ReportDate"]).dt.year
    cal["Month"] = pd.to_datetime(cal["ReportDate"]).dt.month
    cal["WorkingDaysInMonth"] = cal.groupby(["Year","Month"])["IsWorkingDay"].transform("sum").astype(int)
    cal["WorkingDaysElapsed"] = cal.groupby(["Year","Month"])["IsWorkingDay"].cumsum().astype(int)
    return cal[["ReportDate","WorkingDaysInMonth","WorkingDaysElapsed","IsWorkingDay","IsSaturday","SaturdayOrdinal"]]

# ---------------------- Measure Engine (all days) ---------------------- #
def filter_for_measure(df0: pd.DataFrame, spec: pd.Series) -> pd.DataFrame:
    atypes = set(spec["ActionTypeIn_normlist"])
    ag_in  = set(spec["AG_In_normlist"])
    ag_ex  = set(spec["AG_NotIn_normlist"])
    m = pd.Series(True, index=df0.index)
    if atypes: m &= df0["action_type"].isin(atypes)
    if ag_in:  m &= df0["action_group"].isin(ag_in)
    if ag_ex:  m &= ~df0["action_group"].isin(ag_ex)
    return df0.loc[m].copy()

def compute_kpis_all_days(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["ActionDateIST_Date","EmployeeEmail"], dropna=False)
    out = pd.DataFrame({
        "KPI_TotalActions": g.size()
    })
    out["KPI_UniqueApplications"] = g["AcknowledgementNumber"].nunique(dropna=True)
    out["KPI_UniqueStudents"]     = g["StudentId"].nunique(dropna=True)
    out["KPI_UniquePartners"]     = g["PartnerKey"].nunique(dropna=True)
    out["KPI_UniqueUniversities"] = g["University"].nunique(dropna=True)
    out = out.reset_index().rename(columns={"ActionDateIST_Date":"ReportDate"})
    return out

def compute_measures_all_days(df: pd.DataFrame, measures: pd.DataFrame, alias_map: dict):
    member_col = "EmployeeEmail"; date_col = "ActionDateIST_Date"
    col_for = {
        "ack": alias_map.get("Ack","AcknowledgementNumber"),
        "student": alias_map.get("Student","StudentId"),
        "partner": "PartnerKey",
        "university": alias_map.get("University","University"),
        "": None
    }

    values_frames = []
    ack_sets = {}

    all_pairs = df[[date_col, member_col]].drop_duplicates().sort_values([date_col, member_col])
    if all_pairs.empty:
        mi_all = pd.MultiIndex.from_tuples([], names=["ReportDate", member_col])
    else:
        all_pairs_ren = all_pairs.rename(columns={date_col: "ReportDate"})
        mi_all = pd.MultiIndex.from_frame(all_pairs_ren[["ReportDate", member_col]])

    for _, spec in measures.iterrows():
        code     = str(spec["MeasureCode"])
        aggtype  = str(spec["AggType"]).upper()
        aggfield = str(spec["AggField_norm"])

        dff = filter_for_measure(df, spec)
        if dff.empty:
            zeros = pd.Series(0, index=mi_all, name=code)
            values_frames.append(zeros)
            if aggtype == "DISTINCT_COUNT" and aggfield == "ack":
                ack_sets[code] = pd.Series([set()]*len(mi_all), index=mi_all, name=code)
            continue

        g = dff.groupby([date_col, member_col], dropna=False)
        if aggtype == "ROW_COUNT":
            s = g.size().rename(code)
        elif aggtype == "DISTINCT_COUNT":
            colname = col_for.get(aggfield)
            if not colname:
                raise ValueError(f"Measure '{code}' needs a valid AggField.")
            s = g[colname].nunique(dropna=True).rename(code)
            if aggfield == "ack":
                ack_sets[code] = g[colname].apply(lambda x: set(x.dropna().astype(str))).rename(code)
        else:
            raise ValueError(f"Unsupported AggType: {aggtype}")

        s.index = s.index.set_names([date_col, member_col])
        s.index = pd.MultiIndex.from_tuples([(i[0], i[1]) for i in s.index], names=["ReportDate", member_col])
        s = s.reindex(mi_all, fill_value=0)
        values_frames.append(s)

        if aggtype == "DISTINCT_COUNT" and aggfield == "ack":
            ack_series = ack_sets.get(code, pd.Series(dtype=object))
            if not isinstance(ack_series, pd.Series):
                ack_series = pd.Series(dtype=object)
            if isinstance(ack_series.index, pd.MultiIndex):
                ack_series.index = pd.MultiIndex.from_tuples(
                    [(i[0], i[1]) for i in ack_series.index], names=["ReportDate", member_col]
                )
            ack_series = ack_series.reindex(mi_all).apply(lambda x: x if isinstance(x, set) else set())
            ack_sets[code] = ack_series

    values_wide = pd.concat(values_frames, axis=1) if values_frames else pd.DataFrame(index=mi_all)
    values_wide = values_wide.reset_index()
    day_totals = values_wide.groupby("ReportDate").sum(numeric_only=True)
    return values_wide, day_totals, ack_sets

def apply_ack_exclusions_all_days(values_wide: pd.DataFrame, ack_sets: dict, measures: pd.DataFrame):
    mi = pd.MultiIndex.from_frame(values_wide[["ReportDate","EmployeeEmail"]], names=["ReportDate","EmployeeEmail"])

    def ensure_set_series(s: pd.Series | None) -> pd.Series:
        if s is None or not isinstance(s, pd.Series):
            return pd.Series(index=mi, dtype=object).apply(lambda _: set())
        s = s.reindex(mi)
        return s.apply(lambda x: x if isinstance(x, set) else set())

    for _, spec in measures.iterrows():
        code = str(spec["MeasureCode"])
        if str(spec["AggType"]).upper() != "DISTINCT_COUNT" or str(spec["AggField_norm"]) != "ack":
            continue
        excl = str(spec["ExcludeAcksFrom"]).strip()
        if not excl:
            continue
        pool = ensure_set_series(ack_sets.get(code))
        sub  = ensure_set_series(ack_sets.get(excl))
        adj = (pool.combine(sub, lambda a, b: a - b)).apply(lambda s: len(s))
        values_wide.loc[:, code] = adj.values
    return values_wide

# ---------------------- KRA & Ranking ---------------------- #
def compute_kra_per_day(df_day: pd.DataFrame, role_meas: pd.DataFrame, ignore_cap: bool = True) -> pd.Series:
    kra = pd.Series(0.0, index=df_day.index)
    for role, cfg in role_meas.groupby("Role"):
        cfg = cfg.copy()
        if cfg["Weight"].sum() > 0:
            cfg["Weight"] = cfg["Weight"] / cfg["Weight"].sum()
        idx = (df_day["Role"] == role)
        if not idx.any(): 
            continue
        part = pd.Series(0.0, index=df_day.index[idx])
        for _, rr in cfg.iterrows():
            mcode = rr["MeasureCode"]
            if mcode not in df_day.columns:
                continue
            val = df_day.loc[idx, mcode].astype(float)
            denom = rr["DailyTarget"] if rr["DailyTarget"] and rr["DailyTarget"] > 0 else 1.0
            norm = val / denom
            if rr["MinFloor"] and rr["MinFloor"]>0:
                floor = rr["MinFloor"] / denom
                norm = np.where(val>0, np.maximum(norm, floor), norm)
            if (not ignore_cap) and rr["Cap"] and rr["Cap"]>0:
                norm = np.minimum(norm, rr["Cap"])
            part = part + (norm * float(rr["Weight"]))
        kra.loc[idx] = (part * 100.0)
    return kra.round(1)

# ---------------------- MTD (rolling) & targets (STRICT from daily) ---------------------- #
def add_mtd_and_targets_workdays(daily: pd.DataFrame,
                                 role_meas: pd.DataFrame,
                                 measure_codes: list[str],
                                 work_cal: pd.DataFrame | None) -> pd.DataFrame:
    out = daily.copy()
    needs_calendar = not {"WorkingDaysInMonth","WorkingDaysElapsed","IsWorkingDay"}.issubset(out.columns)
    if work_cal is not None and needs_calendar:
        merge_cols = ["ReportDate"]
        if "EmployeeEmail" in work_cal.columns and "EmployeeEmail" in out.columns:
            merge_cols.append("EmployeeEmail")
        cols_keep = [c for c in ["WorkingDaysInMonth","WorkingDaysElapsed","IsWorkingDay"] if c in work_cal.columns]
        out = out.merge(work_cal[merge_cols + cols_keep], on=merge_cols, how="left")
    out = out.sort_values(["EmployeeEmail","ReportDate"]).reset_index(drop=True)

    out["Year"]  = pd.to_datetime(out["ReportDate"]).dt.year
    out["Month"] = pd.to_datetime(out["ReportDate"]).dt.month
    grp = out.groupby(["EmployeeEmail","Year","Month"], group_keys=False)
    if ("WorkingDaysElapsed" not in out.columns) or out["WorkingDaysElapsed"].isna().any():
        out["WorkingDaysElapsed"] = grp["IsWorkingDay"].cumsum().astype(int)
    if ("WorkingDaysInMonth" not in out.columns) or out["WorkingDaysInMonth"].isna().any():
        out["WorkingDaysInMonth"] = grp["IsWorkingDay"].transform("sum").astype(int)

    tgt_map = {(str(r.Role), str(r.MeasureCode)): float(r.DailyTarget) for _, r in role_meas.iterrows()}

    for code in measure_codes:
        if code not in out.columns:
            out[code] = 0
        out[f"MTD_{code}"] = grp[code].cumsum().astype(float)
        dtarget = out.apply(lambda r: tgt_map.get((str(r["Role"]), str(code)), 0.0), axis=1)
        out[f"Expected_MTD_{code}"] = (dtarget * out["WorkingDaysElapsed"]).astype(float)

    return out.drop(columns=["Year","Month"])

def build_measure_families(measures: pd.DataFrame) -> list[dict]:
    def key_of(row):
        parts = []
        for c in ["ActionTypeIn_normlist","AG_In_normlist","AG_NotIn_normlist"]:
            vals = tuple(sorted(row[c])) if isinstance(row[c], list) else ()
            parts.append("|".join(vals))
        return "||".join(parts)
    m = measures.copy()
    m["__famkey"] = m.apply(key_of, axis=1)
    families = []
    for famkey, df in m.groupby("__famkey"):
        row_meas = df[df["AggType"].str.upper()=="ROW_COUNT"]["MeasureCode"].tolist()
        dst_meas = df[df["AggType"].str.upper()=="DISTINCT_COUNT"]["MeasureCode"].tolist()
        families.append({"famkey": famkey, "rows": row_meas, "distinct": dst_meas})
    return families

def add_uniqueness_ratios(daily: pd.DataFrame, measures: pd.DataFrame, thr: float) -> pd.DataFrame:
    out = daily.copy()
    fams = build_measure_families(measures)
    for fam in fams:
        rows_list, distinct_list = fam["rows"], fam["distinct"]
        if not rows_list or not distinct_list:
            continue
        rows_code = rows_list[0]
        uniq_code = distinct_list[0]
        if rows_code in out.columns and uniq_code in out.columns:
            ratio_col = f"Ratio_{uniq_code}_per_{rows_code}"
            out[ratio_col] = np.where(out[rows_code].astype(float)>0,
                                      out[uniq_code].astype(float)/out[rows_code].astype(float),
                                      0.0).round(3)
            out[f"IsHighUniq_{uniq_code}"] = (out[ratio_col] >= thr).astype(int)
    return out

# ---------------------- Build ALL days (single sheet) ---------------------- #
def build_all_days(settings_path: Path, out_path: Path,
                   ignore_cap: bool = True, uniq_thr: float = 0.60):
    cfg = load_settings(settings_path)
    team = cfg["team"]
    measures = cfg["measures"]
    role_meas = cfg["role_meas"]
    alias_map = cfg["alias_map"]
    holidays = cfg["holidays"]
    role_targets = cfg["role_targets"]
    label_map = cfg["label_map"]
    badge_ok = cfg["badge_ok_map"]

    val = validate_settings(measures, role_meas)
    print(f"[Validator] Status: {val['status']}")
    for issue in val["issues"]:
        print("  -", issue)

    df = load_merged()

    kpi_all = compute_kpis_all_days(df)
    vals_wide, day_totals, ack_sets = compute_measures_all_days(df, measures, alias_map)
    vals_wide = apply_ack_exclusions_all_days(vals_wide, ack_sets, measures)

    valid_dates = df["ActionDateIST_Date"].dropna()
    if valid_dates.empty:
        raise ValueError("No valid action dates found in data")
    min_d, max_d = valid_dates.min(), valid_dates.max()
    full_dates = pd.DataFrame({"ReportDate": pd.date_range(min_d, max_d, freq="D").date})

    work_cal = build_working_calendar(min_d, max_d, holidays)
    sat_off_map = {
        str(r.EmployeeEmail).strip().lower(): parse_saturday_pattern(r.get("SaturdayOffPattern", ""))
        for _, r in team.iterrows()
    }

    roster = team[["EmployeeEmail","EmployeeName","Role","Region","SubRegion","Manager","Include","SaturdayOffPattern"]].copy()
    base_cal = work_cal.rename(columns={"IsWorkingDay":"BaseIsWorkingDay"})
    all_members = (
        full_dates.assign(key=1)
        .merge(roster.assign(key=1), on="key")
        .drop(columns="key")
        .merge(base_cal, on="ReportDate", how="left")
    )

    def is_member_workday(row) -> bool:
        base = bool(row.get("BaseIsWorkingDay", False))
        if not base:
            return False
        if bool(row.get("IsSaturday", False)):
            offs = sat_off_map.get(str(row.get("EmployeeEmail", "")).lower(), set())
            if int(row.get("SaturdayOrdinal", 0)) in offs:
                return False
        return True

    all_members["IsWorkingDay"] = all_members.apply(is_member_workday, axis=1)
    all_members["Year"] = pd.to_datetime(all_members["ReportDate"]).dt.year
    all_members["Month"] = pd.to_datetime(all_members["ReportDate"]).dt.month
    grp_wd = all_members.groupby(["EmployeeEmail","Year","Month"], group_keys=False)
    all_members["WorkingDaysElapsed"] = grp_wd["IsWorkingDay"].cumsum().astype(int)
    all_members["WorkingDaysInMonth"] = grp_wd["IsWorkingDay"].transform("sum").astype(int)
    all_members.drop(columns=["Year","Month","BaseIsWorkingDay","IsSaturday","SaturdayOrdinal"], inplace=True)
    work_cal_member = all_members[["ReportDate","EmployeeEmail","WorkingDaysInMonth","WorkingDaysElapsed","IsWorkingDay"]].copy()

    daily = (all_members
             .merge(kpi_all, on=["ReportDate","EmployeeEmail"], how="left")
             .merge(vals_wide, on=["ReportDate","EmployeeEmail"], how="left"))

    for c in ["EmployeeName","Role","Region","SubRegion","Manager"]:
        daily[c] = daily[c].fillna("")
    measure_codes = list(measures["MeasureCode"])
    kpi_cols = ["KPI_TotalActions","KPI_UniqueApplications","KPI_UniqueStudents","KPI_UniquePartners","KPI_UniqueUniversities"]
    for c in kpi_cols + measure_codes:
        if c in daily.columns:
            daily[c] = pd.to_numeric(daily[c], errors="coerce").fillna(0).astype(int)
        else:
            daily[c] = 0

    daily = daily.sort_values(["ReportDate","EmployeeEmail"]).reset_index(drop=True)
    kra_scores = []
    for _, df_day in daily.groupby("ReportDate", sort=False):
        df_d = df_day.copy()
        df_d["KRA_Score"] = compute_kra_per_day(df_d, role_meas, ignore_cap=ignore_cap)
        kra_scores.append(df_d["KRA_Score"])
    daily["KRA_Score"] = pd.concat(kra_scores).sort_index().values

    daily["KRA_Rank_Role"] = daily.groupby(["ReportDate","Role"])["KRA_Score"].rank(method="dense", ascending=False).astype(int)
    daily["KRA_Rank_Overall"] = daily.groupby("ReportDate")["KRA_Score"].rank(method="dense", ascending=False).astype(int)
    for code in measure_codes:
        if code in daily.columns:
            daily[f"Rank_{code}"] = daily.groupby("ReportDate")[code].rank(method="dense", ascending=False).astype(int)
            daily[f"IsTop3_{code}"] = ((daily[code] > 0) &
                                       (daily[f"Rank_{code}"] <= BADGE_MEASURE_TOP_N) &
                                       badge_ok.get(code, True)).astype(int)

    day_totals = day_totals.rename_axis("ReportDate").reset_index()
    daily = daily.merge(day_totals, on="ReportDate", suffixes=("", "__DAYTOT"), how="left")
    for code in measure_codes:
        tot_col = f"{code}__DAYTOT"; pct_col = f"{code}_PctOfDay"
        if code in daily.columns and tot_col in daily.columns:
            total = daily[tot_col].replace(0, np.nan)
            daily[pct_col] = np.where(total.notna(), (daily[code] / total * 100.0).round(1), 0.0)
            daily.drop(columns=[tot_col], inplace=True)

    daily = add_mtd_and_targets_workdays(daily, role_meas, measure_codes, work_cal_member)
    daily = add_uniqueness_ratios(daily, measures, thr=uniq_thr)

    rt_map = {str(r.Role): (float(r.KRATargetScore) if not pd.isna(r.KRATargetScore) else 100.0)
              for _, r in role_targets.iterrows()}
    daily["KRATargetScore"] = daily["Role"].map(lambda x: rt_map.get(str(x), 100.0))
    daily["Hit_Target_Today"] = (daily["KRA_Score"] >= daily["KRATargetScore"]).astype(int)
    daily["_IsWD"] = daily.get("IsWorkingDay", False).fillna(False).astype(bool)

    daily = daily.sort_values(["EmployeeEmail","ReportDate"]).reset_index(drop=True)
    def streak_series(df_emp: pd.DataFrame) -> pd.Series:
        s = []
        cur = 0
        for _, r in df_emp.iterrows():
            if r["_IsWD"]:
                if r["Hit_Target_Today"]: cur += 1
                else: cur = 0
            s.append(cur)
        return pd.Series(s, index=df_emp.index)
    daily["Daily_Streak"] = daily.groupby("EmployeeEmail", group_keys=False).apply(streak_series)

    badge_cols = []
    for code in measure_codes:
        rank_col = f"Rank_{code}"
        if rank_col in daily.columns and badge_ok.get(code, True):
            human = label_map.get(code, code)
            bcol = f"Badge_{code}"
            daily[bcol] = np.where((daily[rank_col] == 1) & (daily[code] > 0), f"Top {human}", "")
            badge_cols.append(bcol)

    daily["Badge_KRA"] = np.where(daily["KRA_Rank_Overall"] == 1, "Top KRA", "")
    badge_cols.append("Badge_KRA")

    def collect_badges(row):
        labels = [str(row[c]) for c in badge_cols if c in row and str(row[c])]
        return ", ".join(labels)

    daily["Badges_Today"] = daily.apply(collect_badges, axis=1)
    daily.drop(columns=[c for c in badge_cols if c in daily.columns], inplace=True)

    id_cols = ["ReportDate","EmployeeEmail","EmployeeName","Role","Region","SubRegion","Manager","Include"]
    kpi_cols = ["KPI_TotalActions","KPI_UniqueApplications","KPI_UniqueStudents","KPI_UniquePartners","KPI_UniqueUniversities"]
    kra_cols = ["KRA_Score","KRA_Rank_Role","KRA_Rank_Overall","Hit_Target_Today","Daily_Streak"]
    pct_cols = [f"{m}_PctOfDay" for m in measure_codes]
    rank_cols = [c for c in daily.columns if c.startswith("Rank_") or c.startswith("IsTop3_")]
    mtd_cols  = [c for c in daily.columns if c.startswith("MTD_") or c.startswith("Expected_MTD_")]
    uniq_cols = [c for c in daily.columns if c.startswith("Ratio_") or c.startswith("IsHighUniq_")]
    extra = ["Badges_Today", "KRATargetScore", "_IsWD", "IsWorkingDay", "WorkingDaysInMonth", "WorkingDaysElapsed", "SaturdayOffPattern"]
    ordered = id_cols + kpi_cols + measure_codes + pct_cols + kra_cols + rank_cols + mtd_cols + uniq_cols + extra
    others = [c for c in daily.columns if c not in ordered]
    daily = daily.reindex(columns=ordered + others)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as w:
        daily.to_excel(w, sheet_name="summary_daily_all", index=False)

    # NEW: Write to Database for performance/reliability
    try:
        engine = get_local_db_engine()
        daily_db = daily.copy()
        daily_db["ReportDate"] = daily_db["ReportDate"].astype(str)
        daily_db.to_sql("daily_summary", engine, if_exists="replace", index=False)
        print(f"[DB] Successfully synced daily_summary table ({len(daily_db)} rows).")
    except Exception as e:
        print(f"[DB] Failed to write to daily_summary table: {e}")

    return daily, measure_codes, label_map

# ---------------------- IST helpers ---------------------- #
IST = dt.timezone(dt.timedelta(hours=5, minutes=30))

def _today_ist() -> dt.date:
    return (dt.datetime.now(dt.timezone.utc) + dt.timedelta(hours=5, minutes=30)).date()

def parse_report_date(s: str|dt.date) -> dt.date:
    if isinstance(s, dt.date): return s
    s = str(s).strip().lower()
    if s == "yesterday":
        return _today_ist() - dt.timedelta(days=1)
    return dt.date.fromisoformat(s)

def last_working_day_upto(daily_all: pd.DataFrame, upto_date_ist: dt.date) -> dt.date | None:
    sub = daily_all.loc[(daily_all["ReportDate"] < upto_date_ist) & (daily_all["_IsWD"] == True), "ReportDate"]
    return max(sub) if not sub.empty else None

# ---------------------- Announcement Builder ---------------------- #
def _topn(df: pd.DataFrame, col: str, n=3):
    if col not in df.columns: return []
    sub = df[df[col] > 0].copy()
    if sub.empty: return []
    sub["_rank"] = sub[col].rank(method="dense", ascending=False)
    sub = sub[sub["_rank"] <= n].sort_values(["_rank", col, "EmployeeName"], ascending=[True, False, True])
    return [(int(r._rank), r.EmployeeName or r.EmployeeEmail, int(r[col])) for _, r in sub.iterrows()]

def _topn_role(df: pd.DataFrame, col: str, n=3):
    res = {}
    if col not in df.columns: return res
    for role, part in df.groupby("Role"):
        res[role] = _topn(part, col, n)
    return res

def _format_topn_list(items, label):
    if not items: return f"- No {label} today."
    return "\n".join([f"- **#{rk}** {name} — {val}" for rk, name, val in items])

def _format_topn_per_role(dct, label):
    if not dct: return f"- No {label} today."
    lines = []
    for role, items in dct.items():
        if not items: continue
        lines.append(f"**{role}**")
        lines.append(_format_topn_list(items, label))
    return "\n".join(lines) if lines else f"- No {label} today."

def _streak_milestones(df_day: pd.DataFrame, milestones=STREAK_MILESTONES):
    out = []
    for _, r in df_day.iterrows():
        d = int(r.get("Daily_Streak", 0))
        if d in milestones:
            name = r["EmployeeName"] or r["EmployeeEmail"]
            out.append((name, d))
    return out

def build_announcement_text(daily_all: pd.DataFrame, measures: list[str], report_date: dt.date, label_map: dict[str, str]) -> str:
    day = report_date
    df_day = daily_all[daily_all["ReportDate"] == day].copy()
    if df_day.empty:
        return f"### KC Daily — {day:%d %b %Y}\nNo activity recorded today."

    title = f"### KC Daily — {day:%d %b %Y} (Gamified Leaderboard)"
    lines = [title, ""]

    top_kra = _topn(df_day, "KRA_Score", 3)
    lines += ["**Overall — Top 3 by KRA**", _format_topn_list(top_kra, "KRA"), ""]

    top_kra_role = _topn_role(df_day, "KRA_Score", 3)
    lines += ["**By Role — Top 3 by KRA**", _format_topn_per_role(top_kra_role, "KRA"), ""]

    lines.append("**Measure Highlights (Top 3)**")
    any_meas = False
    for code in measures:
        if code not in df_day.columns: continue
        top_m = _topn(df_day, code, 3)
        if top_m:
            any_meas = True
            human = label_map.get(code, code)
            lines += [f"- *{human}*", _format_topn_list(top_m, human)]
    if not any_meas: lines.append("- No measure activity today.")
    lines.append("")

    badges = df_day[df_day["Badges_Today"].astype(str).str.len() > 0][["EmployeeName","EmployeeEmail","Badges_Today"]]
    lines.append("**Badges Today**")
    if badges.empty: lines.append("- No badges.")
    else:
        for _, r in badges.iterrows():
            nm = r["EmployeeName"] or r["EmployeeEmail"]
            lines.append(f"- **{nm}** — {r['Badges_Today']}")
    lines.append("")

    lines.append("**Streak Milestones**")
    ms = _streak_milestones(df_day)
    if not ms: lines.append("- No streak milestones today.")
    else:
        for nm, d in ms: lines.append(f"- **{nm}** reached **{d}**-day streak 🎯")
    lines.append("")

    lines.append("**MTD Pulse (rolling)**")
    core_candidates = [c for c in measures if c.lower().startswith(("submit","assess","stg","cmnt","ofu","pend"))]
    show_measures = core_candidates[:3] if core_candidates else measures[:3]
    if not show_measures: lines.append("- (No measures configured)")
    else:
        for m in show_measures:
            col = f"MTD_{m}"
            if col in daily_all.columns:
                human = label_map.get(m, m)
                mtdd = daily_all[daily_all["ReportDate"] == day][["EmployeeName","EmployeeEmail",col]].sort_values(col, ascending=False)
                top = mtdd.head(3)
                if not top.empty and float(top[col].max()) > 0:
                    lines.append(f"- **{human}** MTD leaders:")
                    for _, r in top.iterrows():
                        nm = r["EmployeeName"] or r["EmployeeEmail"]
                        lines.append(f"  - {nm} — {int(r[col])}")
    lines.append("")
    return "\n".join(lines)

def post_announcement(text: str, webhook_url: str) -> tuple[bool, str]:
    if not webhook_url: return False, "No webhook configured."
    try: import requests
    except Exception: return False, "The 'requests' package is not installed."
    headers = {"Content-Type": "application/json"}
    payload = {"text": text}
    try:
        resp = requests.post(webhook_url, headers=headers, data=json.dumps(payload), timeout=10)
        return 200 <= resp.status_code < 300, f"HTTP {resp.status_code}"
    except Exception as e: return False, f"POST failed: {e}"

def determine_announcement_date(daily_all: pd.DataFrame) -> dt.date:
    available = [d for d in daily_all["ReportDate"].dropna().tolist()]
    if not available: raise ValueError("No ReportDate values found in summary_daily_all.")
    if REPORT_DATE_OVERRIDE:
        forced = parse_report_date(REPORT_DATE_OVERRIDE)
        if forced not in set(available): raise ValueError(f"Forced date {forced} not in summary.")
        return forced
    announce_cfg = str(ANNOUNCE_DATE).strip().lower()
    if announce_cfg == "yesterday":
        rep_date = last_working_day_upto(daily_all, _today_ist())
        return rep_date if rep_date is not None else max(available)
    return parse_report_date(ANNOUNCE_DATE)

def main() -> dt.date:
    print("[RUN] Building all-days summary (single-sheet) ...")
    daily_df, measure_codes, label_map = build_all_days(
        SETTINGS_PATH, OUT_ALL_PATH,
        ignore_cap=IGNORE_CAP, uniq_thr=UNIQUENESS_HIGH_THRESH
    )
    print(f"[DONE] Rows written: {len(daily_df)} -> {OUT_ALL_PATH}")
    rep_date = determine_announcement_date(daily_df)
    txt = build_announcement_text(daily_df, measure_codes, rep_date, label_map)
    print(f"[ANNOUNCE] Markdown ready for {rep_date.isoformat()}")
    md_path = ANNOUNCE_DIR / f"daily_announcement_{rep_date.isoformat()}.md"
    md_path.write_text(txt, encoding="utf-8")
    print(f"[ANNOUNCE] Markdown saved -> {md_path}")
    if POST_ANNOUNCEMENT:
        ok, info = post_announcement(txt, ANNOUNCE_WEBHOOK_URL)
        print(f"[ANNOUNCE] Webhook post: {'OK' if ok else 'FAILED'} ({info})")
    return rep_date

if __name__ == "__main__":
    main()
