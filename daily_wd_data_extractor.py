# -*- coding: utf-8 -*-
"""
KC – Data Extractor (v2.0 - Database Centric)
------------------------------------------------------------------------
- Extracts raw actions and stores them in a local SQLite database (`kc_reports.db`).
- Performs an initial one-time backfill from existing daily CSVs.
- Daily runs fetch only the latest data from the source DB and upsert it into SQLite.
- Applies 'Action Group' mapping directly during the database load.
- Idempotency is handled by checking the last fetched date in the database.
"""

import warnings
import datetime as dt
import re
from pathlib import Path
from typing import List, Tuple, Dict

import pandas as pd
from sqlalchemy import create_engine, text, inspect
from sqlalchemy.types import String
from reporting_config import LOCAL_DB_PATH

warnings.filterwarnings("ignore", category=FutureWarning, module="pandas")

# ------------------------ CONFIG ------------------------
CONFIG = {
    # --- Source DB (MSSQL) ---
    "DB_SERVER": "sql-web-edition.cqu5cuepe15y.ap-south-1.rds.amazonaws.com",
    "DB_NAME":   "Kc_WebAppDb_New_Migrated",
    "ODBC_DRIVER": "ODBC Driver 17 for SQL Server",
    "DB_UID": "gaurav_sakhare_prod",
    "DB_PWD": "{G>865yQhe5NP",
    "COUNTRY_ID": 4,
    "USE_REMOTE_DB": True,  # set True to actively fetch recent data from source DB for daily runs

    # --- Local DB (SQLite) ---
    "SQLITE_DB_PATH": LOCAL_DB_PATH,
    "ACTIONS_TABLE": "actions",

    # --- Settings/Team list ---
    "SETTINGS_DIR": r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\setting",
    "TEAM_XLSX": "team_members.xlsx",
    "SETTINGS_FILE": "settings.xlsx",
    "SETTINGS_SHEET": None,

    # --- Raw CSV directory (for one-time backfill) ---
    "RAW_DIR": r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\raw\by_day",

    # --- Safety limits ---
    "MAX_EMAILS_PER_SQL": 800,
    "MAX_DAYS_PER_RUN": 40, # Max recent days to import per run (set high to allow full month backfill on first import)
}

IST = dt.timezone(dt.timedelta(hours=5, minutes=30))
UTC = dt.timezone.utc
BASE_SQL = """
WITH NewApps AS (
    SELECT
        ack.AcknowledgementNumber, anu.Email AS EmployeeEmail, avs.SubmittedDate AS ActionDate,
        CAST('NewApplication' AS NVARCHAR(100)) AS ActionType, CAST('NewApplication' AS NVARCHAR(100)) AS SubType1,
        avs.UniversityId, avs.University, avs.PartnerName, avs.Assignee, avs.TagGroupId, avs.Region, avs.RegionalManager, avs.StudentId, avs.Intake, avs.InYear
    FROM  dbo.Acknowledgements AS ack
    JOIN  dbo.AspNetUsers AS anu ON ack.StaffId = anu.Id
    JOIN  dbo.Application_View_Staff AS avs ON ack.AcknowledgementNumber = avs.AcknowledgementNumber
    WHERE avs.CountryId = :country_id AND avs.SubmittedDate >= :start_utc AND avs.SubmittedDate < :end_utc {TEAM_FILTER_NEW}
), StageChanges AS (
    SELECT
        avs.AcknowledgementNumber, anu.Email AS EmployeeEmail, ah.ChangedOn AS ActionDate,
        CAST('StageChange' AS NVARCHAR(100)) AS ActionType, CAST(ast.Stage AS NVARCHAR(100)) AS SubType1,
        avs.UniversityId, avs.University, avs.PartnerName, avs.Assignee, avs.TagGroupId, avs.Region, avs.RegionalManager, avs.StudentId, avs.Intake, avs.InYear
    FROM  dbo.ApplicationHistory AS ah
    JOIN  dbo.Application_View_Staff AS avs ON ah.AcknowledgementId = avs.AcknowledgementId
    JOIN  dbo.AspNetUsers AS anu ON ah.ChangedBy = anu.Id
    JOIN  dbo.ApplicationStages AS ast ON ah.ApplicationStageId = ast.Id
    WHERE avs.CountryId = :country_id AND ah.ChangedOn >= :start_utc AND ah.ChangedOn < :end_utc {TEAM_FILTER_STAGE}
), OfferFollowUp AS (
    SELECT
        ofu.AcknowledgementNumber, anu.Email AS EmployeeEmail, ofu.CreatedAt AS ActionDate,
        CAST(ofu.FollowupType AS NVARCHAR(100)) AS ActionType, CAST(ofu.InteractionType AS NVARCHAR(100)) AS SubType1,
        avs.UniversityId, avs.University, avs.PartnerName, avs.Assignee, avs.TagGroupId, avs.Region, avs.RegionalManager, ofu.StudentId, avs.Intake, avs.InYear
    FROM  dbo.StudentInteractions AS ofu
    JOIN  dbo.AspNetUsers AS anu ON ofu.CreatedBy = anu.Id
    JOIN  dbo.Application_View_Staff AS avs ON ofu.AcknowledgementNumber = avs.AcknowledgementNumber
    WHERE avs.CountryId = :country_id AND ofu.CreatedAt >= :start_utc AND ofu.CreatedAt < :end_utc {TEAM_FILTER_OFU}
), Comments AS (
    SELECT
        ack.AcknowledgementNumber, u.Email AS EmployeeEmail, sl.CreatedOn AS ActionDate,
        CAST('Comment' AS NVARCHAR(100)) AS ActionType, CAST('Comment' AS NVARCHAR(100)) AS SubType1,
        avs.UniversityId, avs.University, avs.PartnerName, avs.Assignee, avs.TagGroupId, avs.Region, avs.RegionalManager, avs.StudentId, avs.Intake, avs.InYear
    FROM  dbo.SendEmailLog AS sl
    JOIN  dbo.AspNetUsers AS u ON sl.CreatedBy = u.Id AND u.IsStaff = 1
    JOIN  dbo.Acknowledgements AS ack ON ack.AcknowledgementId = sl.AcknowledgementId
    JOIN  dbo.Application_View_Staff AS avs ON ack.AcknowledgementId = avs.AcknowledgementId
    WHERE avs.CountryId = :country_id AND sl.EmailType = 'comment' AND sl.CreatedOn >= :start_utc AND sl.CreatedOn < :end_utc {TEAM_FILTER_COMMENT}
)
SELECT * FROM NewApps UNION ALL SELECT * FROM StageChanges UNION ALL SELECT * FROM OfferFollowUp UNION ALL SELECT * FROM Comments;
"""

# ------------------------ Database Engines ------------------------
def get_source_db_engine():
    """Returns the engine for the source MSSQL database."""
    uid, pwd = CONFIG["DB_UID"], CONFIG["DB_PWD"]
    if not uid or not pwd:
        raise RuntimeError("Source DB credentials missing.")
    from urllib.parse import quote_plus
    conn_str = f"mssql+pyodbc://{quote_plus(uid)}:{quote_plus(pwd)}@{CONFIG['DB_SERVER']}/{CONFIG['DB_NAME']}?driver={quote_plus(CONFIG['ODBC_DRIVER'])}&Encrypt=yes&TrustServerCertificate=yes"
    return create_engine(conn_str, fast_executemany=True, pool_pre_ping=True)

def get_local_db_engine():
    """Returns the engine for the local SQLite database."""
    db_path = Path(CONFIG["SQLITE_DB_PATH"])
    db_path.parent.mkdir(parents=True, exist_ok=True)
    return create_engine(f"sqlite:///{db_path}")

# ------------------------ Helpers ------------------------
def load_team_emails() -> List[str]:
    p = Path(CONFIG["SETTINGS_DIR"]) / CONFIG["TEAM_XLSX"]
    if not p.exists(): raise FileNotFoundError(f"Team file not found: {p}")
    df = pd.read_excel(p)
    email_col = next((c for c in ["Email","OfficialEmail","EmployeeEmail","Email ID (Official)"] if c in df.columns), None)
    if not email_col: raise ValueError("Team file must contain an email column.")
    return df[email_col].astype(str).str.strip().str.lower().dropna().drop_duplicates().tolist()

def load_action_group_mapping() -> pd.DataFrame:
    map_path = Path(CONFIG["SETTINGS_DIR"]) / CONFIG["SETTINGS_FILE"]
    if not map_path.exists(): raise FileNotFoundError(f"Mapping file not found: {map_path}")
    
    sheet = CONFIG["SETTINGS_SHEET"]
    xls = pd.read_excel(map_path, sheet_name=None)
    if sheet is None:
        for s_name, df_sheet in xls.items():
            cols = {c.lower().strip() for c in df_sheet.columns}
            if {"action type", "action taken", "action group"}.issubset(cols):
                sheet = s_name
                break
        if sheet is None: sheet = list(xls.keys())[0] # Fallback
    
    map_df = xls[sheet]
    col_map = {c.lower().strip(): c for c in map_df.columns}
    map_df = map_df.rename(columns={
        col_map.get("action type", "Action Type"): "Action Type",
        col_map.get("action taken", "Action Taken"): "Action Taken",
        col_map.get("action group", "Action Group"): "Action Group"
    })

    if not {"Action Type", "Action Taken", "Action Group"}.issubset(map_df.columns):
        raise ValueError("Mapping sheet is missing required columns.")

    map_df["MapKey"] = (map_df["Action Type"].astype(str).str.strip() + "-" +
                        map_df["Action Taken"].astype(str).str.strip())
    return map_df[["MapKey", "Action Group"]].dropna(subset=["MapKey"]).drop_duplicates()

def apply_mapping_and_normalize(df: pd.DataFrame, mapping_df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df

    # Normalize emails and create MapKey for joining
    df["EmployeeEmail"] = df["EmployeeEmail"].astype(str).str.strip().str.lower()

    # Normalize datetimes to UTC text for SQLite (prefer ActionDateIST if present)
    if "ActionDateIST" in df.columns:
        adt = pd.to_datetime(df["ActionDateIST"], errors="coerce")
        try:
            if adt.dt.tz is None:
                adt = adt.dt.tz_localize(IST)
        except Exception:
            adt = adt.dt.tz_localize(IST)
        adt_utc = adt.dt.tz_convert(UTC)
        df["ActionDate"] = adt_utc.dt.strftime("%Y-%m-%d %H:%M:%S")
    else:
        df["ActionDate"] = pd.to_datetime(df["ActionDate"], errors="coerce", utc=True).dt.strftime("%Y-%m-%d %H:%M:%S")

    # Ensure ActionDateIST / ActionDateIST_Date exist for consistency
    if "ActionDateIST" not in df.columns or "ActionDateIST_Date" not in df.columns:
        adt_utc = pd.to_datetime(df["ActionDate"], errors="coerce", utc=True)
        adt_ist = adt_utc.dt.tz_convert(IST)
        if "ActionDateIST" not in df.columns:
            df["ActionDateIST"] = adt_ist.dt.strftime("%Y-%m-%d %H:%M:%S")
        if "ActionDateIST_Date" not in df.columns:
            df["ActionDateIST_Date"] = adt_ist.dt.date
    df["MapKey"] = (df["ActionType"].fillna("").astype(str).str.strip() + "-" +
                    df["SubType1"].fillna("").astype(str).str.strip())
    
    # Join mapping
    df = df.merge(mapping_df, on="MapKey", how="left")

    # Define a robust unique key for each action
    uk_cols = ["EmployeeEmail", "ActionType", "SubType1", "AcknowledgementNumber", "StudentId", "ActionDate"]
    df["unique_id"] = df[uk_cols].astype(str).agg("|".join, axis=1).str.encode('utf-8').apply(lambda x: __import__('hashlib').sha256(x).hexdigest())
    
    return df


def _get_table_columns(engine, table_name: str) -> list[str]:
    with engine.begin() as conn:
        rows = conn.execute(text(f"PRAGMA table_info({table_name})")).fetchall()
    return [r[1] for r in rows]


def _upsert_dataframe(local_eng, df_processed: pd.DataFrame):
    if df_processed.empty:
        print("[DB] No rows to upsert.")
        return

    table = CONFIG["ACTIONS_TABLE"]
    cols = _get_table_columns(local_eng, table)
    if not cols:
        raise RuntimeError("Target table has no columns.")

    for c in cols:
        if c not in df_processed.columns:
            df_processed[c] = None
    df_use = df_processed[cols].copy()

    temp_table = f"temp_{table}"
    df_use.to_sql(name=temp_table, con=local_eng, if_exists="replace", index=False)

    col_list = ", ".join([f'"{c}"' for c in cols])
    with local_eng.begin() as conn:
        conn.execute(text(f"""
            DELETE FROM {table}
            WHERE unique_id IN (SELECT unique_id FROM {temp_table});
        """))
        conn.execute(text(f"""
            INSERT INTO {table} ({col_list})
            SELECT {col_list} FROM {temp_table};
        """))
        conn.execute(text(f"DROP TABLE {temp_table};"))


def _parse_raw_date_from_name(name: str) -> dt.date | None:
    m = re.search(r"raw_(\d{4}-\d{2}-\d{2})\.csv", name)
    if not m:
        return None
    try:
        return dt.date.fromisoformat(m.group(1))
    except Exception:
        return None


def _iter_raw_csv_files(raw_dir: Path) -> list[tuple[dt.date, Path]]:
    files = []
    for p in sorted(raw_dir.glob("raw_*.csv")):
        d = _parse_raw_date_from_name(p.name)
        if d:
            files.append((d, p))
    return files


def _get_db_date_range(local_eng) -> tuple[dt.date | None, dt.date | None]:
    with local_eng.begin() as conn:
        try:
            row = conn.execute(text(f"SELECT MIN(ActionDateIST_Date), MAX(ActionDateIST_Date) FROM {CONFIG['ACTIONS_TABLE']}")).fetchone()
        except Exception:
            row = None
    if not row:
        return None, None
    try:
        min_d = dt.date.fromisoformat(row[0]) if row[0] else None
    except Exception:
        min_d = None
    try:
        max_d = dt.date.fromisoformat(row[1]) if row[1] else None
    except Exception:
        max_d = None
    return min_d, max_d

# ------------------------ Database Operations ------------------------
def initialize_database():
    """(One-Time) Create database and backfill from existing CSVs."""
    local_eng = get_local_db_engine()
    inspector = inspect(local_eng)

    if inspector.has_table(CONFIG["ACTIONS_TABLE"]):
        print("[DB] 'actions' table already exists. Skipping initialization.")
        return

    print("[DB] Initializing new SQLite database and backfilling from CSVs...")
    local_eng = get_local_db_engine()
    
    # Find and process all raw CSVs
    raw_dir = Path(CONFIG["RAW_DIR"])
    files = sorted(list(raw_dir.glob("raw_*.csv")))
    if not files:
        print("[DB] No raw CSVs found to backfill. Database will be empty.")
        # Create an empty table with the correct schema
        pd.DataFrame(columns=['unique_id']).to_sql(CONFIG['ACTIONS_TABLE'], local_eng, if_exists='replace', index=False, dtype={'unique_id': 'TEXT PRIMARY KEY'})
        return

    print(f"[DB] Found {len(files)} daily CSVs to backfill...")
    mapping_df = load_action_group_mapping()
    
    all_dfs = []
    for i, p in enumerate(files, 1):
        print(f"  - Processing {p.name} ({i}/{len(files)})")
        df = pd.read_csv(p, dtype=str)
        df_processed = apply_mapping_and_normalize(df, mapping_df)
        all_dfs.append(df_processed)

    if not all_dfs:
        print("[DB] No data in CSVs. Database will be empty.")
        return
        
    full_df = pd.concat(all_dfs, ignore_index=True).drop_duplicates(subset=["unique_id"])
    
    print(f"[DB] Writing {len(full_df)} total unique records to database...")
    # Set unique_id as index to create it as a primary key
    full_df = full_df.set_index('unique_id')
    
    full_df.to_sql(
        CONFIG["ACTIONS_TABLE"],
        local_eng,
        if_exists="replace",
        index=True,
        index_label='unique_id',
        dtype={'unique_id': String(64)},
        chunksize=10000,
        method='multi'
    )
    print("[DB] Database initialization complete.")

def fetch_from_source_db(start_utc: dt.datetime, end_utc: dt.datetime, team_emails: List[str] | None = None) -> pd.DataFrame:
    """
    Fetch data directly from the source MSSQL database for a given UTC date range.
    Returns a DataFrame with the fetched records (before normalization).
    """
    source_eng = get_source_db_engine()
    team_emails = team_emails or load_team_emails()
    
    # Build team email filters for each section of the query
    if team_emails:
        email_list = [f"'{e}'" for e in team_emails]
        team_filter = f"AND anu.Email IN ({','.join(email_list)})" if email_list else ""
        team_filter_stage = f"AND anu.Email IN ({','.join(email_list)})" if email_list else ""
        team_filter_ofu = f"AND anu.Email IN ({','.join(email_list)})" if email_list else ""
        team_filter_comment = f"AND u.Email IN ({','.join(email_list)})" if email_list else ""
    else:
        team_filter = team_filter_stage = team_filter_ofu = team_filter_comment = ""
    
    # Replace placeholders
    query = BASE_SQL.format(
        TEAM_FILTER_NEW=team_filter,
        TEAM_FILTER_STAGE=team_filter_stage,
        TEAM_FILTER_OFU=team_filter_ofu,
        TEAM_FILTER_COMMENT=team_filter_comment
    )
    
    params = {
        "country_id": CONFIG["COUNTRY_ID"],
        "start_utc": start_utc,
        "end_utc": end_utc
    }
    
    try:
        # Use text() to properly handle parameter binding with pyodbc
        df = pd.read_sql(text(query), source_eng, params=params, dtype=str)
        return df
    except Exception as e:
        print(f"[WARN] Failed to fetch from source DB: {e}")
        return pd.DataFrame()

def fetch_and_upsert_recent_data():
    """Fetch recent data and UPSERT into local SQLite DB."""
    print("[DB] Updating local SQLite...")
    local_eng = get_local_db_engine()

    # Ensure table exists
    if not inspect(local_eng).has_table(CONFIG['ACTIONS_TABLE']):
        print("[DB] Actions table not found. Please run initialization first.")
        return

    min_db, max_db = _get_db_date_range(local_eng)
    today_ist = dt.datetime.now(IST).date()
    yesterday_ist = today_ist - dt.timedelta(days=1)
    
    # Try to fetch yesterday's data from source DB first
    mapping_df = load_action_group_mapping()
    total = 0
    
    if CONFIG.get("USE_REMOTE_DB", False):
        # Fetch from source MSSQL for yesterday (IST) → UTC range
        yesterday_start_ist = dt.datetime.combine(yesterday_ist, dt.time.min).replace(tzinfo=IST)
        yesterday_end_ist = dt.datetime.combine(yesterday_ist, dt.time.max).replace(tzinfo=IST)
        
        yesterday_start_utc = yesterday_start_ist.astimezone(UTC)
        yesterday_end_utc = yesterday_end_ist.astimezone(UTC)
        
        print(f"[DB] Fetching from source DB for {yesterday_ist.isoformat()}...")
        df_source = fetch_from_source_db(yesterday_start_utc, yesterday_end_utc)
        
        if not df_source.empty:
            # Normalize the fetched data
            df_processed = apply_mapping_and_normalize(df_source, mapping_df)
            if not df_processed.empty:
                _upsert_dataframe(local_eng, df_processed)
                total += len(df_processed)
                print(f"[DB] Upserted {len(df_processed)} records from source DB for {yesterday_ist}")
        else:
            print(f"[DB] No data found in source DB for {yesterday_ist}")
    
    # Also import from CSVs if available (as fallback/backfill)
    print("[DB] Importing from CSVs...")
    raw_dir = Path(CONFIG["RAW_DIR"])
    files = _iter_raw_csv_files(raw_dir)
    if not files:
        print("[DB] No raw CSVs found for update.")
        if total == 0:
            return
    else:
        # Determine DB date range (IST) and CSV range
        earliest_csv = files[0][0] if files else None
        latest_csv = files[-1][0] if files else None

        # Select CSVs to import:
        # 1) If DB empty -> import all CSVs (one-time full backfill)
        # 2) Always import recent files from last date - 1 day (daily updates)
        # 3) If DB is missing early dates, backfill those CSVs too (accuracy fix)
        selected = []
        if not min_db or not max_db:
            selected = files
        else:
            recent_from = max_db - dt.timedelta(days=1)
            selected = [(d, p) for (d, p) in files if d >= recent_from]
            if earliest_csv and min_db and earliest_csv < min_db:
                missing_early = [(d, p) for (d, p) in files if d < min_db]
                if missing_early:
                    print(f"[DB] Backfilling {len(missing_early)} earlier CSVs for accuracy ({earliest_csv} to {min_db - dt.timedelta(days=1)})")
                    selected = missing_early + selected

        # Safety cap for DAILY updates only (skip cap if we are doing a backfill)
        max_days = int(CONFIG.get("MAX_DAYS_PER_RUN", 0) or 0)
        if max_days > 0 and len(selected) > max_days and (not min_db or (earliest_csv and min_db and earliest_csv >= min_db)):
            selected = selected[-max_days:]

        if selected:
            for d, p in selected:
                print(f"[DB] Importing CSV: {p.name}")
                try:
                    df_raw = pd.read_csv(p, dtype=str)
                except Exception as exc:
                    print(f"[WARN] Failed to read {p.name}: {exc}")
                    continue
                if df_raw.empty:
                    continue
                df_processed = apply_mapping_and_normalize(df_raw, mapping_df)
                if df_processed.empty:
                    continue
                _upsert_dataframe(local_eng, df_processed)
                total += len(df_processed)

    if total > 0:
        print(f"[DB] Upserted {total} total records into the local database.")
    else:
        print("[DB] No new records upserted.")

# --- Helper for chunking and IN clauses ---
def _chunks(seq, n):
    for i in range(0, len(seq), n):
        yield seq[i:i+n]

def _in_clause(prefix: str, items: List[str]) -> Tuple[str, Dict[str,str]]:
    params, placeholders = {}, []
    for i, it in enumerate(items):
        key = f"{prefix}{i}"; params[key] = it; placeholders.append(f":{key}")
    return ",".join(placeholders), params

# ------------------------ MAIN ------------------------
if __name__ == "__main__":
    # The new flow: initialize once, then update daily.
    # To re-run initialization, delete the kc_reports.db file.
    initialize_database()
    fetch_and_upsert_recent_data()
