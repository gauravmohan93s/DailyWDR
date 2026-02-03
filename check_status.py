import pandas as pd
from sqlalchemy import create_engine, text
from reporting_config import LOCAL_DB_PATH, SENT_LOG_PATH
import datetime as dt

def verify_status():
    print(f"Checking Database at: {LOCAL_DB_PATH}")
    if not LOCAL_DB_PATH.exists():
        print("[ERROR] Database file not found!")
        return

    try:
        engine = create_engine(f"sqlite:///{LOCAL_DB_PATH}")
        with engine.connect() as conn:
            # Check date range
            query = text("SELECT MIN(ActionDate), MAX(ActionDate), COUNT(*) FROM actions")
            result = conn.execute(query).fetchone()
            print(f"Database Range: {result[0]} to {result[1]}")
            print(f"Total Rows: {result[2]}")

            # Check for recent data (last 7 days)
            today = dt.date.today()
            seven_days_ago = today - dt.timedelta(days=7)
            query_recent = text(f"SELECT COUNT(*) FROM actions WHERE ActionDate >= '{seven_days_ago}'")
            recent_count = conn.execute(query_recent).scalar()
            print(f"Rows in last 7 days: {recent_count}")

    except Exception as e:
        print(f"[ERROR] Database check failed: {e}")

    print("\nChecking Sent Log...")
    if not SENT_LOG_PATH.exists():
        print("No sent_log.csv found. All dates might be considered pending.")
    else:
        try:
            df = pd.read_csv(SENT_LOG_PATH)
            print("Recent Sent Logs:")
            print(df.tail(10))
            
            # Simple check for pending (conceptually)
            # This is just a quick look; the actual logic is in run_daily_flow.py
        except Exception as e:
            print(f"[ERROR] Reading sent log failed: {e}")

if __name__ == "__main__":
    verify_status()
