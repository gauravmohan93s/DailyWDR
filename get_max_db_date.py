
import sys
from pathlib import Path
from sqlalchemy import create_engine, text
from reporting_config import LOCAL_DB_PATH
from daily_wd_data_extractor import CONFIG

def get_max_date():
    if not LOCAL_DB_PATH.exists():
        print("DB File does NOT exist.")
        return

    engine = create_engine(f"sqlite:///{LOCAL_DB_PATH}")
    
    try:
        with engine.begin() as conn:
            row = conn.execute(text(f"SELECT MAX(ActionDateIST_Date) FROM {CONFIG['ACTIONS_TABLE']}")).fetchone()
        
        if row and row[0]:
            print(f"The latest date found in the database is: {row[0]}")
        else:
            print("Could not find a max date in the database. The table might be empty or the date column is null.")

    except Exception as e:
        print(f"An error occurred while querying the database: {e}")

if __name__ == "__main__":
    get_max_date()
