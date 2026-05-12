# -*- coding: utf-8 -*-
"""
One-time migration script to move settings from Excel to SQLite.
Ensures the target directory and database exist.
"""

import pandas as pd
import sqlite3
import os
from pathlib import Path
from reporting_config import SETTINGS_PATH, LOCAL_DB_PATH

def migrate():
    print(f"[MIGRATION] Checking: {SETTINGS_PATH}")
    if not SETTINGS_PATH.exists():
        print(f"[ERROR] Settings file not found: {SETTINGS_PATH}")
        return

    # Ensure the directory for the database exists
    try:
        LOCAL_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    except:
        pass
    
    print(f"[MIGRATION] Opening database: {LOCAL_DB_PATH}")
    conn = sqlite3.connect(LOCAL_DB_PATH)
    
    try:
        xls = pd.read_excel(SETTINGS_PATH, sheet_name=None)
        
        # 1. Migrate Team Roster
        team = xls.get("Team")
        if team is not None:
            team["EmployeeEmail"] = team["EmployeeEmail"].astype(str).str.strip().str.lower()
            team = team.drop_duplicates(subset=["EmployeeEmail"], keep="first")
            team.to_sql("roster", conn, if_exists="replace", index=False)
            try:
                conn.execute("CREATE UNIQUE INDEX idx_roster_email ON roster (EmployeeEmail)")
            except: pass
            print(f"[MIGRATION] Rostered {len(team)} members to DB.")

        # 2. Migrate Measures
        measures = xls.get("Measures")
        if measures is not None:
            measures.to_sql("measures_config", conn, if_exists="replace", index=False)
            print(f"[MIGRATION] Migrated {len(measures)} measure definitions.")

        # 3. Migrate RoleMeasures
        role_meas = xls.get("RoleMeasures")
        if role_meas is not None:
            role_meas.to_sql("role_measures", conn, if_exists="replace", index=False)
            print(f"[MIGRATION] Migrated {len(role_meas)} role-measure mappings.")
            
        conn.commit()
        print("[MIGRATION] Complete! ✅")
    except Exception as e:
        print(f"[ERROR] Migration failed: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    migrate()
