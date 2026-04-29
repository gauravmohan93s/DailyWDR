# -*- coding: utf-8 -*-
"""
Pre-flight settings and data validator for DailyWD-UK.
Catches common entry errors in wd_settings.xlsx.
"""

from pathlib import Path
import pandas as pd
import sys

def validate_settings(path: Path) -> bool:
    print(f"[VALIDATOR] Checking settings: {path.name}")
    if not path.exists():
        print(f"[ERROR] Settings file not found: {path}")
        return False

    try:
        xls = pd.read_excel(path, sheet_name=None)
        
        # 1. Check Team sheet
        team = xls.get("Team")
        if team is None:
            print("[ERROR] Missing 'Team' sheet.")
            return False
            
        # Duplicate Emails
        emails = team["EmployeeEmail"].dropna().astype(str).str.strip().str.lower()
        dups = emails[emails.duplicated()].unique()
        if len(dups) > 0:
            print(f"[ERROR] Duplicate emails found in roster: {', '.join(dups)}")
            return False
            
        # Missing Managers
        missing_mg = team[team["Include"].astype(str).str.lower().isin(["yes","y","1"]) & team["Manager"].isna()]
        if not missing_mg.empty:
            print(f"[WARN] Employees missing managers: {', '.join(missing_mg['EmployeeEmail'].tolist())}")

        # 2. Check Measures
        measures = xls.get("Measures")
        if measures is None:
            print("[ERROR] Missing 'Measures' sheet.")
            return False
        
        m_codes = set(measures["MeasureCode"].dropna().astype(str).str.strip())
        
        # 3. Check RoleMeasures
        role_meas = xls.get("RoleMeasures")
        if role_meas is None:
            print("[ERROR] Missing 'RoleMeasures' sheet.")
            return False
            
        unknown = role_meas[~role_meas["MeasureCode"].astype(str).str.strip().isin(m_codes)]
        if not unknown.empty:
            bad = unknown["MeasureCode"].unique()
            print(f"[ERROR] RoleMeasures refer to unknown codes: {', '.join(bad)}")
            return False

        print("[VALIDATOR] Settings look good! ✅")
        return True

    except Exception as e:
        print(f"[ERROR] Validation failed with exception: {e}")
        return False

if __name__ == "__main__":
    # If run standalone, use default path
    from reporting_config import SETTINGS_PATH # Wait, need to check if SETTINGS_PATH is in reporting_config
    # I'll use a hardcoded path for the standalone test if needed, or just import it.
    pass
