
import pandas as pd
from pathlib import Path

SETTINGS_PATH = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\setting\wd_settings.xlsx")

def read_measures():
    if not SETTINGS_PATH.exists():
        print(f"Settings file not found at {SETTINGS_PATH}")
        return

    xls = pd.read_excel(SETTINGS_PATH, sheet_name=None)
    measures = xls.get("Measures")
    if measures is not None:
        print("Measures:")
        print(measures[['MeasureCode', 'MeasureLabel', 'AggType', 'AggField', 'ActionTypeIn', 'AG_In', 'AG_NotIn']].to_string())
    else:
        print("Measures sheet not found")

if __name__ == "__main__":
    read_measures()
