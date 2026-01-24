# -*- coding: utf-8 -*-
"""
On-Demand Performance Reporter (v4 - Final)

Generates a high-quality, two-sheet Excel report with corrected logic.
1. 'Performance_Dashboard': KPIs, charts, and monthly summary.
2. 'KRA_Daily_Breakdown': A sanitized view including all KRA-contributing
   measures and any badges earned.
"""
import pandas as pd
import argparse
from pathlib import Path
import sys
import warnings

# --- CONFIGURATION ---
SUMMARY_FILE_PATH = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\raw\summary_all_days.xlsx")
SETTINGS_FILE_PATH = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\setting\wd_settings.xlsx")
REPORTS_BASE_DIR = Path("on_demand_reports")

# Suppress pandas FutureWarnings for cleaner output
warnings.simplefilter(action='ignore', category=FutureWarning)

def get_kra_measure_columns() -> list[str]:
    """Reads the settings file to find all measures included in KRA."""
    if not SETTINGS_FILE_PATH.exists():
        print(f"[WARN] Settings file not found at {SETTINGS_FILE_PATH}. Cannot determine KRA columns.")
        return []
    
    measures_df = pd.read_excel(SETTINGS_FILE_PATH, sheet_name='Measures')
    # Ensure 'IncludeInKRA' column exists, defaulting to 'N' if not
    if 'IncludeInKRA' not in measures_df.columns:
        return []
    
    # Filter for measures where IncludeInKRA is 'Y' or 'Yes'
    kra_measures = measures_df[measures_df['IncludeInKRA'].astype(str).str.upper().isin(['Y', 'YES'])]
    return kra_measures['MeasureCode'].tolist()

def generate_report(start_date: str, end_date: str, output_path: Path, filtered_df: pd.DataFrame, kra_columns: list[str]):
    """Generates the final, corrected multi-sheet performance report."""
    if filtered_df.empty or filtered_df['IsWorkingDay'].sum() == 0:
        employee_name = output_path.stem.replace('Performance_Review_', '')
        print(f"[INFO] No working day data for {employee_name} in this period. Skipping.")
        return

    daily_view = filtered_df.sort_values(by="ReportDate").copy()
    
    # --- 1. Create the KRA Daily Breakdown View ---
    core_cols = ['ReportDate', 'EmployeeName', 'KRA_Score', 'Hit_Target_Today', 'Badges_Today']
    sanitized_cols = core_cols + kra_columns
    sanitized_cols = [col for col in sanitized_cols if col in daily_view.columns] # Ensure all exist
    kra_daily_breakdown_view = daily_view[sanitized_cols]

    # --- 2. Create Aggregated Views for Dashboard ---
    aggregation_map = {'KRA_Score': 'mean', 'Hit_Target_Today': 'sum'}
    monthly_view = daily_view.resample('MS', on='ReportDate').agg(aggregation_map).round(2)
    monthly_view['Month'] = monthly_view.index.strftime('%B %Y')

    # --- 3. Calculate Dashboard KPIs ---
    work_days_df = daily_view[daily_view['IsWorkingDay'] == True]
    avg_kra = work_days_df['KRA_Score'].mean()
    hit_target_days = work_days_df['Hit_Target_Today'].sum()
    hit_rate = (hit_target_days / len(work_days_df)) * 100 if not work_days_df.empty else 0
    
    longest_streak = 0
    current_streak = 0
    for val in daily_view.sort_values('ReportDate')['Hit_Target_Today']:
        if val == 1:
            current_streak += 1
        else:
            longest_streak = max(longest_streak, current_streak)
            current_streak = 0
    longest_streak = max(longest_streak, current_streak)

    # --- 4. Write to Excel ---
    print(f"Writing corrected report to {output_path}...")
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        kra_daily_breakdown_view.to_excel(writer, sheet_name='KRA_Daily_Breakdown', index=False, startrow=1)
        
        workbook = writer.book
        dashboard_ws = workbook.add_worksheet('Performance_Dashboard')
        data_ws = writer.sheets['KRA_Daily_Breakdown']
        data_ws.freeze_panes(2, 0)

        header_format = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter'})
        subheader_format = workbook.add_format({'bold': True, 'font_size': 14})
        kpi_title_format = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'center'})
        kpi_value_format = workbook.add_format({'font_size': 18, 'align': 'center', 'bold': True})
        percent_format = workbook.add_format({'font_size': 18, 'align': 'center', 'bold': True, 'num_format': '0.0"%"'})

        dashboard_ws.merge_range('B2:G3', 'Performance Report', header_format)
        info = daily_view.iloc[0].fillna('')
        dashboard_ws.write('B5', 'Employee:', subheader_format)
        dashboard_ws.write('C5', f"{info['EmployeeName']} ({info['EmployeeEmail']})")
        dashboard_ws.write('B6', 'Date Range:', subheader_format)
        dashboard_ws.write('C6', f"{start_date} to {end_date}")

        dashboard_ws.merge_range('B10:C10', 'Average KRA Score', kpi_title_format)
        dashboard_ws.merge_range('B11:C12', f"{avg_kra:.1f}" if pd.notna(avg_kra) else "N/A", kpi_value_format)

        dashboard_ws.merge_range('D10:E10', 'Target Achievement', kpi_title_format)
        dashboard_ws.merge_range('D11:E12', hit_rate, percent_format)
        
        dashboard_ws.merge_range('F10:G10', 'Longest Target Streak', kpi_title_format)
        dashboard_ws.merge_range('F11:G12', f"{longest_streak} Days", kpi_value_format)

        dashboard_ws.write('B15', 'Monthly Summary', subheader_format)
        monthly_view.to_excel(writer, sheet_name='Performance_Dashboard', startrow=16, startcol=1, index=False)

        num_data_points = len(kra_daily_breakdown_view)
        if num_data_points > 1:
            date_col_idx = kra_daily_breakdown_view.columns.get_loc('ReportDate') + 1
            kra_col_idx = kra_daily_breakdown_view.columns.get_loc('KRA_Score') + 1

            kra_chart = workbook.add_chart({'type': 'line'})
            kra_chart.set_title({'name': 'Daily KRA Score Trend'})
            kra_chart.add_series({
                'name': 'KRA Score',
                'categories': ['KRA_Daily_Breakdown', 2, date_col_idx - 1, num_data_points + 1, date_col_idx - 1],
                'values':     ['KRA_Daily_Breakdown', 2, kra_col_idx - 1, num_data_points + 1, kra_col_idx - 1],
            })
            kra_chart.set_x_axis({'date_axis': True, 'num_font': {'rotation': -45}})
            kra_chart.set_legend({'position': 'none'})
            dashboard_ws.insert_chart('I2', kra_chart, {'x_scale': 2, 'y_scale': 1.5})

    print(f"[SUCCESS] Finished report for {info['EmployeeName']}")

def main():
    parser = argparse.ArgumentParser(description="Generate final, corrected performance reports.")
    parser.add_argument("--start-date", required=True, help="Start date (YYYY-MM-DD).")
    parser.add_argument("--end-date", required=True, help="End date (YYYY-MM-DD).")
    parser.add_argument("--all-members", action='store_true', help="Run for all members.", required=True)
    
    args = parser.parse_args()

    if not SUMMARY_FILE_PATH.exists():
        print(f"[ERROR] Summary file not found: {SUMMARY_FILE_PATH}")
        sys.exit(1)
        
    print("Loading base data and KRA settings...")
    df = pd.read_excel(SUMMARY_FILE_PATH, sheet_name='summary_daily_all')
    df['ReportDate'] = pd.to_datetime(df['ReportDate'])
    kra_columns = get_kra_measure_columns()
    
    mask = (df['ReportDate'] >= args.start_date) & (df['ReportDate'] <= args.end_date)
    date_filtered_df = df.loc[mask]

    members_to_process = date_filtered_df['EmployeeEmail'].str.lower().dropna().unique().tolist()
        
    if not members_to_process:
        print("[WARN] No employees found for the specified period.")
        return

    run_folder = f"{args.start_date}_to_{args.end_date}"
    output_dir = REPORTS_BASE_DIR / run_folder
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Reports will be saved in: {output_dir}")

    total_members = len(members_to_process)
    for i, email in enumerate(members_to_process):
        member_df = date_filtered_df[date_filtered_df['EmployeeEmail'].str.lower() == email].copy()
        
        if member_df.empty:
            continue
            
        employee_name = member_df.iloc[0]['EmployeeName'].replace(' ', '_').replace('.', '')
        filename = f"Performance_Review_{employee_name}.xlsx"
        output_path = output_dir / filename
        
        print(f"\n({i+1}/{total_members}) Generating report for {email}...")
        generate_report(args.start_date, args.end_date, output_path, member_df, kra_columns)

if __name__ == "__main__":
    main()