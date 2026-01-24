# -*- coding: utf-8 -*-
"""
On-Demand Performance Reporter (v2 - Designed)

This script generates a visually enhanced, multi-sheet Excel report with a
performance dashboard, KPIs, and charts for a given team member or manager.

It saves reports to a structured folder: `on_demand_reports`.
"""
import pandas as pd
import argparse
from pathlib import Path
import sys
import numpy as np

# --- CONFIGURATION ---
SUMMARY_FILE_PATH = Path(r"C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\raw\summary_all_days.xlsx")
REPORTS_BASE_DIR = Path("on_demand_reports")

def generate_report(start_date: str, end_date: str, output_path: Path, filtered_df: pd.DataFrame):
    """
    Generates the multi-sheet performance report with a designed dashboard.
    """
    if filtered_df.empty:
        print("[WARN] No data found for the specified filters. Report will not be generated.")
        return

    # --- 1. Prepare Data Views ---
    daily_view = filtered_df.sort_values(by=["ReportDate"])
    
    aggregation_map = {
        'KRA_Score': 'mean',
        'Hit_Target_Today': 'sum',
        'KPI_TotalActions': 'sum',
    }
    
    # Ensure columns exist before trying to aggregate them
    valid_agg_map = {k: v for k, v in aggregation_map.items() if k in daily_view.columns}

    # Set index for resampling
    df_for_agg = daily_view.set_index('ReportDate')

    weekly_view = df_for_agg.resample('W-MON').agg(valid_agg_map).round(2).reset_index()
    monthly_view = df_for_agg.resample('MS').agg(valid_agg_map).round(2).reset_index()
    monthly_view['Month'] = monthly_view['ReportDate'].dt.strftime('%B %Y')

    # --- 2. Calculate Dashboard KPIs ---
    avg_kra = daily_view['KRA_Score'].mean()
    total_work_days = daily_view['IsWorkingDay'].sum()
    hit_target_days = daily_view['Hit_Target_Today'].sum()
    hit_rate = (hit_target_days / total_work_days) * 100 if total_work_days > 0 else 0
    
    # Calculate longest streak
    longest_streak = 0
    current_streak = 0
    for val in daily_view['Hit_Target_Today']:
        if val == 1:
            current_streak += 1
        else:
            longest_streak = max(longest_streak, current_streak)
            current_streak = 0
    longest_streak = max(longest_streak, current_streak)

    # --- 3. Write to Excel using XlsxWriter ---
    print(f"Writing designed report to {output_path}...")
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        # Write data sheets first (can be hidden later if needed)
        daily_view.to_excel(writer, sheet_name='Detailed_Daily_Data', index=False, startrow=1)
        weekly_view.to_excel(writer, sheet_name='Weekly_Trends', index=False)
        monthly_view.to_excel(writer, sheet_name='Monthly_Summary', index=False)

        # Get workbook and worksheet objects
        workbook = writer.book
        dashboard_ws = workbook.add_worksheet('Performance_Dashboard')

        # --- 4. Design the Dashboard ---
        # Define formats
        header_format = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center'})
        subheader_format = workbook.add_format({'bold': True, 'font_size': 14})
        kpi_title_format = workbook.add_format({'bold': True, 'font_size': 11})
        kpi_value_format = workbook.add_format({'font_size': 18, 'align': 'center'})
        percent_format = workbook.add_format({'font_size': 18, 'align': 'center', 'num_format': '0.0"%"'})

        # Header
        dashboard_ws.merge_range('B2:G3', 'On-Demand Performance Report', header_format)
        info = daily_view.iloc[0].fillna('')
        dashboard_ws.write('C5', f"{info['EmployeeName']} ({info['EmployeeEmail']})")
        dashboard_ws.write('B6', 'Manager:', subheader_format)
        dashboard_ws.write('C6', info['Manager'])
        dashboard_ws.write('B7', 'Date Range:', subheader_format)
        dashboard_ws.write('C7', f"{start_date} to {end_date}")

        # KPIs
        dashboard_ws.merge_range('B10:C10', 'Average KRA Score', kpi_title_format)
        dashboard_ws.merge_range('B11:C12', f"{avg_kra:.2f}", kpi_value_format)

        dashboard_ws.merge_range('D10:E10', 'Target Hit Rate', kpi_title_format)
        dashboard_ws.merge_range('D11:E12', hit_rate, percent_format)
        
        dashboard_ws.merge_range('F10:G10', 'Longest Target Streak', kpi_title_format)
        dashboard_ws.merge_range('F11:G12', f"{longest_streak} Days", kpi_value_format)

        # --- 5. Add Charts ---
        num_data_points = len(daily_view)

        # Chart 1: KRA Score Over Time (Line Chart)
        kra_chart = workbook.add_chart({'type': 'line'})
        kra_chart.set_title({'name': 'Daily KRA Score Trend'})
        # Add data series: =Detailed_Daily_Data!$I$3:$I$[num_data_points+2]
        # Note: Finding column index is safer than hardcoding
        date_col_idx = daily_view.columns.get_loc('ReportDate') + 1
        kra_col_idx = daily_view.columns.get_loc('KRA_Score') + 1
        
        kra_chart.add_series({
            'name': 'KRA Score',
            'categories': ['Detailed_Daily_Data', 2, date_col_idx - 1, num_data_points + 1, date_col_idx - 1],
            'values':     ['Detailed_Daily_Data', 2, kra_col_idx - 1, num_data_points + 1, kra_col_idx - 1],
        })
        kra_chart.set_x_axis({'date_axis': True})
        dashboard_ws.insert_chart('B15', kra_chart, {'x_scale': 1.8, 'y_scale': 1})

        # Chart 2: Weekly KRA Average (Bar Chart)
        weekly_chart = workbook.add_chart({'type': 'column'})
        weekly_chart.set_title({'name': 'Weekly KRA Score Average'})
        
        week_date_col_idx = weekly_view.columns.get_loc('ReportDate') + 1
        week_kra_col_idx = weekly_view.columns.get_loc('KRA_Score') + 1
        num_weeks = len(weekly_view)

        weekly_chart.add_series({
            'name': 'Avg KRA/Week',
            'categories': ['Weekly_Trends', 1, week_date_col_idx - 1, num_weeks, week_date_col_idx - 1],
            'values':     ['Weekly_Trends', 1, week_kra_col_idx - 1, num_weeks, week_kra_col_idx - 1],
        })
        weekly_chart.set_x_axis({'date_axis': True})
        dashboard_ws.insert_chart('B30', weekly_chart, {'x_scale': 1.8, 'y_scale': 1})
    
    print(f"[SUCCESS] Finished report for {daily_view.iloc[0]['EmployeeName']}")

def main():
    """Main function to parse arguments and run the report generation."""
    parser = argparse.ArgumentParser(description="Generate on-demand performance reports.")
    parser.add_argument("--start-date", required=True, help="Start date (YYYY-MM-DD).")
    parser.add_argument("--end-date", required=True, help="End date (YYYY-MM-DD).")
    
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--email", help="Email of a single team member.")
    group.add_argument("--manager", help="Manager's email for their team.")
    group.add_argument("--all-members", action='store_true', help="Run for all members in the summary file.")

    parser.add_argument("--output-dir", help="Optional: Directory to save reports.")

    args = parser.parse_args()

    # --- Load and Prepare Data ---
    if not SUMMARY_FILE_PATH.exists():
        print(f"[ERROR] Summary file not found: {SUMMARY_FILE_PATH}")
        sys.exit(1)
        
    print("Loading base data...")
    df = pd.read_excel(SUMMARY_FILE_PATH)
    df['ReportDate'] = pd.to_datetime(df['ReportDate'])
    
    # Filter by date range once
    mask = (df['ReportDate'] >= args.start_date) & (df['ReportDate'] <= args.end_date)
    date_filtered_df = df.loc[mask].copy()

    # --- Determine Members to Process ---
    members_to_process = []
    if args.email:
        members_to_process.append(args.email.lower())
    elif args.manager:
        manager_mask = date_filtered_df['Manager'].str.lower() == args.manager.lower()
        members_to_process = date_filtered_df[manager_mask]['EmployeeEmail'].str.lower().unique().tolist()
    elif args.all_members:
        members_to_process = date_filtered_df['EmployeeEmail'].str.lower().unique().tolist()
    else:
        print("[ERROR] You must specify --email, --manager, or --all-members.")
        sys.exit(1)
        
    if not members_to_process:
        print("[WARN] No employees found for the specified criteria.")
        return

    # --- Setup Output Directory ---
    if args.output_dir:
        output_dir = Path(args.output_dir)
    else:
        run_folder = f"{args.start_date}_to_{args.end_date}"
        output_dir = REPORTS_BASE_DIR / run_folder
    
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Reports will be saved in: {output_dir}")

    # --- Loop and Generate Reports ---
    total_members = len(members_to_process)
    for i, email in enumerate(members_to_process):
        member_df = date_filtered_df[date_filtered_df['EmployeeEmail'].str.lower() == email].copy()
        
        if member_df.empty:
            continue
            
        employee_name = member_df.iloc[0]['EmployeeName'].replace(' ', '_')
        filename = f"Performance_Review_{employee_name}.xlsx"
        output_path = output_dir / filename
        
        print(f"\n({i+1}/{total_members}) Generating report for {email}...")
        generate_report(args.start_date, args.end_date, output_path, member_df)

if __name__ == "__main__":
    main()