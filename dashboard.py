# -*- coding: utf-8 -*-
"""
KC Report Command Center — Interactive Dashboard (v1.0)
Non-technical UI for running reports and managing settings.
"""

import streamlit as st
import pandas as pd
import datetime as dt
import subprocess
import sys
import os
import time
from pathlib import Path
from reporting_config import (
    SUMMARY_PATH, SETTINGS_PATH, LOCAL_DB_PATH, 
    REPORT_DATE_OVERRIDE, DRY_RUN_GLOBAL,
    load_app_config, save_app_config
)
from data_loader import load_settings

# Page Config
st.set_page_config(page_title="KC Report Center", page_icon="📊", layout="wide")

# Custom CSS for a professional look
st.markdown("""
<style>
    .main { background-color: #F5F7FB; }
    .stButton>button { width: 100%; border-radius: 8px; height: 3em; background-color: #2563EB; color: white; font-weight: bold; }
    .stButton>button:hover { background-color: #1D4ED8; border: none; }
    .status-box { padding: 20px; border-radius: 12px; background-color: white; border: 1px solid #E5E7EB; }
</style>
""", unsafe_allow_html=True)

# --- Sidebar Navigation ---
st.sidebar.title("📊 KC Analytics")
st.sidebar.markdown("---")
page = st.sidebar.radio("Navigation", ["Operations", "Team Management", "View History"])

# --- Load Current Config ---
app_cfg = load_app_config()

# --- Page: Operations ---
if page == "Operations":
    st.title("🚀 Report Operations")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("1. Setup")
        target_date = st.date_input("Target Report Date", 
                                    value=dt.date.fromisoformat(app_cfg["REPORT_DATE_OVERRIDE"]))
        
        is_dry_run = st.checkbox("Dry Run Mode (Preview Only)", value=app_cfg["DRY_RUN"], 
                                help="If checked, reports will be generated but NOT sent to anyone.")
        
        if st.button("💾 Save Configuration"):
            app_cfg["REPORT_DATE_OVERRIDE"] = target_date.isoformat()
            app_cfg["DRY_RUN"] = is_dry_run
            save_app_config(app_cfg)
            st.success("Configuration updated! Pipeline is ready.")

    with col2:
        st.subheader("2. Quick Stats")
        try:
            # Show member count
            cfg = load_settings(SETTINGS_PATH)
            st.metric("Total Team Members", len(cfg['team']))
            st.metric("Active Roles", len(cfg['role_kra_codes']))
        except:
            st.warning("Could not load stats. Check database connection.")

    st.markdown("---")
    st.subheader("3. Execute Pipeline")
    
    if st.button("🔥 START DAILY FLOW"):
        st.info("Pipeline started... please wait.")
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # We run the actual run_daily_flow.py as a subprocess to capture logs
        try:
            process = subprocess.Popen(
                [sys.executable, "run_daily_flow.py"],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1
            )
            
            log_area = st.empty()
            full_log = ""
            
            for line in process.stdout:
                full_log += line
                log_area.code(full_log[-2000:]) # Show last 2000 chars
                
                # Mock progress for UI feel
                if "Validate Settings" in line: progress_bar.progress(10)
                if "Database" in line: progress_bar.progress(30)
                if "Summariser" in line: progress_bar.progress(50)
                if "Executive" in line: progress_bar.progress(70)
                if "Manager" in line: progress_bar.progress(85)
                if "Staff" in line: progress_bar.progress(95)

            process.wait()
            progress_bar.progress(100)
            if process.returncode == 0:
                st.success("✅ DAILY FLOW COMPLETED SUCCESSFULLY!")
            else:
                st.error("❌ Pipeline failed. Check the logs above for details.")
                
        except Exception as e:
            st.error(f"Failed to trigger script: {e}")

# --- Page: Team Management ---
elif page == "Team Management":
    st.title("👥 Team Management")
    st.markdown("Manage your team roster directly in the database. These changes take effect immediately.")
    
    try:
        import sqlite3
        conn = sqlite3.connect(LOCAL_DB_PATH)
        df_roster = pd.read_sql("SELECT * FROM roster", conn)
        
        st.subheader("Current Roster")
        edited_df = st.data_editor(df_roster, num_rows="dynamic", use_container_width=True)
        
        if st.button("💾 Save Roster Changes"):
            # Basic Validation
            if edited_df['EmployeeEmail'].duplicated().any():
                st.error("Error: Duplicate emails are not allowed.")
            else:
                edited_df.to_sql("roster", conn, if_exists="replace", index=False)
                st.success("Roster updated and synced to database!")
        
        conn.close()
    except Exception as e:
        st.error(f"Could not load roster: {e}")

# --- Page: View History ---
elif page == "View History":
    st.title("📜 Sending History")
    
    history_file = Path("sent_log_detail.csv")
    if history_file.exists():
        df_hist = pd.read_csv(history_file)
        st.dataframe(df_hist.sort_values("TimestampUTC", ascending=False), use_container_width=True)
    else:
        st.info("No sending history found yet.")

# Footer
st.sidebar.markdown("---")
st.sidebar.caption(f"v1.0.0-Stable | May 2026")
