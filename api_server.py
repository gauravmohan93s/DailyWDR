# -*- coding: utf-8 -*-
"""
KC Analytics API Server (FastAPI) - Enterprise Version (v1.2.4)
- SMART MERGE (UPSERT) for Roster Import: Preserves existing metadata.
- Strict Role Mapping: Prevents overwriting Roles with raw Designation strings.
"""

from fastapi import FastAPI, BackgroundTasks, WebSocket, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import pandas as pd
import sqlite3
import datetime as dt
import subprocess
import sys
import os
import asyncio
import io
import json
from pathlib import Path
from typing import List, Optional

from reporting_config import (
    LOCAL_DB_PATH, load_app_config, save_app_config
)
from data_loader import get_filter_options, norm_token

app = FastAPI(title="KC Reporting API", version="1.2.4")

app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

# --- Models ---
class AppConfig(BaseModel):
    REPORT_DATE_OVERRIDE: str
    DRY_RUN: bool
    DISK_CLEANUP_DAYS: int
    NAME_FILTER: Optional[str] = ""
    ROLE_FILTER: Optional[str] = ""
    REGION_FILTER: Optional[str] = ""
    SUBREGION_FILTER: Optional[str] = ""
    TEAM_FILTER: Optional[str] = ""

class MatrixEntry(BaseModel):
    Role: str
    Region: str
    SubRegion: str
    ManagerEmail: str
    IncludeInReporting: Optional[bool] = True

# --- Log Manager ---
class LogManager:
    def __init__(self):
        self.active_websockets: List[WebSocket] = []
        self.log_history: List[dict] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_websockets.append(websocket)
        for entry in self.log_history[-100:]: await websocket.send_json(entry)

    def disconnect(self, websocket: WebSocket):
        if websocket in self.active_websockets: self.active_websockets.remove(websocket)

    async def broadcast(self, message: str, type: str = "log", progress: int = None):
        entry = {"msg": message, "type": type, "progress": progress, "ts": dt.datetime.now().strftime("%H:%M:%S")}
        self.log_history.append(entry)
        if len(self.log_history) > 2000: self.log_history.pop(0)
        for connection in self.active_websockets:
            try: await connection.send_json(entry)
            except: pass

log_manager = LogManager()

# --- Utility Tasks ---
async def run_pipeline_task():
    await log_manager.broadcast("Initializing Daily Work Done Report pipeline...", type="status", progress=5)
    try:
        process = subprocess.Popen(
            [sys.executable, "run_daily_flow.py", "--json-progress"],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1, encoding='utf-8', errors='ignore'
        )
        for line in process.stdout:
            line = line.strip()
            if not line: continue
            if line.startswith('{"step":'):
                try:
                    data = json.loads(line)
                    await log_manager.broadcast(data["msg"], type="status", progress=data["progress"])
                except: await log_manager.broadcast(line)
            else: await log_manager.broadcast(line)
            await asyncio.sleep(0.01)
        process.wait()
        if process.returncode == 0: await log_manager.broadcast("Pipeline completed successfully!", type="success", progress=100)
        else: await log_manager.broadcast(f"Pipeline failed (Exit Code {process.returncode})", type="error")
    except Exception as e: await log_manager.broadcast(f"Fatal Error: {str(e)}", type="error")

def sync_roster_managers(conn: sqlite3.Connection):
    """Update Roster 'Manager' and 'Include' from Matrix."""
    try:
        roster = pd.read_sql("SELECT * FROM roster", conn)
        matrix = pd.read_sql("SELECT * FROM management_matrix", conn)
        if matrix.empty: return
        
        matrix["key"] = matrix.apply(lambda r: f"{norm_token(r.get('Role',''))}|{norm_token(r.get('Region',''))}|{norm_token(r.get('SubRegion',''))}", axis=1)
        m_map = {r["key"]: str(r.get("ManagerEmail","")).strip() for _, r in matrix.iterrows()}
        i_map = {r["key"]: int(r.get("IncludeInReporting", 1)) for _, r in matrix.iterrows()}

        def update_row(row):
            k = f"{norm_token(row.get('Role',''))}|{norm_token(row.get('Region',''))}|{norm_token(row.get('SubRegion',''))}"
            if k in m_map: row["Manager"] = m_map[k]
            if k in i_map: row["Include"] = i_map[k]
            return row

        roster = roster.apply(update_row, axis=1)
        roster.to_sql("roster", conn, if_exists="replace", index=False)
        conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_roster_email ON roster (EmployeeEmail)")
        print("[SYNC] Roster managers and inclusion synced.")
    except Exception as e: print(f"[SYNC ERROR] {e}")

# --- Endpoints ---
@app.get("/config")
def get_config():
    cfg = load_app_config()
    return {**cfg, "options": get_filter_options()}

@app.post("/config")
def update_config(config: AppConfig):
    save_app_config(config.dict())
    return {"status": "success"}

@app.get("/roster")
def get_roster(search: Optional[str] = None):
    conn = sqlite3.connect(LOCAL_DB_PATH)
    query = "SELECT * FROM roster"
    if search:
        query += f" WHERE EmployeeName LIKE '%{search}%' OR EmployeeEmail LIKE '%{search}%' OR Role LIKE '%{search}%' OR Region LIKE '%{search}%' OR SubRegion LIKE '%{search}%'"
    df = pd.read_sql(query, conn)
    conn.close()
    return df.to_dict(orient="records")

@app.post("/roster/import")
async def import_roster(file: UploadFile = File(...)):
    """Import roster with Smart Merge (UPSERT) to preserve existing metadata."""
    print(f"[IMPORT] Processing file: {file.filename}")
    contents = await file.read()
    try:
        if file.filename.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(io.BytesIO(contents), engine='openpyxl')
        else:
            df = pd.read_csv(io.BytesIO(contents))
        
        raw_cols = {str(c).lower().strip(): c for c in df.columns}
        
        # Intelligent Mapping (STRICT Role definition)
        mapping = {
            "EmployeeEmail": ["email", "e-mail", "employee email", "user email", "email id", "email id (official)", "official email"],
            "EmployeeName": ["name", "employee name", "full name", "staff name", "user name"],
            "Role": ["role", "job role", "position", "job title"], # 'designation' removed to prevent bad overwrites
            "Region": ["region", "zone", "state", "territory"],
            "SubRegion": ["subregion", "sub region", "city", "branch", "location"],
            "WeekOffs": ["weekoff", "week-off", "off day", "week offs", "allocated weekoffs", "week off"]
        }
        
        final_data = {}
        for target, aliases in mapping.items():
            found_col = next((raw_cols[a] for a in aliases if a in raw_cols), None)
            if found_col:
                final_data[target] = df[found_col]
            else:
                final_data[target] = None # Mark as missing in import

        incoming_df = pd.DataFrame(final_data)
        incoming_df["EmployeeEmail"] = incoming_df["EmployeeEmail"].astype(str).str.strip().str.lower()
        incoming_df = incoming_df[incoming_df["EmployeeEmail"].str.contains("@")].drop_duplicates(subset=["EmployeeEmail"]).copy()
        
        conn = sqlite3.connect(LOCAL_DB_PATH)
        # 1. Load existing roster
        try:
            existing_df = pd.read_sql("SELECT * FROM roster", conn)
        except:
            existing_df = pd.DataFrame(columns=["EmployeeEmail", "EmployeeName", "Role", "Region", "SubRegion", "WeekOffs", "Include", "Manager", "SendTo", "SaturdayOffPattern"])

        # 2. PERFORM SMART MERGE (UPSERT)
        processed_emails = set()
        for _, row in incoming_df.iterrows():
            email = row["EmployeeEmail"]
            processed_emails.add(email)
            # Only update columns that were actually present in the Excel file
            update_data = {k: v for k, v in row.items() if v is not None and str(v).strip() != ""}
            
            if email in existing_df["EmployeeEmail"].values:
                # Update existing record
                idx = existing_df[existing_df["EmployeeEmail"] == email].index[0]
                for col, val in update_data.items():
                    existing_df.at[idx, col] = val
            else:
                # Insert new record
                new_row = {c: "" for c in existing_df.columns}
                new_row.update(update_data)
                new_row["Include"] = 1 # Default to active
                existing_df = pd.concat([existing_df, pd.DataFrame([new_row])], ignore_index=True)

        # 3. Save merged result
        existing_df.to_sql("roster", conn, if_exists="replace", index=False)
        conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_roster_email ON roster (EmployeeEmail)")
        
        # Sync with Matrix to ensure Managers are correct for updated members
        sync_roster_managers(conn)
        
        conn.commit(); conn.close()
        
        print(f"[IMPORT] Success! {len(incoming_df)} members processed via Smart Merge.")
        return {"status": "success", "message": f"Successfully processed {len(incoming_df)} members. Existing Roles were preserved unless an explicit 'Role' column was found."}
    except Exception as e:
        print(f"[IMPORT ERROR] {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/matrix")
def get_matrix(search: Optional[str] = None):
    conn = sqlite3.connect(LOCAL_DB_PATH)
    try:
        try: matrix_df = pd.read_sql("SELECT * FROM management_matrix", conn)
        except: matrix_df = pd.DataFrame(columns=["Role", "Region", "SubRegion", "ManagerEmail", "IncludeInReporting"])

        # Auto-generate combinations
        roster_df = pd.read_sql("SELECT DISTINCT Role, Region, SubRegion FROM roster", conn)
        for df in [matrix_df, roster_df]:
            for c in ["Role", "Region", "SubRegion"]: df[c] = df[c].fillna("").astype(str).str.strip()

        merged = roster_df.merge(matrix_df, on=["Role", "Region", "SubRegion"], how="left")
        merged["ManagerEmail"] = merged["ManagerEmail"].fillna("")
        merged["IncludeInReporting"] = merged["IncludeInReporting"].fillna(1)
        
        if search:
            s = search.lower()
            mask = merged["Role"].str.lower().str.contains(s) | \
                   merged["Region"].str.lower().str.contains(s) | \
                   merged["SubRegion"].str.lower().str.contains(s) | \
                   merged["ManagerEmail"].str.lower().str.contains(s)
            merged = merged[mask]

        merged.to_sql("management_matrix", conn, if_exists="replace", index=False)
        conn.commit()
        return merged.to_dict(orient="records")
    except Exception as e: raise HTTPException(status_code=500, detail=str(e))
    finally: conn.close()

@app.post("/matrix")
def update_matrix(entries: List[MatrixEntry]):
    conn = sqlite3.connect(LOCAL_DB_PATH)
    try:
        df = pd.DataFrame([e.dict() for e in entries])
        df.to_sql("management_matrix", conn, if_exists="replace", index=False)
        sync_roster_managers(conn)
        conn.commit()
        return {"status": "success"}
    except Exception as e: raise HTTPException(status_code=500, detail=str(e))
    finally: conn.close()

@app.post("/run-pipeline")
def trigger_pipeline(background_tasks: BackgroundTasks):
    background_tasks.add_task(run_pipeline_task)
    return {"status": "success"}

@app.websocket("/ws/logs")
async def websocket_endpoint(websocket: WebSocket):
    await log_manager.connect(websocket)
    try:
        while True: await websocket.receive_text()
    except: log_manager.disconnect(websocket)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
