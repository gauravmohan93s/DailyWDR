# -*- coding: utf-8 -*-
"""
KC Analytics API Server (FastAPI) - Enterprise Version
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
from data_loader import get_filter_options

app = FastAPI(title="KC Reporting API", version="1.2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- Models ---
class AppConfig(BaseModel):
    REPORT_DATE_OVERRIDE: str
    DRY_RUN: bool
    DISK_CLEANUP_DAYS: int
    NAME_FILTER: Optional[str] = ""
    ROLE_FILTER: Optional[str] = ""
    TEAM_FILTER: Optional[str] = ""

class MatrixEntry(BaseModel):
    Role: str
    Region: str
    SubRegion: str
    ManagerEmail: str

# --- Global State for Logs & Progress ---
class LogManager:
    def __init__(self):
        self.active_websockets: List[WebSocket] = []
        self.log_history: List[dict] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_websockets.append(websocket)
        for entry in self.log_history[-100:]:
            await websocket.send_json(entry)

    def disconnect(self, websocket: WebSocket):
        if websocket in self.active_websockets:
            self.active_websockets.remove(websocket)

    async def broadcast(self, message: str, type: str = "log", progress: int = None):
        entry = {"msg": message, "type": type, "progress": progress, "ts": dt.datetime.now().strftime("%H:%M:%S")}
        self.log_history.append(entry)
        if len(self.log_history) > 2000: self.log_history.pop(0)
        for connection in self.active_websockets:
            try: await connection.send_json(entry)
            except: pass

log_manager = LogManager()

# --- Utility: Run Pipeline in Background ---
async def run_pipeline_task():
    await log_manager.broadcast("Initializing Daily Work Done Report pipeline...", type="status", progress=5)
    try:
        # Use --json-progress flag (we'll update run_daily_flow.py to support this)
        process = subprocess.Popen(
            [sys.executable, "run_daily_flow.py", "--json-progress"],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            encoding='utf-8',
            errors='ignore'
        )
        
        for line in process.stdout:
            line = line.strip()
            if not line: continue
            
            # Check if line is a progress JSON
            if line.startswith('{"step":'):
                try:
                    data = json.loads(line)
                    await log_manager.broadcast(data["msg"], type="status", progress=data["progress"])
                except:
                    await log_manager.broadcast(line)
            else:
                await log_manager.broadcast(line)
            await asyncio.sleep(0.01)

        process.wait()
        if process.returncode == 0:
            await log_manager.broadcast("Pipeline completed successfully!", type="success", progress=100)
        else:
            await log_manager.broadcast(f"Pipeline failed (Exit Code {process.returncode})", type="error")
            
    except Exception as e:
        await log_manager.broadcast(f"Fatal Error: {str(e)}", type="error")

# --- Endpoints ---

@app.get("/config")
def get_config():
    cfg = load_app_config()
    opts = get_filter_options()
    return {**cfg, "options": opts}

@app.post("/config")
def update_config(config: AppConfig):
    save_app_config(config.dict())
    return {"status": "success"}

@app.get("/roster")
def get_roster(search: Optional[str] = None):
    conn = sqlite3.connect(LOCAL_DB_PATH)
    query = "SELECT * FROM roster"
    if search:
        query += f" WHERE EmployeeName LIKE '%{search}%' OR EmployeeEmail LIKE '%{search}%' OR Role LIKE '%{search}%'"
    df = pd.read_sql(query, conn)
    conn.close()
    return df.to_dict(orient="records")

@app.post("/roster/import")
async def import_roster(file: UploadFile = File(...)):
    contents = await file.read()
    try:
        df = pd.read_excel(io.BytesIO(contents)) if file.filename.endswith(('.xlsx', '.xls')) else pd.read_csv(io.BytesIO(contents))
        # Intelligent Mapping: search for likely column names
        mapping = {
            "EmployeeEmail": ["email", "e-mail", "employee email", "user email"],
            "EmployeeName": ["name", "employee name", "full name", "staff name"],
            "Role": ["role", "designation", "position"],
            "Region": ["region", "zone", "state"],
            "SubRegion": ["subregion", "sub region", "city", "branch"],
            "WeekOffs": ["weekoff", "week-off", "off day", "week offs"]
        }
        final_cols = {}
        for target, aliases in mapping.items():
            for col in df.columns:
                if str(col).lower().strip() in aliases:
                    final_cols[target] = col
                    break
        
        if "EmployeeEmail" not in final_cols:
            raise HTTPException(status_code=400, detail="Could not identify 'Email' column in the file.")
            
        # Rebuild DF with standard columns
        new_df = pd.DataFrame()
        for target, source in final_cols.items():
            new_df[target] = df[source]
        
        new_df["EmployeeEmail"] = new_df["EmployeeEmail"].astype(str).str.strip().str.lower()
        new_df = new_df.drop_duplicates(subset=["EmployeeEmail"])
        
        conn = sqlite3.connect(LOCAL_DB_PATH)
        # Ensure target table exists with same schema or replace
        new_df.to_sql("roster", conn, if_exists="replace", index=False)
        conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_roster_email ON roster (EmployeeEmail)")
        conn.commit(); conn.close()
        
        return {"status": "success", "message": f"Successfully imported {len(new_df)} members with intelligent mapping."}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/matrix")
def get_matrix():
    conn = sqlite3.connect(LOCAL_DB_PATH)
    try:
        df = pd.read_sql("SELECT * FROM management_matrix", conn)
    except:
        df = pd.DataFrame(columns=["Role", "Region", "SubRegion", "ManagerEmail"])
    conn.close()
    return df.to_dict(orient="records")

@app.post("/matrix")
def update_matrix(entries: List[MatrixEntry]):
    conn = sqlite3.connect(LOCAL_DB_PATH)
    try:
        df = pd.DataFrame([e.dict() for e in entries])
        df.to_sql("management_matrix", conn, if_exists="replace", index=False)
        conn.commit()
        return {"status": "success"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        conn.close()

@app.post("/run-pipeline")
def trigger_pipeline(background_tasks: BackgroundTasks):
    background_tasks.add_task(run_pipeline_task)
    return {"status": "success"}

@app.websocket("/ws/logs")
async def websocket_endpoint(websocket: WebSocket):
    await log_manager.connect(websocket)
    try:
        while True:
            await websocket.receive_text()
    except:
        log_manager.disconnect(websocket)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
