# Technical Refactor Plan: DailyWD-UK V2.0

This document outlines the steps to move the DailyWD-UK pipeline into a professional, server-ready state.

## 🟢 PHASE 1: Core Optimization (COMPLETED)
1. **Centralized Data & Settings Loader:** Consolidated logic into `data_loader.py`.
2. **Database-Backed Summaries:** Migrated Excel to SQLite `daily_summary`.
3. **Automated Settings Validator:** Pre-flight check to catch human errors.
4. **SQL-Level Filtering:** High-performance data fetching.
5. **Fully Centralized Paths:** Single source of truth for all file paths.
6. **Decoupled Parallelism:** Stability for Outlook sending.
7. **Automated Disk Cleanup:** Storage maintenance.

## 🔵 PHASE 2: Full-Stack Modernization (NEXT.JS + FASTAPI)
**Goal:** Replace basic scripts with an enterprise-grade web application.

### 1. Backend Service (FastAPI)
* **Status:** IN PROGRESS
* Implement `api_server.py` to expose reporting logic via REST.
* Add WebSocket support for real-time log streaming.
* Wrap existing `run_daily_flow.py` as a background task.

### 2. Frontend Dashboard (Next.js 14)
* **Goal:** A beautiful, responsive UI for non-technical users.
* Use Tailwind CSS and modern components for the "Command Center".
* Create a dedicated "Team Manager" grid with real-time database syncing.

### 3. API Integration & Packaging
* Connect the Next.js frontend to the FastAPI backend.
* Create a single `start_system.bat` file to launch both servers.

---
**Updated on:** 12 May 2026 (Full-Stack Modernization Phase)
