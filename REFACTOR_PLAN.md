# Technical Refactor Plan: DailyWD-UK V2.0

This document outlines the specific technical steps needed to move the DailyWD-UK pipeline into a professional, server-ready state.

## 1. Centralized Data & Settings Loader (DRY)
**Goal:** Consolidate data loading and settings logic into a single source of truth.
*   **Current Issue:** `load_settings` and `load_summary` are duplicated across 4 scripts, leading to "logic drift."
*   **Implementation:** 
    *   Create `data_loader.py`.
    *   Implement centralized, deduplicated `load_settings()` and `load_summary()`.
    *   Update all scripts to import from `data_loader.py`.

## 2. Transition to Database-Backed Summaries
**Goal:** Replace `summary_all_days.xlsx` with a dedicated SQLite table.
*   **Current Issue:** Large Excel files are slow, prone to corruption, and caused "Total Work Done" inflation due to roster duplication.
*   **Implementation:**
    *   Create a `daily_summary` table in `kc_reports.db`.
    *   Update `wd_summariser.py` to write results to both DB and Excel (for human viewing).
    *   Update email scripts to read from the DB table for maximum performance and reliability.

## 3. Automated Settings & Data Validator
**Goal:** Proactively catch manual entry errors before reports are sent.
*   **Implementation:**
    *   Create `settings_validator.py`.
    *   Check for: Duplicate emails in roster, missing managers, invalid MeasureCodes.
    *   Integrate into `run_daily_flow.py` as a "Pre-flight Check."

## 4. System Health & Alerting
**Goal:** Monitoring background task execution.
*   **Implementation:**
    *   Implement centralized logging to `logs/system_health.log`.
    *   Send a 1-line "Health Status" email to the admin after `run_daily_flow.py` completes.

## 5. Parallel Report Generation
**Goal:** Scale the system for large teams.
*   **Implementation:**
    *   Use Python's `concurrent.futures` or `multiprocessing` to generate PDFs in parallel.
    *   Significantly reduces execution time for `staff_email.py`.

---
**Updated on:** 29 Apr 2026 (Fix for Repetitive Team Members)
