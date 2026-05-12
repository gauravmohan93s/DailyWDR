# Technical Refactor Plan: DailyWD-UK V2.0

This document outlines the specific technical steps needed to move the DailyWD-UK pipeline into a professional, server-ready state.

## 1. Centralized Data & Settings Loader (DRY)
**Goal:** Consolidate data loading and settings logic into a single source of truth.
*   **Status:** ✅ DONE
*   **Implementation:** Created `data_loader.py`. All scripts now import `load_settings` and `load_summary` from here.

## 2. Transition to Database-Backed Summaries
**Goal:** Replace `summary_all_days.xlsx` with a dedicated SQLite table for reliability.
*   **Status:** ✅ DONE
*   **Implementation:** Created `daily_summary` table in `kc_reports.db`. `wd_summariser.py` syncs to both DB and Excel.

## 3. Automated Settings & Data Validator
**Goal:** Proactively catch manual entry errors before reports are sent.
*   **Status:** ✅ DONE
*   **Implementation:** Created `settings_validator.py`. Integrated as "Pre-flight Check" in `run_daily_flow.py`.

## 4. Fully Centralized Paths
**Goal:** Move every hardcoded file path into `reporting_config.py`.
*   **Implementation:** Consolidate `SETTINGS_PATH`, `SUMMARY_PATH`, `OUT_DIR`, etc., into one configuration file.

## 5. SQL-Level Filtering (Performance)
**Goal:** Optimize `staff_email.py` by fetching only required rows from SQLite.
*   **Implementation:** Replace `SELECT *` with `WHERE EmployeeEmail=? AND ActionDateIST_Date=?`.

## 6. Decoupled Parallel Processing (Stability)
**Goal:** Use parallel generation for files (fast) but sequential sending for Outlook (stable).
*   **Implementation:** Refactor `staff_email.py` into two loops: one parallel (PDF/Excel) and one sequential (Email).

## 7. Automated Disk Cleanup
**Goal:** Prevent preview folders from growing indefinitely.
*   **Implementation:** Add a step to `run_daily_flow.py` to delete preview files older than 14 days.

---
**Updated on:** 12 May 2026 (Advanced Optimization Phase)
