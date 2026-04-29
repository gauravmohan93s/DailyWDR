# Project Status: DailyWD-UK Automation

**Date:** 23 Mar 2026
**Version:** V1.4

## 1. Goal
To fully automate the daily reporting of Work Done (WD) for the KC UK team, ensuring accurate performance tracking, gamified leaderboards, and timely delivery of insights to Staff, Managers, and Executives via email.

## 2. Objective
*   **Automation:** Eliminate manual data compilation from raw daily dumps.
*   **Accuracy:** Ensure "Action Group" logic and "KRA Scores" are calculated consistently across all reports.
*   **Timeliness:** Deliver reports by the start of the next working day.
*   **Engagement:** Use gamification (Badges, Streaks, Leaderboards) to motivate staff.
*   **Reliability:** Handle missing data, holidays, and Sundays gracefully.

## 3. Current Status & Architecture
The pipeline is currently **functional** and **active**.

### Architecture
1.  **Extraction (`daily_wd_data_extractor.py`):** Fetches data from SQL Server (primary) or CSVs (fallback) and upserts it into a local SQLite database (`kc_reports.db`).
2.  **Summarization (`wd_summariser.py`):** Processes raw actions from SQLite into a single "All Days" summary Excel file (`summary_all_days.xlsx`), calculating KRAs, streaks, and badges.
3.  **Orchestration (`run_daily_flow.py`):** Manages the sequence: Extract -> Summarize -> Email.
4.  **Distribution:**
    *   `exec_email.py`: Organization-wide roll-up for executives.
    *   `mg_tl_email.py`: Team-specific roll-up for Managers/TLs.
    *   `staff_email.py`: Individual performance cards and certificates for staff.

### Progress (Recent Maintenance)
*   **Missing Submission Fix:** Resolved an issue where records with NULL timestamps in the source DB were being excluded from reports despite having a valid IST date assigned during extraction.
*   **Resilient Summarization:** Modified `wd_summariser.py` to preserve existing database dates, preventing destructive overwrites when source timestamps are missing.
*   **Robust Extraction:** Enhanced `daily_wd_data_extractor.py` to more effectively handle empty or invalid source dates using a fallback date system.
*   **Database Repair:** Successfully repaired 900+ historical records that were missing action timestamps by intelligently matching them with other actions on the same ID.
*   **Cleanup:** Removed redundant debugging and repair scripts.

## 4. Critique & Issues
While functional, the project has several technical debts and risks:

*   **Security (CRITICAL):**
    *   `daily_wd_data_extractor.py` contains **hardcoded database credentials** in plain text. This must be moved to environment variables or a secure configuration file immediately.
*   **Portability:**
    *   Scripts utilize **hardcoded absolute paths** (e.g., `C:\Users\gsakhare\...`). This prevents the code from running on other machines or CI/CD pipelines without modification.
*   **Scalability / Efficiency:**
    *   `wd_summariser.py` re-calculates the *entire* history (from `kc_reports.db`) on every run. As the dataset grows, this will become a performance bottleneck. Incremental processing should be considered for future versions.
    *   The system relies on local SQLite and Excel files on OneDrive, which is suitable for a single-user setup but fragile for enterprise scaling.
*   **Dependency:**
    *   Heavy reliance on `win32com` (Outlook desktop automation). This requires an active logged-in Windows session and cannot be easily deployed to a cloud server.

## 5. Roadmap & Future Context (The Path to V2.0)
To transition from a functional local automation to a robust enterprise-grade pipeline, the following phases are planned:

### Phase 1: Security & Portability (High Priority)
*   **Credential Externalization:** Move RDS and MSSQL credentials to an encrypted `.env` file (using `python-dotenv`).
*   **Path Decoupling:** Replace all absolute `C:\Users\gsakhare\...` paths with `pathlib` relative paths or a `config.json` base path.
*   **Environment Validation:** Add a pre-flight check to ensure all dependencies (`wkhtmltopdf`, Outlook, etc.) are available before the run begins.

### Phase 2: Architectural Consolidation
*   **Unified Reporting Engine:** Refactor `exec_email.py`, `mg_tl_email.py`, and `staff_email.py` into a shared `ReportGenerator` and `EmailDispatcher` module. This will reduce code duplication by ~60%.
*   **Validation Layer:** Implement a script to validate `wd_settings.xlsx` (checking for orphan MeasureCodes or missing Role weights) before processing starts.

### Phase 3: Performance & Scalability
*   **Incremental Summarization:** Modify `wd_summariser.py` to only process the current month’s delta and merge it with a "Historical Archive" table in SQLite.
*   **Data Quality Auditing:** Implement a "Dead-Letter Log" to track and report records that require "Mature Date Guessing" (NULL timestamp fallbacks).

### Phase 4: Cloud Readiness (Optional)
*   **API Integration:** Transition from `win32com` (Outlook Desktop) to **Microsoft Graph API**.
*   **Containerization:** Dockerize the pipeline to allow it to run on a lightweight Windows Server or a CI/CD runner.

---
**Status:** 🟢 OPERATIONAL | **Current Focus:** Phase 1 (Security & Portability)
