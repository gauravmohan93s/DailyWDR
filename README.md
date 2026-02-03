# DailyWD-UK Automation

**Status:** 🟢 OPERATIONAL | **Version:** 3.10

## 1. Overview

**Goal:** To fully automate the daily reporting of Work Done (WD) for the KC UK team, ensuring accurate performance tracking, gamified leaderboards, and timely delivery of insights to Staff, Managers, and Executives via email.

**Objectives:**
*   **Automation:** Eliminate manual data compilation from raw daily dumps.
*   **Accuracy:** Ensure "Action Group" logic and "KRA Scores" are calculated consistently.
*   **Timeliness:** Deliver reports by the start of the next working day.
*   **Engagement:** Use gamification (Badges, Streaks, Leaderboards) to motivate staff.
*   **Reliability:** Handle missing data, holidays, and Sundays gracefully.

---

## 2. Architecture

The system follows a linear pipeline:

1.  **Extraction (`daily_wd_data_extractor.py`)**
    *   Fetches data from SQL Server (primary) or CSVs (fallback).
    *   Upserts data into a local SQLite database (`kc_reports.db`).
    *   Uses SHA256 hashing to ensure idempotent upserts (no duplicates).

2.  **Summarization (`wd_summariser.py`)**
    *   Processes raw actions from SQLite.
    *   Calculates KRAs, streaks, and badges.
    *   Generates a summary Excel file (`summary_all_days.xlsx`) for the reporting layer.

3.  **Orchestration (`run_daily_flow.py`)**
    *   Manages the sequence: Extract -> Summarize -> Email.

4.  **Distribution**
    *   `exec_email.py`: Organization-wide roll-up for executives.
    *   `mg_tl_email.py`: Team-specific roll-up for Managers/TLs.
    *   `staff_email.py`: Individual performance cards and certificates for staff.

### Directory Structure

```text
DailyWD-UK/
├── logs/                   # (Output) Execution logs
├── on_demand_reports/      # (Output) Custom generated reports
├── kc_reports.db           # (Local DB) SQLite database storing all history
├── summary_all_days.xlsx   # (Intermediate) Processed summary data
├── daily_wd_data_extractor.py
├── wd_summariser.py
├── run_daily_flow.py
├── exec_email.py
├── mg_tl_email.py
├── staff_email.py
├── win32_utils.py          # Outlook automation helpers
└── wd_settings.xlsx        # (Config) Roster, Measures, Holidays (Keep in OneDrive)
```

---

## 3. Installation & Setup

### Prerequisites
*   Python 3.10+
*   Dependencies: `pandas`, `numpy`, `openpyxl`, `sqlalchemy`, `pyodbc`, `pdfkit`
*   **wkhtmltopdf** installed and in system PATH (for PDF generation).
*   **Microsoft Outlook** (Desktop) installed and logged in (for email sending).

### Configuration
Set the following environment variables for security:
*   `KC_DB_PWD`: MSSQL database password.
*   `SMTP_PASSWORD`: SMTP password (if using SMTP instead of Outlook).

### Scheduled Execution (Windows Task Scheduler)

Recommended schedule to ensure data is ready by next morning:

| Task | Frequency | Time (IST) | Script |
|------|-----------|-----------|--------|
| Database Extract | Weekly | Mon 22:00 | `daily_wd_data_extractor.py` |
| Daily Summarize | Daily | 23:00 | `wd_summariser.py` |
| Staff Mailer | Daily | 23:30 | `staff_email.py` |
| Manager Digest | Daily | 23:35 | `mg_tl_email.py` |
| Executive Digest | Daily | 23:40 | `exec_email.py` |

**PowerShell Setup Command:**
```powershell
# Example for Summarizer (Run as Admin)
$Action = New-ScheduledTaskAction -Execute "python" -Argument "C:\Path\To\wd_summariser.py" -WorkingDirectory "C:\Path\To\Project"
$Trigger = New-ScheduledTaskTrigger -Daily -At "23:00"
Register-ScheduledTask -Action $Action -Trigger $Trigger -TaskName "DailyWDR-Summarize"
```

---

## 4. Troubleshooting

### Common Issues
1.  **Emails not sending**:
    *   Check `logs/` for error messages.
    *   Ensure Outlook is running.
    *   Verify `DRY_RUN = False` in script if ready for production.
2.  **PDF Attachment Missing**:
    *   Ensure `wkhtmltopdf` is installed.
    *   Check logs for "PDF generation skipped" warnings.
3.  **Database Connection Failed**:
    *   Verify VPN connection (if required).
    *   Check `KC_DB_PWD` environment variable.

---

## 5. Refactoring Plan (Future V2.0)

**Goal:** Transform into a modular, API-driven system.

*   **Phase 1 (Core)**: Move to a package structure (`core/`, `logic/`, `reporting/`).
*   **Phase 2 (Logic)**: Separate calculation logic from summarization.
*   **Phase 3 (Reporting)**: Implement a query engine for dynamic reports.
*   **Phase 4 (Cleanup)**: Remove dependency on `summary_all_days.xlsx`.

---

## 6. Changelog

### [3.10] - 2026-01-24
*   **Fixed**: MTD calculation issue (increased `MAX_DAYS_PER_RUN` to 40).
*   **Fixed**: Empty PDF attachments (added validation and size checks).
*   **Added**: Comprehensive Documentation & Cleanup.

### [3.9] - 2025-10-15
*   **Changed**: Migrated from Excel-only to SQLite (`kc_reports.db`).
*   **Added**: `daily_wd_data_extractor.py` with idempotent upsert.

### [3.8] - 2025-08-20
*   **Changed**: Executive & Manager digests expanded with leaderboards/badges.

### [3.0] - 2025-06-01
*   **Initial Release**: Multi-tier email reporting, Gamification, PDF generation.