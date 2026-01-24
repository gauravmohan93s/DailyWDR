# Changelog

All notable changes to the DailyWDR project are documented here.

## [3.10] - 2026-01-24

### Fixed

#### MTD (Month-to-Date) Calculation Issue
- **Problem**: MTD values were showing only the last 10 days instead of cumulative from the 1st of the month
- **Root Cause**: The `MAX_DAYS_PER_RUN` configuration in `daily_wd_data_extractor.py` was set to 10 days, limiting CSV imports during daily updates
- **Solution**: 
  - Increased `MAX_DAYS_PER_RUN` from 10 to 40 days
  - This allows full month backfill during initialization and subsequent daily imports
  - The database now accumulates all historical data; MTD calculations in `wd_summariser.py` correctly compute cumulative sums from database records
- **Files Modified**: `daily_wd_data_extractor.py` (line 51)

#### PDF Attachment Issue ("Install wkhtmltopdf..." with No Content)
- **Problem**: PDF attachments were empty or showed an error message instead of the generated report
- **Root Causes**:
  1. HTML content passed to `pdfkit` was not validated before generation
  2. PDF generation errors were silently caught without proper feedback
  3. Empty or malformed PDF files were still attached to emails
- **Solution**:
  - Added HTML content validation (check for non-empty string)
  - Added file size check after PDF generation to confirm successful output
  - Improved error messages with specific failure reasons
  - Modified attachment logic to only attach PDFs that successfully generated
  - Added warning logs when PDF generation is skipped
- **Files Modified**:
  - `exec_email.py` (lines 737–763, 948–957)
  - `mg_tl_email.py` (lines 804–821, 1045–1050)
  - `staff_eamil.py` (lines 587–612, 1276–1282)

### Changed

- Improved error handling in all three email scripts for PDF generation failures
- Enhanced logging output to aid troubleshooting PDF generation issues
- Conditional PDF attachment (only if generation succeeded) prevents sending incomplete reports

### Added

- **Documentation**:
  - Comprehensive `README.md` with installation, usage, configuration, and troubleshooting guides
  - `LICENSE` file (MIT License) for open-source distribution
  - `.gitignore` for Python, IDE, logs, and sensitive credentials
  - `CHANGELOG.md` (this file) for version history

- **Repository Cleanup**:
  - Removed all temporary debug markdown files (AGENTS.md, FIX_SUMMARY.md, etc.)
  - Removed test scripts (test_run.py, verify_fixes.py, check_db_*.py)
  - Removed generated cache (`__pycache__`, logs/)
  - Removed temporary data files (sent_log.csv, etc.)

### Repository Status

- Cleaned for public distribution
- Ready for GitHub push to development branch
- All configuration and secrets must be stored in environment variables (enforced via `.gitignore`)

## [3.9] - 2025-10-15

### Changed

- Database-centric architecture: Migrated from Excel-only to SQLite (`kc_reports.db`)
- wd_summariser now loads all data from local SQLite instead of merged_actions.xlsx
- Improved performance for large datasets (6+ months of data)

### Added

- `daily_wd_data_extractor.py`: New extraction & upsert module
- Idempotent upsert logic via unique_id (SHA256 hash of action key fields)
- Support for both MSSQL source DB and CSV backfill

## [3.8] - 2025-08-20

### Changed

- Executive digest now includes ORG-wide leaderboards and team badges
- Manager digest expanded with team-level KRA analysis

## [3.0] - 2025-06-01

### Initial Release

- Multi-tier email reporting: staff, managers, executives
- Dynamic measure system via wd_settings.xlsx
- Availability-weighted KRA calculations
- Month-to-date tracking with working-day exclusions
- Badge system with gamification
- PDF generation via wkhtmltopdf
- Audit trail via sent_log.csv

---

**Next Steps**:
- Push to GitHub repository
- Set up CI/CD pipeline for automated testing
- Document additional run configurations (scheduling with Task Scheduler)
