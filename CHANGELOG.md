# Changelog

All notable changes to the DailyWDR project are documented here.

## [3.13] - 2026-04-29

### Fixed

#### Repetitive Team Members in Reports
- **Problem**: Team members were appearing multiple times in "Role-wise Work" tables, and "Total Work Done" was inflated.
- **Root Cause**: Duplicate email entries in the "Team" sheet of `wd_settings.xlsx` caused the roster to balloon during merges.
- **Solution**: 
  - **Roster Deduplication**: Added logic to `wd_summariser.py`, `exec_email.py`, `mg_tl_email.py`, and `staff_email.py` to deduplicate the team roster based on email during settings loading.
  - **Data Sanitization**: Added secondary deduplication to report loading logic to ensure each employee has only one record per day, even if existing summary data contains duplicates.
- **Files Modified**: `wd_summariser.py`, `exec_email.py`, `mg_tl_email.py`, `staff_email.py`
