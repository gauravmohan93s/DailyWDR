# Windows Task Scheduler Setup Guide

This guide explains how to schedule DailyWDR scripts to run automatically using Windows Task Scheduler.

## Overview

The DailyWDR system should run on the following schedule:

| Task | Frequency | Time (IST) | Script | Duration |
|------|-----------|-----------|--------|----------|
| Database Extract | Weekly | Monday 22:00 | `daily_wd_data_extractor.py` | 5-10 mins |
| Daily Summarize | Daily | 23:00 | `wd_summariser.py` | 2-3 mins |
| Staff Mailer | Daily | 23:30 | `staff_eamil.py` | 3-5 mins |
| Manager Digest | Daily | 23:35 | `mg_tl_email.py` | 1 min |
| Executive Digest | Daily | 23:40 | `exec_email.py` | 30 secs |

## Prerequisites

1. **Python Path**: Confirm Python 3.10+ is installed
   ```powershell
   python --version
   ```

2. **Required Packages**: Install dependencies
   ```bash
   pip install pandas numpy openpyxl sqlalchemy pyodbc pdfkit
   ```

3. **Script Location**: All scripts in `C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2\`

## Setup Instructions

### Method 1: Using PowerShell (Recommended)

Open PowerShell as Administrator and execute:

```powershell
# 1. Database Extract - Weekly Monday 22:00
$Action = New-ScheduledTaskAction -Execute "python" -Argument "C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2\daily_wd_data_extractor.py" -WorkingDirectory "C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2"
$Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At "22:00"
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive
Register-ScheduledTask -Action $Action -Trigger $Trigger -Principal $Principal -TaskName "DailyWDR-Extract" -Description "Weekly extraction from MSSQL to SQLite"

# 2. Daily Summarize - Daily 23:00
$Action = New-ScheduledTaskAction -Execute "python" -Argument "C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2\wd_summariser.py" -WorkingDirectory "C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2"
$Trigger = New-ScheduledTaskTrigger -Daily -At "23:00"
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive
Register-ScheduledTask -Action $Action -Trigger $Trigger -Principal $Principal -TaskName "DailyWDR-Summarize" -Description "Daily summary and metrics generation"

# 3. Staff Mailer - Daily 23:30
$Action = New-ScheduledTaskAction -Execute "python" -Argument "C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2\staff_eamil.py" -WorkingDirectory "C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2"
$Trigger = New-ScheduledTaskTrigger -Daily -At "23:30"
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive
Register-ScheduledTask -Action $Action -Trigger $Trigger -Principal $Principal -TaskName "DailyWDR-StaffMail" -Description "Send individual staff emails"

# 4. Manager Digest - Daily 23:35
$Action = New-ScheduledTaskAction -Execute "python" -Argument "C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2\mg_tl_email.py" -WorkingDirectory "C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2"
$Trigger = New-ScheduledTaskTrigger -Daily -At "23:35"
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive
Register-ScheduledTask -Action $Action -Trigger $Trigger -Principal $Principal -TaskName "DailyWDR-Manager" -Description "Send manager digests"

# 5. Executive Digest - Daily 23:40
$Action = New-ScheduledTaskAction -Execute "python" -Argument "C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2\exec_email.py" -WorkingDirectory "C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2"
$Trigger = New-ScheduledTaskTrigger -Daily -At "23:40"
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive
Register-ScheduledTask -Action $Action -Trigger $Trigger -Principal $Principal -TaskName "DailyWDR-Executive" -Description "Send executive digests"
```

### Method 2: Manual Setup in Task Scheduler GUI

1. **Open Task Scheduler**:
   - Press `Win + R`, type `taskschd.msc`, press Enter

2. **Create New Task**:
   - Right-click "Task Scheduler Library" → "Create Basic Task..."
   - Name: `DailyWDR-Extract`
   - Trigger: Weekly (Monday, 22:00)
   - Action:
     - Program: `python.exe`
     - Arguments: `C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2\daily_wd_data_extractor.py`
     - Start in: `C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2`

3. **Repeat for other scripts** with appropriate times:
   - `DailyWDR-Summarize` (23:00 daily)
   - `DailyWDR-StaffMail` (23:30 daily)
   - `DailyWDR-Manager` (23:35 daily)
   - `DailyWDR-Executive` (23:40 daily)

## Environment Configuration

### Set Credentials as Environment Variables

1. **Open System Environment Variables**:
   - Search "Edit the system environment variables"
   - Click "Environment Variables..."

2. **Add User Variables**:
   - `KC_DB_PWD`: Your MSSQL database password
   - `SMTP_PASSWORD`: Your SMTP password (if using SMTP)

   **Format**:
   ```
   Variable Name: KC_DB_PWD
   Variable Value: your_actual_password
   ```

3. **Verify in Scripts**:
   All scripts use `os.environ.get()` to fetch these values safely.

## Task Execution Verification

### View Task Logs

1. **Last Run Status**:
   - Open Task Scheduler
   - Select task → Check "Last Run Time" and "Last Run Result"

2. **Output Logs**:
   - Logs are saved in:
     ```
     C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\raw\
     ```
   - Files: `sent_log.csv`, `sent_log_detail.csv`

3. **Email Preview Folders**:
   ```
   C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\UK - Analytics\UK Team Reports\Reports\DailyReport\CF_Action\raw\
   ├── daily_email_previews\       (staff mailer HTML/PDF)
   ├── manager_digest_previews\    (manager emails)
   └── executive_digest_previews\  (executive emails)
   ```

## Troubleshooting

### Task Not Running

**Problem**: Task shows "Last Run Result: Task Disabled"

**Solution**:
1. Right-click task → "Properties"
2. Check "Run with highest privileges"
3. Set "Configure for: Windows Server 2016" (or latest available)
4. Click "OK" → "Enable task"

### Script Fails to Execute

**Problem**: Last Run Result shows error code

**Solution**:
1. Run script manually to test:
   ```powershell
   python "C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2\wd_summariser.py"
   ```
2. Check error output and verify:
   - Python path is correct
   - Database file exists at: `C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2\kc_reports.db`
   - All dependencies installed: `pip list | grep -E "pandas|openpyxl|sqlalchemy|pdfkit"`

### Emails Not Sending

**Problem**: Task runs successfully but emails not received

**Solution**:
1. Check sent_log_detail.csv for error messages
2. Verify `DRY_RUN = False` in email scripts
3. Check Outlook is running (if using Outlook send)
4. Test with `DRY_RUN = True` to preview emails first

### Database Errors

**Problem**: Task fails with "Database not found" or connection error

**Solution**:
1. Verify VPN is connected (if DB is remote)
2. Check network connectivity: `ping sql-web-edition.cqu5cuepe15y.ap-south-1.rds.amazonaws.com`
3. Verify ODBC Driver 17 installed: `odbcad32.exe` → Drivers tab
4. Test extraction manually:
   ```powershell
   python daily_wd_data_extractor.py
   ```

## Monitoring & Maintenance

### Weekly Checks

- [ ] Verify all 5 tasks completed successfully
- [ ] Check sent_log.csv for delivery status
- [ ] Review summary_all_days.xlsx for data quality

### Monthly Tasks

- [ ] Rotate database backups
- [ ] Review error logs and fix any issues
- [ ] Update .gitignore if new sensitive files appear
- [ ] Check for Python/library updates

### Quarterly Reviews

- [ ] Verify VPN connectivity settings
- [ ] Update email recipient lists in wd_settings.xlsx
- [ ] Test disaster recovery (database restore)
- [ ] Review Task Scheduler logs for any warnings

## Backup Strategy

### Database Backup

```powershell
# Create weekly backups (add as Task Scheduler task)
$date = Get-Date -Format "yyyy-MM-dd"
$source = "C:\Users\gsakhare\OneDrive - KC OVERSEAS EDUCATION PVT LTD\Python Scripts\DailyWD-UK - V1.2\kc_reports.db"
$backup = "C:\Backups\kc_reports_$date.db"
Copy-Item -Path $source -Destination $backup
```

### Settings Backup

Keep `wd_settings.xlsx` in OneDrive (auto-synced):
- Team roster
- Measure definitions
- Role configurations
- Holiday calendar

## Support & Contact

For issues or assistance:
- **Author**: Gaurav Mohan Sakhare
- **Email**: gsakhare@kcoverseas.com
- **Phone**: +91 712 2222061/62/63

---

**Last Updated**: January 2026  
**Version**: 3.10+
