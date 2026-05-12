@echo off
TITLE KC Analytics - Management Dashboard (DEBUG MODE)
SETLOCAL EnableDelayedExpansion

echo ============================================================
echo      KC ANALYTICS - SYSTEM STARTUP (DEBUG)
echo ============================================================

:: 1. Check Python
echo [1/5] Checking Python...
python --version
if %errorlevel% neq 0 (
    echo [ERROR] Python not found.
    pause
    exit /b
)

:: 2. Dependencies
echo [2/5] Installing/Updating dependencies...
python -m pip install fastapi uvicorn pandas openpyxl sqlalchemy pydantic
if %errorlevel% neq 0 (
    echo [ERROR] Python pip install failed.
    pause
    exit /b
)

:: 3. Database
echo [3/5] Initializing Database...
python migrate_settings_to_db.py
if %errorlevel% neq 0 (
    echo [ERROR] migrate_settings_to_db.py failed.
    pause
    exit /b
)

:: 4. Node.js / npm
echo [4/5] Checking Node.js...
node -v
if %errorlevel% neq 0 (
    echo [ERROR] Node.js not found.
    pause
    exit /b
)

:: 5. Frontend Setup
echo [5/5] Checking Frontend...
if not exist "frontend\package.json" (
    echo [ERROR] frontend\package.json not found! Are you in the right folder?
    dir
    pause
    exit /b
)

cd frontend
if not exist "node_modules" (
    echo [ACTION] node_modules missing. Running npm install...
    call npm.cmd install
    if %errorlevel% neq 0 (
        echo [ERROR] npm install failed.
        pause
        exit /b
    )
)

:: 6. Launch
echo.
echo Launching Backend (in new window)...
start "KC API Backend" cmd /k "cd .. && python api_server.py"

echo Launching Frontend...
echo.
echo ------------------------------------------------------------
echo    DASHBOARD SHOULD BE AT:  http://localhost:3000
echo    IF THIS WINDOW CLOSES, CHECK FOR ERRORS ABOVE.
echo ------------------------------------------------------------
echo.
call npm.cmd run dev

if %errorlevel% neq 0 (
    echo [ERROR] Frontend (npm run dev) failed.
    pause
)

:: Keep window open if something goes wrong
echo.
echo [SYSTEM] Process ended or interrupted.
pause
cmd /k
