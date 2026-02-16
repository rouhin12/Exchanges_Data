@echo off
REM Daily Exchange Data Download Scheduler - Windows Setup
REM This batch file helps you set up the scheduler in Windows Task Scheduler

echo ===================================================================
echo Exchange Data Download Scheduler - Windows Setup
echo ===================================================================
echo.

REM Get current directory and Python path
set SCRIPT_DIR=%~dp0
set PYTHON_EXE=%SCRIPT_DIR%.venv\Scripts\python.exe
set SCHEDULER_SCRIPT=%SCRIPT_DIR%scheduler.py

REM Check if Python exists
if not exist "%PYTHON_EXE%" (
    echo ERROR: Python virtual environment not found at: %PYTHON_EXE%
    echo Please activate your virtual environment first.
    pause
    exit /b 1
)

REM Check if schedule package is installed
echo Checking dependencies...
"%PYTHON_EXE%" -c "import schedule" 2>nul
if errorlevel 1 (
    echo.
    echo Installing required package: schedule
    "%PYTHON_EXE%" -m pip install schedule
    if errorlevel 1 (
        echo ERROR: Failed to install schedule package
        pause
        exit /b 1
    )
)

echo.
echo ===================================================================
echo Setup Complete!
echo ===================================================================
echo.
echo Choose an option:
echo.
echo 1. Test run (download all data now)
echo 2. Test NSE only
echo 3. Create Windows Task Scheduler entry (manual)
echo 4. Run scheduler in foreground (Ctrl+C to stop)
echo 5. Exit
echo.

set /p choice="Enter choice (1-5): "

if "%choice%"=="1" (
    echo.
    echo Running test download...
    "%PYTHON_EXE%" "%SCHEDULER_SCRIPT%" --test
    pause
    goto :eof
)

if "%choice%"=="2" (
    echo.
    echo Running NSE-only test...
    "%PYTHON_EXE%" "%SCHEDULER_SCRIPT%" --test --nse-only
    pause
    goto :eof
)

if "%choice%"=="3" (
    echo.
    echo ===================================================================
    echo Windows Task Scheduler Setup Instructions
    echo ===================================================================
    echo.
    echo 1. Open Task Scheduler (Win+R, type: taskschd.msc)
    echo 2. Click "Create Basic Task"
    echo 3. Name: Exchange Data Daily Download
    echo 4. Trigger: Daily at 9:00 AM
    echo 5. Action: Start a program
    echo 6. Program: %PYTHON_EXE%
    echo 7. Arguments: "%SCHEDULER_SCRIPT%"
    echo 8. Start in: %SCRIPT_DIR%
    echo.
    echo Press any key to copy paths to clipboard (optional)...
    pause >nul
    echo.
    echo Python path: %PYTHON_EXE%
    echo Script path: %SCHEDULER_SCRIPT%
    echo Working dir: %SCRIPT_DIR%
    echo.
    pause
    goto :eof
)

if "%choice%"=="4" (
    echo.
    echo Starting scheduler in foreground...
    echo Press Ctrl+C to stop.
    echo.
    "%PYTHON_EXE%" "%SCHEDULER_SCRIPT%"
    goto :eof
)

if "%choice%"=="5" (
    exit /b 0
)

echo Invalid choice. Please run again.
pause
