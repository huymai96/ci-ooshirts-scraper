@echo off
echo ============================================
echo Updating "Run Receiving Tool Scrapers" Task
echo - Change to HOURLY schedule
echo ============================================
echo.

REM Delete existing task
schtasks /Delete /TN "Run Receiving Tool Scrapers" /F 2>nul

REM Create new task with hourly schedule
schtasks /Create /TN "Run Receiving Tool Scrapers" /TR "\"%~dp0run_scrapers.bat\"" /SC HOURLY /MO 1 /ST 00:00 /RL HIGHEST /F

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [SUCCESS] Task updated to run EVERY HOUR!
    echo.
    echo Task Details:
    schtasks /Query /TN "Run Receiving Tool Scrapers" /V /FO LIST | findstr /I "Task\|Status\|Next\|Repeat"
) else (
    echo.
    echo [ERROR] Failed to create task. Make sure you're running as Administrator.
)

echo.
pause
