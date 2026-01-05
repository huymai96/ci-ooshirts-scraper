@echo off
setlocal enableextensions

REM === SETTINGS ===
REM Uses current directory - works when cloned anywhere
set "BASE=%~dp0"
set "PY=python"
set "XLS=%BASE%customink_orders.xlsx"
set "BATCH_LOG=%BASE%batch_run.log"

REM === PREP: run from the scripts folder ===
pushd "%BASE%"

REM === Log start time ===
echo ============================================ >> "%BATCH_LOG%"
echo [%date% %time%] HOURLY SCRAPER RUN STARTED >> "%BATCH_LOG%"
echo ============================================ >> "%BATCH_LOG%"

REM === 1) DELETE THE EXCEL FILE FIRST (fresh run) ===
if exist "%XLS%" (
  del /f /q "%XLS%"
  echo [%date% %time%] Deleted old Excel file >> "%BATCH_LOG%"
) else (
  echo [%date% %time%] No existing Excel file >> "%BATCH_LOG%"
)

REM === 2) RUN OOSHIRT SCRAPER (HEADLESS - no prompts) ===
echo [%date% %time%] Starting ooshirts_order_scraper.py ... >> "%BATCH_LOG%"
"%PY%" "%BASE%ooshirts_order_scraper.py"
if %ERRORLEVEL% NEQ 0 (
  echo [%date% %time%] ERROR: Ooshirts scraper failed with code %ERRORLEVEL% >> "%BATCH_LOG%"
) else (
  echo [%date% %time%] Ooshirts scraper completed >> "%BATCH_LOG%"
)

REM === 3) Small wait ensures file handles are closed ===
powershell -NoProfile -Command "Start-Sleep -Seconds 3"

REM === 4) RUN CI SCRAPER (HEADLESS - no prompts) ===
echo [%date% %time%] Starting CI_order_scraper.py ... >> "%BATCH_LOG%"
"%PY%" "%BASE%CI_order_scraper.py"
if %ERRORLEVEL% NEQ 0 (
  echo [%date% %time%] ERROR: CI scraper failed with code %ERRORLEVEL% >> "%BATCH_LOG%"
) else (
  echo [%date% %time%] CI scraper completed >> "%BATCH_LOG%"
)

REM === 5) UPLOAD INBOUND.CSV TO CLOUD ===
echo [%date% %time%] Starting inbound.csv upload ... >> "%BATCH_LOG%"
"%PY%" "%BASE%upload_inbound.py"
if %ERRORLEVEL% NEQ 0 (
  echo [%date% %time%] ERROR: Inbound upload failed with code %ERRORLEVEL% >> "%BATCH_LOG%"
) else (
  echo [%date% %time%] Inbound upload completed >> "%BATCH_LOG%"
)

echo [%date% %time%] ALL TASKS COMPLETED >> "%BATCH_LOG%"
echo. >> "%BATCH_LOG%"
popd
exit /b 0
