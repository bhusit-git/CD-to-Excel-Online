@echo off
setlocal

set "LAUNCHER_DIR=%~dp0"
set "SCRIPT=%LAUNCHER_DIR%c-d\build_full_billing_workbook.py"
if not exist "%SCRIPT%" set "SCRIPT=%LAUNCHER_DIR%build_full_billing_workbook.py"

if not exist "%SCRIPT%" (
  echo หา build_full_billing_workbook.py ไม่เจอ
  echo ให้วางไฟล์นี้ไว้ข้าง ๆ โฟลเดอร์ c-d หรือไว้โฟลเดอร์เดียวกับ script
  pause
  exit /b 1
)

set "INPUT_PATH=%~1"
set "MONTH_ARG=%~2"

if "%INPUT_PATH%"=="" (
  set /p INPUT_PATH=ใส่ path ของโฟลเดอร์ c-d: 
)

if "%MONTH_ARG%"=="" (
  set /p MONTH_ARG=ใส่เดือนข้อมูล (YYYY-MM) เช่น 2026-04 [ค่าเริ่มต้น 2026-04]: 
)

if exist "%MONTH_ARG%\" (
  set "INPUT_PATH=%MONTH_ARG%"
  set "MONTH_ARG="
)

if "%MONTH_ARG%"=="" set "MONTH_ARG=2026-04"

if not exist "%INPUT_PATH%\" (
  echo ไม่พบโฟลเดอร์: %INPUT_PATH%
  pause
  exit /b 1
)

python "%SCRIPT%" "%INPUT_PATH%" --month "%MONTH_ARG%"
if errorlevel 1 (
  echo.
  echo ถ้าเครื่องนี้ยังไม่มี dependency ให้รัน:
  echo pip install -r "%LAUNCHER_DIR%c-d\requirements.txt"
  pause
  exit /b 1
)

echo.
echo เสร็จแล้ว
pause
