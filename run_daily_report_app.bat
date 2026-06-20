@echo off
setlocal
cd /d "%~dp0"
if exist "runtime\pythonw.exe" (
  "runtime\pythonw.exe" "scripts\daily_report_app.py"
  exit /b %errorlevel%
)
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" "scripts\daily_report_app.py"
  exit /b %errorlevel%
)
if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
  "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" "scripts\daily_report_app.py"
  exit /b %errorlevel%
)
where py >nul 2>nul
if %errorlevel% equ 0 (
  py -3 "scripts\daily_report_app.py"
  exit /b %errorlevel%
)
where python >nul 2>nul
if %errorlevel% equ 0 (
  python "scripts\daily_report_app.py"
  exit /b %errorlevel%
)
echo Python runtime not found. Contact system administrator.
pause
exit /b 1
