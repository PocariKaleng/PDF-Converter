@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
    if errorlevel 1 goto failed
)
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
".venv\Scripts\python.exe" app.py --open
if errorlevel 1 goto failed
exit /b 0
:failed
echo Aplikasi gagal dijalankan. Pastikan Python 3.11+ dan Word desktop terpasang.
pause
exit /b 1
