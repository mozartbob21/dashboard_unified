@echo off
chcp 65001 >nul
cd /d "%~dp0\..\.."
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" -m services.edds.chrome --setup
) else (
  python -m services.edds.chrome --setup
)
pause
