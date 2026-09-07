@echo off
cd /d "%~dp0"
.venv\Scripts\python.exe -m uvicorn app:app --host 0.0.0.0 --port 8443 --ssl-keyfile=key.pem --ssl-certfile=cert.pem
pause