@echo off
rem Starts Offload on this computer only (nothing is exposed to the network).
cd /d "%~dp0"
set FLASK_DEBUG=0
set HOST=127.0.0.1
set PORT=5000
echo Starting Offload at http://localhost:5000  (close this window to stop it)
start "" cmd /c "timeout /t 4 >nul & start http://localhost:5000"
py app.py
pause
