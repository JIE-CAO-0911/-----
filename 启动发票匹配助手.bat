@echo off
setlocal

cd /d "%~dp0"

set PORT=8501

for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":%PORT% " ^| findstr LISTENING') do (
    taskkill /f /pid %%a >nul 2>nul
)

start "" /min python -m streamlit run "streamlit_app.py" --server.headless true --server.port %PORT% --server.fileWatcherType poll --server.disconnectedSessionTTL 1
start "" "http://localhost:%PORT%"

endlocal
