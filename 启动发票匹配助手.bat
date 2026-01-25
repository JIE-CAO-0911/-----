@echo off
setlocal

cd /d "%~dp0"

set PORT=8501

start "" /min python -m streamlit run "streamlit_app.py" --server.headless true --server.port %PORT% --server.fileWatcherType none
start "" "http://localhost:%PORT%"

endlocal
