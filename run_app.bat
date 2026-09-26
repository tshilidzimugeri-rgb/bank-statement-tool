@echo off
REM Opens the bank statement upload page in your browser. Close this window to stop it.
cd /d "%~dp0"
".venv\Scripts\streamlit.exe" run "src\statement_tool\app.py" --server.address localhost --browser.gatherUsageStats false
