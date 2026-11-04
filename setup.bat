@echo off
rem One-time setup: create the virtual environment and install pinned dependencies.
cd /d "%~dp0"
python -m venv .venv || goto :error
".venv\Scripts\python.exe" -m pip install -r requirements.txt || goto :error
echo.
echo Setup complete. Double-click "PDF Converter.bat" to start.
pause
exit /b 0
:error
echo Setup failed. Python 3.10+ must be installed and on PATH.
pause
exit /b 1
