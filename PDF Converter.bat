@echo off
rem Launch the PDF Converter UI without a console window.
start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0app.py"
