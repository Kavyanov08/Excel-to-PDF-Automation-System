@echo off
REM StartWatcher.bat - double-click to run the watcher without a console window
cd /d "%~dp0"
if exist ".venv\Scripts\pythonw.exe" (
    ".venv\Scripts\pythonw.exe" watch_folder.py
) else (
    pythonw.exe watch_folder.py
)