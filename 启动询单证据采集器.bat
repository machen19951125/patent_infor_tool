@echo off
cd /d "%~dp0"
python "scripts\run_inquiry_evidence_app.py"
if errorlevel 1 pause
