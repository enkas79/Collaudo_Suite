@echo off
cd /d "%~dp0"
python run_suite.py
if errorlevel 1 pause
