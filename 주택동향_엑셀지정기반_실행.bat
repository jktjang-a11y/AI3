@echo off
chcp 65001 >nul
cd /d "%~dp0"
python housing_report_excel_upload.py
if errorlevel 1 pause
