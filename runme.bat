@echo off
echo ========================================
echo Shipment Validation - Web Database Mode
echo ========================================
echo.
echo Starting web-based shipment validator...
echo This will start the web UI server Locally.
echo.
pause
python generate_daily_report.py
pause 