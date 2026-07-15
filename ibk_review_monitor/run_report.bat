@echo off
cd /d "%~dp0"
echo Starting i-ONE Bank review monitor...
echo.

python review_monitor.py

echo.
echo Done. Please check the "reports" folder.
pause
