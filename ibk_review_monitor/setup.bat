@echo off
cd /d "%~dp0"
echo ===================================================
echo  i-ONE Bank Review Monitor - First-time Setup
echo ===================================================
echo.
echo Installing required packages, please wait...
echo.

python --version >nul 2>&1
if errorlevel 1 goto nopython

python -m pip install --upgrade pip -q
python -m pip install -r requirements.txt

echo.
echo ===================================================
echo  Setup complete!
echo  Now double-click run_report.bat to generate a report.
echo ===================================================
pause
exit /b 0

:nopython
echo [ERROR] Python is not installed, or not added to PATH.
echo Please see the "Python install" section in README.md first.
pause
exit /b 1
