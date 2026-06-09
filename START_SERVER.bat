@echo off
title VOLTGRID Energy Dashboard
cd /d "%~dp0"

echo ===================================================
echo    VOLTGRID  -  Energy Monitoring Dashboard
echo ===================================================
echo.

REM --- check Python is installed ---
python --version >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Python was not found.
  echo.
  echo Please install Python from:  https://www.python.org/downloads/
  echo During install, TICK the box:  "Add Python to PATH"
  echo Then run this file again.
  echo.
  pause
  exit /b
)

echo Step 1 of 2:  Installing required libraries ^(first run only^)...
echo.
python -m pip install -r requirements.txt
echo.

echo Step 2 of 2:  Starting the dashboard...
echo.
echo    Open this address in your browser:   http://localhost:5010
echo    Login:   admin   /   admin123
echo.
echo    KEEP THIS WINDOW OPEN while using the dashboard.
echo    Close it ^(or press Ctrl+C^) to stop the server.
echo.
echo ---------------------------------------------------
echo.

python app.py

echo.
echo ===================================================
echo  The server has stopped. Read any message above.
echo ===================================================
pause
