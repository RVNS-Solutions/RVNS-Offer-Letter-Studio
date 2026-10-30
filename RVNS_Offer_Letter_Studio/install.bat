@echo off
cd /d "%~dp0"
py -3 -m venv .venv
if errorlevel 1 (
  echo Python 3 was not found. Install Python 3.10 or newer and enable Add Python to PATH.
  pause
  exit /b 1
)
call .venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Dependency installation failed. Check your internet connection and Python version.
  pause
  exit /b 1
)
echo.
echo Installation completed. Double-click run_dashboard.bat to start.
pause
