@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\streamlit.exe" (
  echo Dependencies are not installed. Please run install.bat first.
  pause
  exit /b 1
)
call .venv\Scripts\activate
python -m streamlit run app.py
pause
