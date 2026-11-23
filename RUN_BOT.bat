@echo off
cd /d "%~dp0"
python main.py
if errorlevel 1 (
  echo.
  echo The assistant stopped with an error. Read the message above.
)
pause
