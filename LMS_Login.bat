@echo off
cd /d "%~dp0"

echo Activating virtual environment...
call venv\Scripts\activate

echo Starting Leave Management System...
python app.py

echo App stopped. Closing...
pause
