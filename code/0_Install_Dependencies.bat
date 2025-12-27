@echo off
echo Installing requirements...
cd /d "%~dp0"
python -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo Error installing requirements! Please make sure Python is installed.
    pause
    exit /b
)
echo.
echo Requirements installed successfully!
pause
