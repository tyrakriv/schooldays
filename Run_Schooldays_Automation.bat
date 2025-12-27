@echo off
:MENU
cls
echo ------------------------------------------
echo    SCHOOL DAYS AUTOMATION MENU
echo ------------------------------------------
echo.
echo 1. Setup (Validate Data + Configure Screen)
echo 2. Run Automation
echo.
echo ------------------------------------------
set /p choice=Select an option (1 or 2): 

if "%choice%"=="1" goto SETUP
if "%choice%"=="2" goto RUN
goto MENU

:SETUP
cls
echo --- Step 1: Validating Excel Data ---
python code\validate_data.py
echo.
echo --- Step 2: Configuring Screen Coordinates ---
python code\config_wizard.py
pause
goto MENU

:RUN
cls
echo --- Starting Automation ---
python code\main.py
pause
goto MENU
