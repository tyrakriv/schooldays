@echo off
cls
echo ---------------------------------------------------
echo    SCHOOL DAYS AUTOMATION (Seamless Mode)
echo ---------------------------------------------------
echo.

:: Step 1: Validation & Cleanup (Always Run)
echo --- Step 1: Validating ^& Cleaning Data ---
python code-yearbook-choice\validate_data.py

:: Check if validation failed (ErrorLevel 1 = Critital, ErrorLevel 0 = OK)
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Validation Failed. Please fix the Excel file and try again.
    pause
    exit /b
)

:: Step 1.5: Config (Always Run)
echo.
echo --- Step 2: Configuring Screen Coordinates ---
python code-yearbook-choice\config_wizard.py

:: Step 2: Automation
echo.
echo --- Step 3: Running Automation ---
python code-yearbook-choice\main.py

if %errorlevel% equ 0 (
    echo.
    echo ***************************************************
    echo       AUTOMATION COMPLETED SUCCESSFULLY!
    echo ***************************************************
    echo.
    echo [SUCCESS] Press any key to close...
    pause >nul
) else (
    echo.
    echo [ABORTED] Automation stopped with errors or was cancelled.
    pause
)
