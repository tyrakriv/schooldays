@echo off
cls
echo ---------------------------------------------------
echo    SCHOOL DAYS AUTOMATION (Package Entry Mode)
echo ---------------------------------------------------
echo.

:: Activate Venv
if exist venv\Scripts\activate.bat call venv\Scripts\activate.bat
if exist ..\venv\Scripts\activate.bat call ..\venv\Scripts\activate.bat
if exist .venv\Scripts\activate.bat call .venv\Scripts\activate.bat
if exist ..\.venv\Scripts\activate.bat call ..\.venv\Scripts\activate.bat

:: Step 1: Validation (Always Run)
echo --- Step 1: Validating Input Data ---
python code-package-choice\validate_package.py

:: Check if validation failed
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Validation Failed. Please fix the issue and try again.
    pause
    exit /b
)

:: Step 2: Config (Always Run)
echo.
echo --- Step 2: Configuring Screen Coordinates ---
python code-package-choice\config_wizard_package.py

:: Step 3: Automation
echo.
echo --- Step 3: Running Automation ---
python code-package-choice\main.py

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
