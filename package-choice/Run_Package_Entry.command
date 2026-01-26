#!/bin/bash
cd "$(dirname "$0")"

echo "---------------------------------------------------"
echo "   SCHOOL DAYS AUTOMATION (Package Entry Mode)"
echo "---------------------------------------------------"
echo ""

# Activate Venv
if [ -d "venv" ]; then
    source venv/bin/activate
elif [ -d "../venv" ]; then
    source ../venv/bin/activate
elif [ -d ".venv" ]; then
    source .venv/bin/activate
elif [ -d "../.venv" ]; then
    source ../.venv/bin/activate
fi

# Step 1: Validation
echo "--- Step 1: Validating Input Data ---"
python3 code-package-choice/validate_package.py
if [ $? -ne 0 ]; then
    echo ""
    echo "[ERROR] Validation Failed. Please fix the issue and try again."
    read -p "Press Enter to exit..."
    exit 1
fi

# Step 2: Config
echo ""
echo "--- Step 2: Configuring Screen Coordinates ---"
python3 code-package-choice/config_wizard_package.py

# Step 3: Automation
echo ""
echo "--- Step 3: Running Automation ---"
python3 code-package-choice/main_package.py
if [ $? -eq 0 ]; then
    echo ""
    echo "***************************************************"
    echo "      AUTOMATION COMPLETED SUCCESSFULLY!"
    echo "***************************************************"
    read -p "Press Enter to close..."
else
    echo ""
    echo "[ABORTED] Automation stopped with errors or was cancelled."
    read -p "Press Enter to close..."
fi
