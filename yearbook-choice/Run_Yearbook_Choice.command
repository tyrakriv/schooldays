#!/bin/bash
cd "$(dirname "$0")"

echo "---------------------------------------------------"
echo "   SCHOOL DAYS AUTOMATION (Yearbook Photo Choice)"
echo "---------------------------------------------------"
echo ""

# Activate venv if present
if [ -d "venv" ]; then
    source venv/bin/activate
elif [ -d "../venv" ]; then
    source ../venv/bin/activate
elif [ -d "code-yearbook-choice/venv" ]; then
    source code-yearbook-choice/venv/bin/activate
fi

# Install dependencies if pyautogui missing (needed for config + automation)
if ! python3 -c "import pyautogui" 2>/dev/null; then
    echo "--- Installing dependencies (one-time) ---"
    pip3 install -r code-yearbook-choice/requirements.txt
    if [ $? -ne 0 ]; then
        echo "[ERROR] Failed to install dependencies. Run: pip3 install -r code-yearbook-choice/requirements.txt"
        read -p "Press Enter to exit..."
        exit 1
    fi
    echo ""
fi

# Step 1: Validation & Cleanup
echo "--- Step 1: Validating & Cleaning Data ---"
python3 code-yearbook-choice/validate_data.py
if [ $? -ne 0 ]; then
    echo ""
    echo "[ERROR] Validation Failed. Please fix the Excel file and try again."
    read -p "Press Enter to exit..."
    exit 1
fi

# Step 2: Config
echo ""
echo "--- Step 2: Configuring Screen Coordinates ---"
python3 code-yearbook-choice/config_wizard.py

# Step 3: Automation
echo ""
echo "--- Step 3: Running Automation ---"
python3 code-yearbook-choice/main.py
if [ $? -eq 0 ]; then
    sleep 1
    echo ""
    echo "---------------------------------------------------"
    echo "Done! Press [ENTER] to close this window."
    read
    osascript -e 'tell application "Terminal" to close first window' &
    exit
else
    sleep 1
    echo ""
    echo "---------------------------------------------------"
    echo "[STOPPED] The process was stopped or encountered an error."
    echo "Check the messages above for details."
    echo ""
    echo "Press [ENTER] to close this window."
    read
    osascript -e 'tell application "Terminal" to close first window' &
    exit
fi
