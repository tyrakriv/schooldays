@echo off

echo Starting Configuration Wizard...
python code-yearbook-choice\validate_data.py
python code-yearbook-choice\config_wizard.py
pause
