import pyautogui
import json
import time
import os
import pyperclip
import pandas as pd
from datetime import datetime
from data_handler import load_and_process_data
from excel_utils import find_column_robust

COORD_FILE = os.path.join(os.path.dirname(__file__), "coordinates.json")
SESSION_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
_error_log = []
_completed_log = []

def load_coordinates():
    if not os.path.exists(COORD_FILE):
        print("Error: coordinates.json not found. Run config_wizard.py first!")
        return None
    with open(COORD_FILE, "r") as f:
        return json.load(f)

def log_runtime_error(student, reason):
    err_entry = student.copy()
    if "error_reason" in err_entry:
        del err_entry["error_reason"]
    err_entry["error_reason"] = reason
    _error_log.append(err_entry)

def _write_error_report():
    if not _error_log:
        return
    session_info_path = os.path.join(os.path.dirname(__file__), "current_session.txt")
    reports_dir = "reports"
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
    filename = os.path.join(reports_dir, f"yearbook-choice-errors-{SESSION_TIMESTAMP}.xlsx")
    if os.path.exists(session_info_path):
        try:
            with open(session_info_path, "r") as f:
                p = f.read().strip()
                if p and os.path.exists(p):
                    filename = p
        except OSError:
            pass
    try:
        if os.path.exists(filename):
            existing = pd.read_excel(filename)
            # Runtime errors only have id, last_name, selection. Map them into existing columns
            # so all rows share the same schema (no extra id/last_name/selection columns).
            id_col = find_column_robust(existing, "student id")
            ln_col = find_column_robust(existing, "student last name")
            sel_col = find_column_robust(existing, ["yearbook photo", "selection"])
            err_col = "error_reason" if "error_reason" in existing.columns else None
            new_rows = []
            for entry in _error_log:
                row = {c: None for c in existing.columns}
                if id_col:
                    row[id_col] = entry.get("id")
                if ln_col:
                    row[ln_col] = entry.get("last_name")
                if sel_col:
                    row[sel_col] = entry.get("selection")
                if err_col:
                    row[err_col] = entry.get("error_reason")
                new_rows.append(row)
            new_df = pd.DataFrame(new_rows, columns=existing.columns)
            new_df = new_df.astype(object)
            combined = pd.concat([existing, new_df], ignore_index=True)
            combined.to_excel(filename, index=False)
        else:
            # No existing file: write runtime errors with clear column names
            df = pd.DataFrame(_error_log)
            if "id" in df.columns:
                df = df.rename(columns={"id": "Student ID", "last_name": "Last Name", "selection": "Selection"})
            df.to_excel(filename, index=False)
        print(f"Error report saved: {filename}")
    except Exception as e:
        print(f"Failed to write error report: {e}")

def log_success(student):
    """Logs successfully completed students (one row per student we finished)."""
    _completed_log.append(student)

def _write_completed_report():
    if not _completed_log:
        return
    reports_dir = "reports"
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
    filename = os.path.join(reports_dir, f"yearbook-choice-completed-{SESSION_TIMESTAMP}.xlsx")
    try:
        pd.DataFrame(_completed_log).to_excel(filename, index=False)
        print(f"Completed list saved: {filename}")
    except Exception as e:
        print(f"Failed to write completed report: {e}")

def verify_field_is_editable(entry, field_name):
    pyautogui.click(entry['x'], entry['y'])
    time.sleep(.5)
    # Try to type
    pyperclip.copy("")
    pyautogui.doubleClick()
    pyautogui.hotkey('ctrl', 'c')
    initial_text = pyperclip.paste()

    checked = False
    if initial_text.lower().strip() in ('auto', 'pkg auto', 'yrbk auto'):
        checked = True
    else:
        pyautogui.doubleClick()
        pyautogui.typewrite("reset")
        time.sleep(.1)
        pyperclip.copy("")
        pyautogui.doubleClick()
        pyautogui.hotkey('ctrl', 'c')
        pasted_text = pyperclip.paste()
        pyautogui.doubleClick()
        if initial_text:
            pyautogui.typewrite(initial_text)
        else:
            pyautogui.press('backspace')
        if pasted_text.lower().strip() == 'reset':
            checked = True

    time.sleep(.1)

    if not checked:
        print(f"{field_name} Field is unchecked. Please fix and restart the program.")
        return False
    return True

def run_automation():
    global _error_log, _completed_log
    _error_log = []
    _completed_log = []
    coords = load_coordinates()
    if not coords:
        return False

    # User Safety Prompt
    print("--- READY TO START ---")
    print("1. Make sure the window is in the SAME position as when you ran the wizard.")
    print("2. Do NOT touch the mouse/keyboard once this starts.")
    print("3. Move your mouse to the corner of the screen to trigger a FAILSAFE abort.")
    print("----------------------")
    
    # data_handler will find the first .xlsx file automatically
    students = load_and_process_data(None) # Passing None as we updated logic to find file internally

    if not students:
        print("No student data found.")
        return False

    # Run summary: what we're about to enter (like package-choice)
    reports_dir = "reports"
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
    summary_rows = [{"Student ID": s["id"], "Last Name": s.get("last_name", ""), "Yearbook Selection": s.get("selection", "")} for s in students]
    if summary_rows:
        summary_file = os.path.join(reports_dir, f"yearbook-choice-run-summary-{SESSION_TIMESTAMP}.xlsx")
        pd.DataFrame(summary_rows).to_excel(summary_file, index=False)
        print(f"Run summary saved: {summary_file}")

    # Wait a sec to switch focus
    print("Starting in 3 seconds...")
    time.sleep(3)

    # 0. INITIALIZATION: Ensure "Web Entry" is UNCHECKED (Reset State)
    # We do this once at the start to ensure we don't carry over manual checks
    if 'web_entry_input_box' in coords:
        if not verify_field_is_editable(coords['web_entry_input_box'], "Web Entry (yrbk auto)"):
            return False
            
    # 0.5. INITIALIZATION: Ensure "Last Name" is UNCHECKED (Reset State)
    if 'last_name_box' in coords:
        if not verify_field_is_editable(coords['last_name_box'], "Last Name"):
            return False
    
    for student in students:
        sid = student['id']
        selection = student['selection']
        excel_last_name = student.get('last_name', '')
                
        # 1. Search
        pyautogui.click(coords['search_box']['x'], coords['search_box']['y'])
        pyautogui.doubleClick() 
        pyautogui.typewrite(sid)
        pyautogui.press('enter') 
        
        time.sleep(.1) 
        
        # 2. VALIDATION: Check Last Name
        if 'last_name_box' in coords:
            pyautogui.click(coords['last_name_box']['x'], coords['last_name_box']['y'])
            pyautogui.tripleClick()
            time.sleep(.1)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(.1)
            
            last_name = pyperclip.paste().strip()
            
            if not last_name:
                print(f"  -> VALIDATION FAILED: Student ID {sid} not found (Last Name empty). Skipping.")
                log_runtime_error(student, "Student ID not found (Empty Last Name)")
                continue
                
            if excel_last_name:
                # Handle hyphenated names (App might select "Walsh-" with trailing hyphen)
                expected_parts = excel_last_name.lower().split('-')
                first_part = expected_parts[0].strip().lower()
                first_part_with_hyphen = first_part + '-'
                
                # Allow match if found name is: full name, first part only, OR first part with hyphen
                is_match = (last_name.lower() == excel_last_name.lower()) or \
                           (last_name.lower() == first_part) or \
                           (last_name.lower() == first_part_with_hyphen)
                
                if not is_match:
                    print(f"  -> NAME MISMATCH: Found '{last_name}', Expected '{excel_last_name}'")
                    log_runtime_error(student, f"Last Name Mismatch (Found: {last_name}, Expected: {excel_last_name})")
                    continue
            
        else:
             pass  # Validation skipped if not configured
 
        # 2. Audit Trail (Web Entry -> "yrbk auto")
        if 'web_entry_input_box' in coords:
            pyautogui.click(coords['web_entry_input_box']['x'], coords['web_entry_input_box']['y'])
            time.sleep(.1)
            pyautogui.tripleClick()
            time.sleep(.1)
            pyautogui.typewrite("yrbk auto")
            time.sleep(.1)
        
        else:
            pass  # Audit trail skipped if not configured
        
        # 3. Select Option
        if selection == 'd':
            pyautogui.click(coords['option_d']['x'], coords['option_d']['y'])
        elif selection == 'a':
            pyautogui.click(coords['option_a']['x'], coords['option_a']['y'])
        elif selection == 'b':
            pyautogui.click(coords['option_b']['x'], coords['option_b']['y'])
        elif selection == 'c':
            pyautogui.click(coords['option_c']['x'], coords['option_c']['y'])
        else:
            print(f"  -> Unknown selection '{selection}'. Skipping.")
        
        # Small pause between records
        time.sleep(.1)
        
        # Log success
        log_success(student)

    _write_error_report()
    _write_completed_report()
    print("Automation Complete!")
    return True

if __name__ == "__main__":
    import sys
    try:
        success = run_automation()
        if success:
            sys.exit(0) # Success
        else:
            sys.exit(1) # Logic error or setup failure
    except (pyautogui.FailSafeException, KeyboardInterrupt):
        print("\n")
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print("   AUTOMATION ABORTED BY USER (FailSafe Triggered)")
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        _write_error_report()
        _write_completed_report()
        sys.exit(1)
    except Exception as e:
        print(f"\nCRITICAL ERROR: {e}")
        _write_error_report()
        _write_completed_report()
        sys.exit(1)
