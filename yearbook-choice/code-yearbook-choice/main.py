import pyautogui
import json
import time
import os
import pyperclip
import pandas as pd
from datetime import datetime
from data_handler import load_and_process_data

COORD_FILE = os.path.join(os.path.dirname(__file__), "coordinates.json")

# Global timestamp for this run instance
SESSION_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

def load_coordinates():
    if not os.path.exists(COORD_FILE):
        print("Error: coordinates.json not found. Run config_wizard.py first!")
        return None
    with open(COORD_FILE, "r") as f:
        return json.load(f)

def log_runtime_error(student, reason):
    # Try to find the shared session file from validate_data.py
    session_info_path = os.path.join(os.path.dirname(__file__), "current_session.txt")
    filename = None
    
    if os.path.exists(session_info_path):
        try:
            with open(session_info_path, "r") as f:
                filename = f.read().strip()
        except:
            pass
            
    # Fallback if running standalone or read failed
    if not filename:
        reports_dir = "reports"
        if not os.path.exists(reports_dir):
            os.makedirs(reports_dir)
        filename = os.path.join(reports_dir, f"run-runtime-errors-{SESSION_TIMESTAMP}.csv")
    
    err_entry = student.copy()
    if 'error_reason' in err_entry:
        del err_entry['error_reason']
        
    err_entry['error_reason'] = reason
    
    # Create ordered dict/df (Natural order is fine since we just added error_reason last)
    df = pd.DataFrame([err_entry])
    header = not os.path.exists(filename)
    try:
        df.to_csv(filename, mode='a', header=header, index=False)
    except Exception as e:
        print(f"Failed to log runtime error: {e}")

def run_automation():
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

    # Wait a sec to switch focus
    print("Starting in 3 seconds...")
    time.sleep(3)

    # 0. INITIALIZATION: Ensure "Web Entry" is UNCHECKED (Reset State)
    # We do this once at the start to ensure we don't carry over manual checks
    if 'web_entry_input_box' in coords and 'web_entry_checkbox' in coords:
        pyautogui.click(coords['web_entry_input_box']['x'], coords['web_entry_input_box']['y'])
        time.sleep(.05)
        # Try to type
        pyperclip.copy("")
        pyautogui.doubleClick()
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(.05)
        
        #if no text is in the input, then we need to click the checkbox
        if pyperclip.paste().lower().strip() != '':
            pyautogui.doubleClick()
            time.sleep(.05)
            pyautogui.click(coords['web_entry_checkbox']['x'], coords['web_entry_checkbox']['y'])
            time.sleep(.05)
        else:
            pass
            
    # 0.5. INITIALIZATION: Ensure "Last Name" is UNCHECKED (Reset State)
    if 'last_name_box' in coords and 'last_name_checkbox' in coords:
        pyautogui.click(coords['last_name_box']['x'], coords['last_name_box']['y'])
        time.sleep(.05)
        # Try to type
        pyperclip.copy("")
        pyautogui.doubleClick()
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(.05)
        
        #if no text is in the input, then we need to click the checkbox
        if pyperclip.paste().lower().strip() != '':
             pyautogui.doubleClick()
             time.sleep(.05)
             pyautogui.click(coords['last_name_checkbox']['x'], coords['last_name_checkbox']['y'])
             time.sleep(.05)
        else:
            pass
    
    for i, student in enumerate(students):
        sid = student['id']
        selection = student['selection']
                
        # 1. Search
        pyautogui.click(coords['search_box']['x'], coords['search_box']['y'])
        # Double click to select all text (to overwrite previous)
        pyautogui.doubleClick() 
        pyautogui.typewrite(sid)
        pyautogui.press('enter') 
        
        # Wait for load (Very important for legacy apps)
        time.sleep(.05) 
        
        # 2. VALIDATION: Check Last Name
        if 'last_name_box' in coords:
            pyautogui.click(coords['last_name_box']['x'], coords['last_name_box']['y'])
            # Select All (Triple Click)
            pyautogui.tripleClick()
            time.sleep(.05)
            # Copy
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(.05)
            
            last_name = pyperclip.paste().strip()
            
            if not last_name:
                print(f"  -> VALIDATION FAILED: Student ID {sid} not found (Last Name empty). Skipping.")
                log_runtime_error(student, "Student ID not found (Empty Last Name)")
                continue
                
            # NEW: Check against Excel Last Name
            excel_last_name = student.get('last_name', '')
            if excel_last_name and last_name.lower() != excel_last_name.lower():
                 print(f"  -> VALIDATION FAILED: Last Name Mismatch! Found: '{last_name}', Expected: '{excel_last_name}'")
                 log_runtime_error(student, f"Last Name Mismatch (Found: {last_name}, Expected: {excel_last_name})")
                 continue
            
        else:
             print("  -> Warning: Skipping validation (last_name_box not configured).")
 
        
        # 2. Select Option
        if selection == 'd':
            #do nothing since d is default
            print(f"  -> DEFAULT SELECTION: Student ID {sid} selected default option (D).")
            pass
        elif selection == 'a':
            pyautogui.click(coords['option_a']['x'], coords['option_a']['y'])
            print(f"  -> SELECTION: Student ID {sid} selected option A.")
        elif selection == 'b':
            pyautogui.click(coords['option_b']['x'], coords['option_b']['y'])
            print(f"  -> SELECTION: Student ID {sid} selected option B.")
        elif selection == 'c':
            pyautogui.click(coords['option_c']['x'], coords['option_c']['y'])
            print(f"  -> SELECTION: Student ID {sid} selected option C.")
        else:
            print(f"  -> Unknown selection '{selection}'. Skipping.")
        
        # 3. Audit Trail (Check "Web Entry" and type "auto")
        if 'web_entry_input_box' in coords and 'web_entry_checkbox' in coords:
            # Step A: Try to type "auto" in source box assuming it's enabled
            pyautogui.click(coords['web_entry_input_box']['x'], coords['web_entry_input_box']['y'])
            time.sleep(.05)
            
            # Select All to overwrite (Clean entry)
            pyautogui.tripleClick()
            time.sleep(.05)
            
            # Type "auto"
            pyautogui.typewrite("auto")
            time.sleep(.05)
        
        else:
            print("  -> Warning: Skipping audit trail (web_entry_input_box or web_entry_checkbox not configured).")
        
        # Small pause between records
        time.sleep(0.1)

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
        sys.exit(1)
    except Exception as e:
        print(f"\nCRITICAL ERROR: {e}")
        sys.exit(1)
