import pyautogui
import json
import time
import os
import pyperclip
import pandas as pd
from datetime import datetime
from data_handler import load_and_process_data

COORD_FILE = os.path.join(os.path.dirname(__file__), "coordinates.json")

def load_coordinates():
    if not os.path.exists(COORD_FILE):
        print("Error: coordinates.json not found. Run config_wizard.py first!")
        return None
    with open(COORD_FILE, "r") as f:
        return json.load(f)

def log_runtime_error(student, reason):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    reports_dir = "reports"
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
    
    filename = os.path.join(reports_dir, f"run-validation-{timestamp}.csv")
    
    err_entry = student.copy()
    if 'error_reason' in err_entry:
        del err_entry['error_reason']
        
    err_entry['error_reason'] = reason
    
    # Create ordered dict/df (Natural order is fine since we just added error_reason last)
    df = pd.DataFrame([err_entry])
    header = not os.path.exists(filename)
    try:
        df.to_csv(filename, mode='a', header=header, index=False)
        print(f"Logged error for {student['id']}: {reason}")
    except Exception as e:
        print(f"Failed to log runtime error: {e}")

def run_automation():
    coords = load_coordinates()
    if not coords:
        return

    # User Safety Prompt
    print("--- READY TO START ---")
    print("1. Open 'School Days Plus'.")
    print("2. Make sure the window is in the SAME position as when you ran the wizard.")
    print("3. Do NOT touch the mouse/keyboard once this starts.")
    print("4. Move your mouse to the corner of the screen to trigger a FAILSAFE abort.")
    print("----------------------")
    input("Press Enter to begin...")
    
    print("Loading Excel data from current folder...")
    # data_handler will find the first .xlsx file automatically
    students = load_and_process_data(None) # Passing None as we updated logic to find file internally

    if not students:
        print("No student data found.")
        return

    # Wait a sec to switch focus
    print("Starting in 3 seconds...")
    time.sleep(3)

    for i, student in enumerate(students):
        sid = student['id']
        selection = student['selection']
        
        print(f"[{i+1}/{len(students)}] Processing ID: {sid}, Selection: {selection}")
        
        # 1. Search
        pyautogui.click(coords['search_box']['x'], coords['search_box']['y'])
        # Double click to select all text (to overwrite previous)
        pyautogui.doubleClick() 
        pyautogui.typewrite(sid)
        pyautogui.press('enter') 
        
        # Wait for load (Very important for legacy apps)
        time.sleep(1.0) 
        
        # 2. VALIDATION: Check Last Name
        if 'last_name_box' in coords:
            pyautogui.click(coords['last_name_box']['x'], coords['last_name_box']['y'])
            # Select All (Double Click)
            pyautogui.doubleClick()
            time.sleep(0.1)
            # Copy
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.1)
            
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
            
            print(f"  -> Validation Passed: Found Last Name '{last_name}'")
        else:
             print("  -> Warning: Skipping validation (last_name_box not configured).")
 
        
        # 2. Select Option
        if selection == 'd':
            print("  -> Selection is 'd' (Default). Skipping clicks.")
        elif selection == 'a':
            pyautogui.click(coords['option_a']['x'], coords['option_a']['y'])
        elif selection == 'b':
            pyautogui.click(coords['option_b']['x'], coords['option_b']['y'])
        elif selection == 'c':
            pyautogui.click(coords['option_c']['x'], coords['option_c']['y'])
        else:
            print(f"  -> Unknown selection '{selection}'. Skipping.")
        
        # 3. Save (if needed)
        # Uncomment if a save click is required
        # pyautogui.click(coords['save_button']['x'], coords['save_button']['y'])
        
        # Small pause between records
        time.sleep(0.5)

    print("Automation Complete!")

if __name__ == "__main__":
    run_automation()
