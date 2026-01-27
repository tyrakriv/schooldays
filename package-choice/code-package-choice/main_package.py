import pyautogui
import json
import time
import os
import pyperclip
import pandas as pd
from datetime import datetime
from data_handler_package import load_and_process_data

COORD_FILE = os.path.join(os.path.dirname(__file__), "coordinates_package.json")
SESSION_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

def load_coordinates():
    if not os.path.exists(COORD_FILE):
        print("Error: coordinates_package.json not found. Run config_wizard_package.py first!")
        return None
    with open(COORD_FILE, "r") as f:
        return json.load(f)

def log_error(student_id, last_name, product_raw, reason):
    reports_dir = "reports"
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
    filename = os.path.join(reports_dir, f"package-errors-{SESSION_TIMESTAMP}.csv")
    
    entry = {
        "student_id": student_id,
        "last_name": last_name,
        "product_raw": product_raw,
        "error_reason": reason,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    df = pd.DataFrame([entry])
    header = not os.path.exists(filename)
    try:
        df.to_csv(filename, mode='a', header=header, index=False)
    except Exception as e:
        print(f"Failed to log error: {e}")

def click_and_type(coord, text, press_enter=True):
    if not coord:
        return
    pyautogui.click(coord['x'], coord['y'])
    time.sleep(0.05)
    
    # Select all to ensure clear or just click? 
    # Usually safer to just type for "Quick Entry" boxes if they clear themselves, 
    # but if they don't, we might append.
    # User screenshot shows "Quick Package Entry" at bottom. It likely clears after Enter.
    # But "Class Pkg" might differ.
    # I'll double click to be safe if it's text.
    pyautogui.doubleClick() 
    time.sleep(0.05)
    
    pyautogui.typewrite(str(text))
    time.sleep(0.05)
    
    if press_enter:
        pyautogui.press('enter')
        time.sleep(0.1) # Wait for add

def run_automation():
    coords = load_coordinates()
    if not coords:
        return False

    pyautogui.FAILSAFE = True # Enabled by default, but making it explicit
    
    print("--- READY TO START PACKAGE ENTRY ---")
    print("1. Ensure School Days app is open and ready.")
    print("2. EMERGENCY STOP: Slam mouse quickly to any corner of the screen.")
    print("3. OR click on this Terminal window and press Ctrl+C.")
    print("------------------------------------")
    
    students = load_and_process_data(None) # Auto-finds Excel

    if not students:
        print("No student data found or processed.")
        return False
        
    print(f"Loaded {len(students)} students to process.")
    print("Starting in 3 seconds...")
    time.sleep(3)
    
    for student in students:
        sid = student['id']
        lname = student['last_name']
        choice_groups = student.get('choices_groups', [])
        errors = student.get('errors', [])
        
        print(f"Processing Student ID: {sid} ({len(choice_groups)} choice groups, {len(errors)} previous errors)")
        
        # 0. Log pre-existing errors (from data_handler logic)
        for err in errors:
            print(f"  -> Skipping invalid item: {err['raw_product']} ({err['reason']})")
            log_error(sid, lname, err['raw_product'], err['reason'])
            
        if not choice_groups:
            # If no valid groups to process, skip automation for this student
            continue
        
def search_student(sid, coords):
    if 'search_box' in coords:
        pyautogui.click(coords['search_box']['x'], coords['search_box']['y'])
        pyautogui.doubleClick() 
        pyautogui.typewrite(sid)
        pyautogui.press('enter')
        time.sleep(0.5) # Wait for load

def read_field_text(coord):
    """
    Clicks field, Selects All, Copies to clipboard, returns text.
    """
    if not coord: return ""
    
    # Click and focus
    pyautogui.click(coord['x'], coord['y'])
    
    # Select All (Cmd+A for Mac, Ctrl+A for Windows - user is on Mac but might be RDP?)
    # User OS is Mac. Standard is Command+A (Meta+A).
    # But often RDP/VMs use Ctrl. I'll stick to what they used before or safer: Triple Click.
    pyautogui.tripleClick()
    time.sleep(0.05)
    
    # Clear clipboard first
    pyperclip.copy("")
    
    # Copy
    pyautogui.hotkey('ctrl', 'c') 
    time.sleep(0.1)
    
    return pyperclip.paste().strip()

def run_automation():
    coords = load_coordinates()
    if not coords:
        return False

    pyautogui.FAILSAFE = True # Enabled by default, but making it explicit
    
    print("--- READY TO START PACKAGE ENTRY ---")
    print("1. Ensure School Days app is open and ready.")
    print("2. EMERGENCY STOP: Slam mouse quickly to any corner of the screen.")
    print("3. OR click on this Terminal window and press Ctrl+C.")
    print("------------------------------------")
    
    students = load_and_process_data(None) # Auto-finds Excel

    if not students:
        print("No student data found or processed.")
        return False
        
    print(f"Loaded {len(students)} students to process.")
    print("Starting in 3 seconds...")
    time.sleep(3)
    
    validated_first_student = False

    for student in students:
        sid = student['id']
        lname = student['last_name']
        choice_groups = student.get('choices_groups', [])
        errors = student.get('errors', [])
        
        print(f"Processing Student ID: {sid} ({len(choice_groups)} choice groups, {len(errors)} previous errors)")
        
        # 0. Log pre-existing errors (from data_handler logic)
        for err in errors:
            print(f"  -> Skipping invalid item: {err['raw_product']} ({err['reason']})")
            log_error(sid, lname, err['raw_product'], err['reason'])
            
        if not choice_groups:
            # If no valid groups to process, skip automation for this student
            continue
        
        # 1. Search Student
        search_student(sid, coords)
        
        # 2. Validate Last Name (Optional)
        if 'last_name_box' in coords and lname:
            pyautogui.click(coords['last_name_box']['x'], coords['last_name_box']['y'])
            pyautogui.tripleClick()
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.05)
            found_name = pyperclip.paste().strip()
            
            if found_name.lower() != lname.lower():
                print(f"  -> NAME MISMATCH: Found '{found_name}', Expected '{lname}'. Logging error.")
                log_error(sid, lname, "ALL", f"Name Mismatch (Found: {found_name})")
                continue # Skip this student

        # Track what we entered for validation
        entry_for_validation = None

        # 3. Process Each Yearbook Choice Group
        for group in choice_groups:
            photo_choice = group['photo_choice']
            standard_string = group['standard_string']
            other_items = group['others']
            
            print(f"  -> Processing Choice '{photo_choice}'...")

            # A. Click Photo Choice Letter (Once per group)
            choice_key = f"choice_{photo_choice}"
            if choice_key in coords:
                pyautogui.click(coords[choice_key]['x'], coords[choice_key]['y'])
                time.sleep(0.1)
            else:
                print(f"  -> Warning: Coordinate for choice '{photo_choice}' not found.")
                
            # B. Input Standard Packages (The combined string, e.g. "xxyy")
            if standard_string:
                print(f"     -> Entering Standard Packages: {standard_string}")
                if 'quick_package_entry_box' in coords:
                    click_and_type(coords['quick_package_entry_box'], standard_string)
                    
                    # Capture for validation if not yet validated
                    if not validated_first_student:
                        entry_for_validation = standard_string
                else:
                    print("     -> Error: 'quick_package_entry_box' coordinate missing.")
            
            # C. Input Other Items (Group, CD, Touchup) - processed individually
            for item in other_items:
                p_code = item['code']
                target_box_name = item['target_box']
                
                if target_box_name:
                    if target_box_name == 'touchup':
                        # Special interaction
                        if 'touchup_dropdown' in coords:
                            pyautogui.click(coords['touchup_dropdown']['x'], coords['touchup_dropdown']['y'])
                            time.sleep(0.2)
                            if 'touchup_pending_option' in coords:
                                pyautogui.click(coords['touchup_pending_option']['x'], coords['touchup_pending_option']['y'])
                                time.sleep(0.1)
                    
                    elif target_box_name in coords:
                         # click_and_type handles double click + enter
                         click_and_type(coords[target_box_name], p_code)
                    else:
                        print(f"     -> Error: Coordinate '{target_box_name}' not defined.")
                        log_error(sid, lname, item['raw_product'], f"Missing Coordinate: {target_box_name}")

        # 4. Perform Validation Check (Only for the first successful standard entry)
        if not validated_first_student and entry_for_validation:
            print("\n*** PERFORMING FIRST-RUN VALIDATION ***")
            print(f"Verifying package entry '{entry_for_validation}' for Student {sid}...")
            
            # A. Re-Search Student (to refresh view)
            search_student(sid, coords)
            
            # B. Check the box
            found_pkg = read_field_text(coords.get('quick_package_entry_box'))
            
            print(f"Expected: '{entry_for_validation}'")
            print(f"Found:    '{found_pkg}'")
            
            if found_pkg.lower() == entry_for_validation.lower():
                print("VALIDATION SUCCESSFUL! Continuing automation...")
                validated_first_student = True
            else:
                print("VALIDATION FAILED! The value in the box does not match what we typed.")
                print("Aborting to prevent errors.")
                return False

        time.sleep(0.1) # Pause between students

    print("Automation Complete!")
    return True

if __name__ == "__main__":
    import sys
    try:
        run_automation()
    except KeyboardInterrupt:
        print("\nABORTED BY USER.")
    except Exception as e:
        print(f"\nCRITICAL ERROR: {e}")
