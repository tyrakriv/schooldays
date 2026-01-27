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

def click_and_type(coord, text):
    if not coord:
        return
    pyautogui.click(coord['x'], coord['y'])
    pyautogui.tripleClick()
    time.sleep(0.05)
    
    pyautogui.typewrite(str(text))
    time.sleep(0.05)

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
        time.sleep(1.0) # Wait for add

def run_automation():
    coords = load_coordinates()
    if not coords:
        return False

    pyautogui.FAILSAFE = True # Enabled by default, but making it explicit
    
    print("--- READY TO START PACKAGE ENTRY ---")
    print("1. Ensure School Days app is open and ready.")
    print("2. IMPORTANT: Manually CHECK all input boxes (like Touchup) so they are editable!")
    print("3. EMERGENCY STOP: Slam mouse quickly to any corner of the screen.")
    print("4. OR click on this Terminal window and press Ctrl+C.")
    print("------------------------------------")
    
    students = load_and_process_data(None) # Auto-finds Excel

    if not students:
        print("No student data found or processed.")
        return False
        
    print(f"Loaded {len(students)} students to process.")
    
    # Generate Verification Excel
    print("Generating Verification Report...")
    reports_dir = "reports"
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
        
    verif_data = []
    for s in students:
        sid = s['id']
        lname = s['last_name']
        for grp in s.get('choices_groups', []):
            # Flatten for report - exclude error items like "Lost Order Form"
            valid_others = [x for x in grp['others'] 
                          if 'lost order' not in x['raw_product'].lower() 
                          and 'invalid' not in x['raw_product'].lower()]
            other_str = "; ".join([f"{x['raw_product']} (Input '{x['code']}')" for x in valid_others])
            verif_data.append({
                'Student ID': sid,
                'Last Name': lname,
                'Photo Choice': grp['photo_choice'] if grp['photo_choice'] else "(NONE)",
                'Standard Pkg String': grp['standard_string'],
                'Other Items': other_str
            })
            
    if verif_data:
        v_df = pd.DataFrame(verif_data)
        v_file = os.path.join(reports_dir, f"students_package_choices_verification-{SESSION_TIMESTAMP}.xlsx")
        v_df.to_excel(v_file, index=False)
        print(f"saved verification file to: {v_file}")

    print("Starting in 3 seconds...")
    time.sleep(3)
    
    validated_first_student = False

    for student in students:
        sid = student['id']
        lname = student['last_name']
        choice_groups = student.get('choices_groups', [])
        errors = student.get('errors', [])
        
        print(f"Processing: {sid} - {lname}")
        
        # 0. Log pre-existing errors (from data_handler logic)
        if errors:
            print(f"  ⚠️  {len(errors)} error(s) logged for this student")
        for err in errors:
            log_error(sid, lname, err['raw_product'], err['reason'])
            
        if not choice_groups:
            # If no valid groups to process, skip automation for this student
            continue
        
        # 1. Search Student
        search_student(sid, coords)
        time.sleep(1.0) # Wait for student to load

        # 2. Validate Last Name (Optional)
        if 'last_name_box' in coords and lname:
            pyautogui.click(coords['last_name_box']['x'], coords['last_name_box']['y'])
            pyautogui.tripleClick()
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.1)
            found_name = pyperclip.paste().strip()
            
            # Handle hyphenated names (App might select "Walsh-" with trailing hyphen)
            expected_parts = lname.lower().split('-')
            first_part = expected_parts[0].strip().lower()
            first_part_with_hyphen = first_part + '-'
            
            # Allow match if found name is: full name, first part only, OR first part with hyphen
            is_match = (found_name.lower() == lname.lower()) or \
                       (found_name.lower() == first_part) or \
                       (found_name.lower() == first_part_with_hyphen)
            
            if not is_match:
                print(f"  -> NAME MISMATCH: Found '{found_name}', Expected '{lname}'")
                log_error(sid, lname, "ALL", f"Name Mismatch (Found: {found_name})")
                continue # Skip this student

        # Track what we entered for validation
        entry_for_validation = None

        # 3. Process Each Yearbook Choice Group
        for group in choice_groups:
            photo_choice = group['photo_choice']
            standard_string = group['standard_string']
            other_items = group['others']
            
            # Process choice group silently

            # A. Click Photo Choice Letter (Once per group)
            if photo_choice:
                choice_key = f"choice_{photo_choice}"
                if choice_key in coords:
                    pyautogui.click(coords[choice_key]['x'], coords[choice_key]['y'])
                    time.sleep(0.5)
                else:
                    log_error(sid, lname, "Photo Choice", f"Coordinate for choice '{photo_choice}' not found")
            else:
                 # If None, we assume we skip this or default is acceptable
                 pass
                
            # B. Input Standard Packages (The combined string, e.g. "xxyy")
            if standard_string:
                if 'quick_package_entry_box' in coords:
                    click_and_type(coords['quick_package_entry_box'], standard_string)
                    
                    # Capture for validation if not yet validated
                    if not validated_first_student:
                        entry_for_validation = standard_string
                else:
                    log_error(sid, lname, "Standard Package", "'quick_package_entry_box' coordinate missing")
            
            # C. Input Other Items (Group, CD, Touchup) - processed individually
            for item in other_items:
                p_code = item['code']
                target_box_name = item['target_box']
                
                if target_box_name:
                    if target_box_name == 'touchup':
                        if 'touchup_dropdown' in coords:
                            click_and_type(coords['touchup_dropdown'], "Pending")
                        else:
                            log_error(sid, lname, "Touchup", "'touchup_dropdown' coordinate missing")

                    elif target_box_name in coords:
                         click_and_type(coords[target_box_name], p_code)
                    else:
                        log_error(sid, lname, item['raw_product'], f"Missing Coordinate: {target_box_name}")

        # 4. Perform Validation Check (Only for the first successful standard entry)
        if not validated_first_student and entry_for_validation:
            print(f"\n*** VALIDATING FIRST ENTRY: {entry_for_validation} ***")
            
            # A. Re-Search Student (to refresh view)
            search_student(sid, coords)
            time.sleep(1.0) # Wait for student to load

            # B. Check the box
            found_pkg = read_field_text(coords.get('quick_package_entry_box'))
            
            if found_pkg.lower() == entry_for_validation.lower():
                print("✓ Validation passed")
                validated_first_student = True
            else:
                print(f"✗ VALIDATION FAILED: Expected '{entry_for_validation}', Found '{found_pkg}'")
                log_error(sid, lname, f"Standard Pkg: {entry_for_validation}", f"VALIDATION FAILED (Found: '{found_pkg}' in quick package entry box when it should be {entry_for_validation})")
                return False

        time.sleep(0.5) # Pause between students

    print("Automation Complete!")
    return True

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

    pyautogui.tripleClick()
    time.sleep(0.5)
    
    # Clear clipboard first
    pyperclip.copy("")
    
    # Copy
    pyautogui.hotkey('ctrl', 'c') 
    time.sleep(0.1)
    
    return pyperclip.paste().strip()

if __name__ == "__main__":
    import sys
    try:
        run_automation()
    except KeyboardInterrupt:
        print("\nABORTED BY USER.")
    except Exception as e:
        print(f"\nCRITICAL ERROR: {e}")
