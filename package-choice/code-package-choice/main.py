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
_error_log = []
_completed_log = []

def load_coordinates():
    if not os.path.exists(COORD_FILE):
        print("Error: coordinates_package.json not found. Run config_wizard_package.py first!")
        return None
    with open(COORD_FILE, "r") as f:
        return json.load(f)

def log_error(student_id, first_name, last_name, product_raw, reason):
    """Append to _error_log; written to Excel on each write (after load, at end, on abort)."""
    entry = {
        "student_id": student_id,
        "first_name": first_name,
        "last_name": last_name,
        "product_raw": product_raw,
        "error_reason": reason,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    _error_log.append(entry)

def _write_error_report():
    if not _error_log:
        return
    reports_dir = "reports"
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
    filename = os.path.join(reports_dir, f"package-choice-errors-{SESSION_TIMESTAMP}.xlsx")
    try:
        pd.DataFrame(_error_log).to_excel(filename, index=False)
        print(f"Error report saved: {filename}")
    except Exception as e:
        print(f"Failed to write error report: {e}")

def log_completed(student):
    """Buffer completed students; written to Excel at end or on abort."""
    row = {
        "student_id": student["id"],
        "first_name": student.get("first_name", ""),
        "last_name": student.get("last_name", ""),
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    _completed_log.append(row)

def _write_completed_report():
    if not _completed_log:
        return
    reports_dir = "reports"
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
    filename = os.path.join(reports_dir, f"package-choice-completed-{SESSION_TIMESTAMP}.xlsx")
    try:
        pd.DataFrame(_completed_log).to_excel(filename, index=False)
        print(f"Completed list saved: {filename}")
    except Exception as e:
        print(f"Failed to write completed report: {e}")

def click_and_type(coord, text, clear_with_backspace=False):
    if not coord:
        return
    pyautogui.click(coord['x'], coord['y'])
    if clear_with_backspace:
        # Class Pkg / Class Pix No Pkg: we clear by backspace. Cursor must be at end of text (user note in startup instructions).
        for _ in range(40):
            pyautogui.press('backspace')
        time.sleep(0.05)
    else:
        pyautogui.tripleClick()
        time.sleep(0.05)
    pyautogui.typewrite(str(text))
    time.sleep(0.05)

def run_automation():
    global _error_log, _completed_log
    _error_log = []
    _completed_log = []
    coords = load_coordinates()
    if not coords:
        return False

    pyautogui.FAILSAFE = True # Enabled by default, but making it explicit
    
    print("--- READY TO START PACKAGE ENTRY ---")
    print("1. Ensure School Days app is open and ready.")
    print("2. IMPORTANT: Manually CHECK all input boxes (like Touchup) so they are editable!")
    print("3. RERUNS: If you are re-running and the Class Pkg or Class Pix No Pkg box already has text,")
    print("   click inside that box and put the cursor at the END of the text (so backspace clears it correctly).")
    print("4. EMERGENCY STOP: Slam mouse quickly to any corner of the screen.")
    print("5. OR click on this Terminal window and press Ctrl+C.")
    print("------------------------------------")
    
    students = load_and_process_data(None) # Auto-finds Excel

    if not students:
        print("No student data found or processed.")
        return False
        
    print(f"Loaded {len(students)} students to process.")

    # Log all data-load errors; then write Excel so it exists even if run is aborted
    for s in students:
        for err in s.get('errors', []):
            log_error(s['id'], s.get('first_name', ''), s.get('last_name', ''), err.get('raw_product', ''), err.get('reason', ''))
    _write_error_report()

    # Verification report: summary of package choices we're entering (before run starts)
    print("Generating run summary...")
    reports_dir = "reports"
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
        
    verif_data = []
    for s in students:
        sid = s['id']
        fname = s.get('first_name', '')
        lname = s.get('last_name', '')
        for grp in s.get('choices_groups', []):
            # Organize items by target box
            quick_pkg = grp['standard_string']
            cd_value = ""
            touchup_value = ""
            class_pix = ""
            class_pix_no_pkg = ""
            
            # Process other items and group by target box
            for item in grp['others']:
                # Skip error items
                if 'lost order' in item['raw_product'].lower() or 'invalid' in item['raw_product'].lower():
                    continue
                    
                target_box = item.get('target_box', '')
                code = item['code']
                
                if target_box == 'cd_box':
                    cd_value = code
                elif target_box == 'touchup':
                    touchup_value = "Pending"  # Always "Pending" for touchup
                elif target_box == 'class_pkg_box':
                    # Append to class_pix (may have multiple group prints)
                    if class_pix:
                        class_pix += ", " + code
                    else:
                        class_pix = code
                elif target_box == 'class_pix_no_pkg_box':
                    # Append to class_pix_no_pkg
                    if class_pix_no_pkg:
                        class_pix_no_pkg += ", " + code
                    else:
                        class_pix_no_pkg = code
            
            verif_data.append({
                'Student ID': sid,
                'Last Name': lname,
                'Photo Choice': grp['photo_choice'] if grp['photo_choice'] else "(NONE)",
                'Quick Package Entry': quick_pkg,
                'CD': cd_value,
                'Touchup': touchup_value,
                'Class Pix': class_pix,
                'Class Pix No Pkg': class_pix_no_pkg
            })
            
    if verif_data:
        v_df = pd.DataFrame(verif_data)
        v_file = os.path.join(reports_dir, f"package-choice-run-summary-{SESSION_TIMESTAMP}.xlsx")
        v_df.to_excel(v_file, index=False)
        print(f"Run summary saved: {v_file}")

    print("Starting in 3 seconds...")
    time.sleep(3)
    
    validated_first_student = False

    for student in students:
        sid = student['id']
        fname = student.get('first_name', '')
        lname = student.get('last_name', '')
        choice_groups = student.get('choices_groups', [])
        errors = student.get('errors', [])
        
        print(f"Processing: {sid} - {lname}")
        
        # 0. Log pre-existing errors (from data_handler logic; data-check errors already in report from load)
        if errors:
            print(f"  ⚠️  {len(errors)} error(s) logged for this student")
        for err in errors:
            if err.get('raw_product') != '(data check)':
                log_error(sid, fname, lname, err['raw_product'], err['reason'])
            
        if not choice_groups:
            # If no valid groups to process, skip automation for this student
            continue
        
        # 1. Search Student
        search_student(sid, coords)
        time.sleep(0.3) # Wait for student to load

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
                log_error(sid, fname, lname, "ALL", f"Name Mismatch (Found: {found_name})")
                continue # Skip this student

        # 2.5 Audit Trail (Web Entry -> "auto")
        if 'web_entry_box' in coords:
            click_and_type(coords['web_entry_box'], "auto")

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
                    time.sleep(0.1)
                else:
                    log_error(sid, fname, lname, "Photo Choice", f"Coordinate for choice '{photo_choice}' not found")
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
                    log_error(sid, fname, lname, "Standard Package", "'quick_package_entry_box' coordinate missing")
            
            # C. Input Other Items (Group, CD, Touchup) - processed individually
            for item in other_items:
                p_code = item['code']
                target_box_name = item['target_box']
                
                if target_box_name:
                    if target_box_name == 'touchup':
                        if 'touchup_dropdown' in coords:
                            click_and_type(coords['touchup_dropdown'], "Pending")
                        else:
                            log_error(sid, fname, lname, "Touchup", "'touchup_dropdown' coordinate missing")

                    elif target_box_name in coords:
                         # Class pkg / class pix no pkg boxes: triple-click doesn't clear reliably, use backspace
                         clear_backspace = target_box_name in ('class_pkg_box', 'class_pix_no_pkg_box')
                         click_and_type(coords[target_box_name], p_code, clear_with_backspace=clear_backspace)
                    else:
                        log_error(sid, fname, lname, item['raw_product'], f"Missing Coordinate: {target_box_name}")

        # 4. Perform Validation Check (Only for the first successful standard entry)
        if not validated_first_student and entry_for_validation:
            print(f"\n*** VALIDATING FIRST ENTRY: {entry_for_validation} ***")
            
            # A. Re-Search Student (to refresh view)
            search_student(sid, coords)
            time.sleep(0.3) # Wait for student to load

            # B. Check the box
            found_pkg = read_field_text(coords.get('quick_package_entry_box'))
            
            if found_pkg.lower() == entry_for_validation.lower():
                print("✓ Validation passed")
                validated_first_student = True
            else:
                print(f"✗ VALIDATION FAILED: Expected '{entry_for_validation}', Found '{found_pkg}'")
                log_error(sid, fname, lname, f"Standard Pkg: {entry_for_validation}", f"VALIDATION FAILED (Found: '{found_pkg}' in quick package entry box when it should be {entry_for_validation})")
                _write_error_report()
                return False

        log_completed(student)
        time.sleep(0.1) # Pause between students

    _write_error_report()
    _write_completed_report()
    print("Automation Complete!")
    return True

def search_student(sid, coords):
    if 'search_box' in coords:
        pyautogui.click(coords['search_box']['x'], coords['search_box']['y'])
        pyautogui.doubleClick() 
        pyautogui.typewrite(sid)
        pyautogui.press('enter')
        time.sleep(0.1) # Wait for load

def read_field_text(coord):
    """
    Clicks field, Selects All, Copies to clipboard, returns text.
    """
    if not coord: return ""
    
    # Click and focus
    pyautogui.click(coord['x'], coord['y'])

    pyautogui.tripleClick()
    time.sleep(0.1)
    
    # Clear clipboard first
    pyperclip.copy("")
    
    # Copy
    pyautogui.hotkey('ctrl', 'c') 
    time.sleep(0.1)
    
    return pyperclip.paste().strip()

import sys

if __name__ == "__main__":
    try:
        if not run_automation():
            sys.exit(1)
    except pyautogui.FailSafeException:
        print("\n[EMERGENCY STOP] Failsafe triggered by moving mouse to corner.")
        _write_error_report()
        _write_completed_report()
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n[ABORTED] Stopped by user (Ctrl+C).")
        _write_error_report()
        _write_completed_report()
        sys.exit(1)
    except Exception as e:
        print(f"\n[CRITICAL ERROR] {e}")
        _write_error_report()
        _write_completed_report()
        sys.exit(1)
