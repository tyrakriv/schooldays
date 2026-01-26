import pyautogui
import json
import os
import time

COORD_FILE = os.path.join(os.path.dirname(__file__), "coordinates.json")

def get_coordinate(prompt_name):
    print(f"\n--- {prompt_name} ---")
    print("1. Move your mouse cursor.")
    print("2. Press 'Enter' when ready (do not click).")
    input("Waiting for Enter...")
    point = pyautogui.position()
    print(f"Captured: {point}")
    return {"x": point.x, "y": point.y}

def run_wizard():
    print("Welcome to the School Days - Package Choice Setup.")
    print("We need to learn where the buttons are on YOUR screen.")
    print("-----------------------------------------------------")
    
    coords = {}
    
    # 1. Search Box
    coords["search_box"] = get_coordinate("SEARCH INPUT BOX (Where you type Student ID)")

    # 1.5 Last Name Box (For Validation)
    coords["last_name_box"] = get_coordinate("LAST NAME FIELD (To verify student exists)")
    coords["last_name_checkbox"] = get_coordinate("LAST NAME CHECKBOX (To allow input)")

    # 2. Yearbook Photo Choice
    print("\n--- YEARBOOK PHOTO CHOICE ---")
    print("We need locations for the yearbook choice selections (A, B, C, etc.)")
    print("If you only use A-D, capture those. If more, add them.")
    coords["choice_a"] = get_coordinate("PHOTO CHOICE 'a'")
    coords["choice_b"] = get_coordinate("PHOTO CHOICE 'b'")
    coords["choice_c"] = get_coordinate("PHOTO CHOICE 'c'")
    coords["choice_d"] = get_coordinate("PHOTO CHOICE 'd'")
    coords["choice_e"] = get_coordinate("PHOTO CHOICE 'e'")
    coords["choice_f"] = get_coordinate("PHOTO CHOICE 'f'")

    # 3. Package Entry Fields
    print("\n--- PACKAGE ENTRY FIELDS ---")
    coords["quick_package_entry_box"] = get_coordinate("QUICK PACKAGE ENTRY BOX (For standard packages)")
    coords["class_pkg_box"] = get_coordinate("CLASS PKG BOX (Group print WITH personal pkg)")
    coords["class_pix_no_pkg_box"] = get_coordinate("CLASS PIX NO PKG BOX (Group print ONLY)")
    coords["cd_box"] = get_coordinate("CD/DIGITAL INPUT BOX")
    
    # 4. Touchups
    print("\n--- TOUCHUP FIELDS ---")
    coords["touchup_dropdown"] = get_coordinate("TOUCHUP DROPDOWN MENU")
    
    print("\n>>> Please CLICK the touchup dropdown now to open it so you can hover over the 'Pending' option. <<<")
    input("Press Enter AFTER you have opened the dropdown...")
    coords["touchup_pending_option"] = get_coordinate("TOUCHUP 'PENDING' OPTION")

    with open(COORD_FILE, "w") as f:
        json.dump(coords, f, indent=4)
    
    print(f"\nSuccess! Coordinates saved to {COORD_FILE}.")

if __name__ == "__main__":
    run_wizard()
