import pyautogui
import json
import os
import time

COORD_FILE = os.path.join(os.path.dirname(__file__), "coordinates_package.json")

def get_coordinate(prompt_name):
    print(f"\n--- {prompt_name} ---")
    print("1. Move your mouse cursor.")
    print("2. Press 'Enter' when ready (do not click).")
    input("Waiting for Enter...")
    point = pyautogui.position()
    print(f"Captured: {point}")
    return {"x": point.x, "y": point.y}

def run_wizard():
    print("Welcome to the School Days Package Entry Setup.")
    print("We need to learn where the buttons are on YOUR screen.")
    print("-----------------------------------------------------")
    
    coords = {}
    
    # 1. Search Box
    coords["search_box"] = get_coordinate("SEARCH INPUT BOX (Where you type Student ID)")

    # 1.5 Last Name Box (For Validation)
    coords["last_name_box"] = get_coordinate("LAST NAME FIELD (To verify student exists)")
    
    # 2. Yearbook Choices (Letters)
    print("\nPlease locate the Yearbook Choice buttons/radios.")
    # Assuming standard letters A-F or similar mapping
    # The user said "click on the photo choice".
    coords["choice_a"] = get_coordinate("PHOTO CHOICE 'A' / 'a'")
    coords["choice_b"] = get_coordinate("PHOTO CHOICE 'B' / 'b'")
    coords["choice_c"] = get_coordinate("PHOTO CHOICE 'C' / 'c'")
    coords["choice_d"] = get_coordinate("PHOTO CHOICE 'D' / 'd'")

    # Some extra common ones?
    # No harm in gathering, but the prompt only specified letters and "click on photo choice"

    # 3. Package Entry Fields
    print("\nNow locate the Data Entry fields.")
    
    coords["quick_package_entry_box"] = get_coordinate("QUICK PACKAGE ENTRY BOX (For standard packages)")
    coords["class_pkg_box"] = get_coordinate("CLASS PKG BOX (For Group Prints WITH personal pkg)")
    coords["class_pix_no_pkg_box"] = get_coordinate("CLASS PIX NO PKG BOX (For Group Prints WITHOUT personal pkg)")
    
    coords["cd_box"] = get_coordinate("CD INPUT BOX (For 'All 4 Digital Portraits')")
    
    # Touchup
    # User requested to treat this as a standard input box and type "Pending"
    coords["touchup_dropdown"] = get_coordinate("TOUCHUP BOX (Will type 'Pending')")
    
    # coords["touchup_pending_option"] = get_coordinate("TOUCHUP 'PENDING' OPTION") # Removed per request

    with open(COORD_FILE, "w") as f:
        json.dump(coords, f, indent=4)
    
    print(f"\nSuccess! Coordinates saved to {COORD_FILE}.")

if __name__ == "__main__":
    run_wizard()
