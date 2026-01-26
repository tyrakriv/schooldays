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
    print("Welcome to the School Days Plus Automation Setup.")
    print("We need to learn where the buttons are on YOUR screen.")
    print("-----------------------------------------------------")
    
    coords = {}
    
    # 1. Search Box
    coords["search_box"] = get_coordinate("SEARCH INPUT BOX (Where you type Student ID)")

    # 1.5 Last Name Box (For Validation)
    coords["last_name_box"] = get_coordinate("LAST NAME FIELD (To verify student exists)")
       
    # 2. Audit Trail Locations
    coords["web_entry_input_box"] = get_coordinate("WEB ENTRY INPUT (Where we type 'auto')")

    # 3. Option Locations
    print("\nNow we need the locations for the Yearbook selection list.")
    
    coords["option_a"] = get_coordinate("OPTION 'A' (The first item in the list)")
    coords["option_b"] = get_coordinate("OPTION 'B' (The second item in the list)")
    coords["option_c"] = get_coordinate("OPTION 'C' (The third item in the list)")
    # We don't need 'd' if it's default, but good to have if we need to explicitly click it later.
    coords["option_d"] = get_coordinate("OPTION 'D' (The fourth/last item)")


    with open(COORD_FILE, "w") as f:
        json.dump(coords, f, indent=4)
    
    print(f"\nSuccess! Coordinates saved to {COORD_FILE}.")

if __name__ == "__main__":
    run_wizard()
