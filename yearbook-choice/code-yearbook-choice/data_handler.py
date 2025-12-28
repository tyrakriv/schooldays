import pandas as pd
import os
from datetime import datetime

def load_and_process_data(excel_path):
    """
    Reads the Excel file, groups by Student ID, and finds the entry 
    with the latest Yearbook Date.
    
    Returns: A list of dicts, each representing a student to process.
    """
    from excel_utils import find_column_robust, get_excel_path
    
    cleaned_path = os.path.join(os.path.dirname(__file__), "cleaned_data.xlsx")
    
    if not os.path.exists(cleaned_path):
        print("Error: 'cleaned_data.xlsx' not found.")
        print("Please run Step 1 (Setup/Validation) first to generate it.")
        return []

    print(f"Loading cleaned data from: {cleaned_path}")

    try:
        df = pd.read_excel(cleaned_path)
    except Exception as e:
        print(f"Error reading {cleaned_path}: {e}")
        return []
    
    processed_data = []

    # Identify columns again (robustly, just in case)
    student_id_col = find_column_robust(df, "student id")
    selection_col = find_column_robust(df, ["yearbook photo", "selection"])
    last_name_col = find_column_robust(df, "student last name")
    
    if not student_id_col or not selection_col or not last_name_col:
        print("Error: 'Student ID', 'Yearbook Selection', or 'Student Last Name' column missing in cleaned data.")
        return []

    # Iterate - Data is already deduped and valid
    for index, row in df.iterrows():
        sid = row[student_id_col]
        
        # Selection
        if selection_col and pd.notna(row[selection_col]):
             sel = str(row[selection_col]).lower().strip()
        else:
             sel = 'd' # Default
        
        # Last Name
        lname = ""
        if last_name_col and pd.notna(row[last_name_col]):
             lname = str(row[last_name_col]).strip()
        
        entry = {
            'id': str(sid),
            'last_name': lname,
            'selection': sel
        }
        processed_data.append(entry)

    return processed_data