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
    
    if not excel_path:
        excel_path = get_excel_path()
        
        if not excel_path:
            print("Error: No Excel file (.xlsx) found in this folder.")
            print("Please assume user has put one .xlsx file here.")
            return []
            
    print(f"Found Excel file: {excel_path}")

    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Error reading {excel_path}: {e}")
        return []
    
    processed_data = []

    # Group by Student ID
    student_id_col = find_column_robust(df, "student id")
    selection_col = find_column_robust(df, ["yearbook photo"])
    date_col = find_column_robust(df, "yearbook date")
    last_name_col = find_column_robust(df, "student last name")

    if not student_id_col or not selection_col or not date_col or not last_name_col:
        if not student_id_col:
            print("Error: Could not find 'Student ID' column.")
        if not selection_col:
            print("Error: Could not find 'Yearbook Selection' column.")
        if not date_col:
            print("Error: Could not find 'Yearbook Date' column.")
        if not last_name_col:
            print("Error: Could not find 'Student Last Name' column.")
        print("Columns found:", list(df.columns))
        return []
    
    unique_ids = df[student_id_col].unique()
    print(f"Found {len(unique_ids)} unique sorted by Student ID.")
    
    # ---------------------------------------------------------
    # ERROR REPORTING & VALIDATION
    # ---------------------------------------------------------
    invalid_rows = []
    
    # ---------------------------------------------------------
    # DATA PROCESSING LOOP
    # ---------------------------------------------------------
    # We moved validation inside the loop to support "Single Row = Date Optional" logic.
    pass



    # SAVE ERROR REPORT - Moved to end of function to include conflicts


    # ---------------------------------------------------------
    # MAIN LOGIC using filtered 'df'
    # ---------------------------------------------------------
    
    unique_ids = df[student_id_col].dropna().unique()

    for student_id in unique_ids:
        # Filter for this student
        student_rows = df[df[student_id_col] == student_id].copy()
        
        if student_rows.empty:
            continue

        # Validate duplicates if we cannot sort by date
        if len(student_rows) > 1 and not date_col:
            print(f"Error: Duplicate rows found for student {student_id} and no 'Yearbook Date' column exists to prioritize.")
            for idx, row in student_rows.iterrows():
                invalid_rows.append(row)
            continue

        # Handle Date Logic
        # New Rule: If single row, Date is optional/can be invalid.
        #           If multiple rows, Date must be valid to sort.
        
        if date_col:
             # Convert dates for this student's rows
             # We do this here (locally) instead of globally.
             
             # Create a series of converted dates
             row_dates = pd.to_datetime(student_rows[date_col], errors='coerce')
             
             # Check for invalid dates
             invalid_date_mask = row_dates.isna() & student_rows[date_col].notna()
             has_invalid_dates = invalid_date_mask.any()
             
             if len(student_rows) > 1:
                 # MULTIPLE ROWS: All dates matching "latest" logic must be valid. 
                 # Actually, if ANY date is invalid, can we trust the sort?
                 # Safest is to require valid dates for ALL rows if we need to sort.
                 
                 if has_invalid_dates:
                     print(f"Error: Student {student_id} has multiple rows but invalid dates. Cannot sort.")
                     # Log all rows as invalid
                     for idx, row in student_rows.iterrows():
                         invalid_rows.append(row)
                     continue
                 
                 # Dates are valid, overwrite with datetime objects for sorting
                 student_rows[date_col] = row_dates
                 
                 # Sort desc
                 student_rows = student_rows.sort_values(by=date_col, ascending=False)
                 
             else:
                 # SINGLE ROW: We ignore date validity.
                 pass
        
        # If no date column, we already filtered for duplicates above?
        # The logic at L103 checked duplicates if NO date_col exists at all.
        
        # ------------------------------------------------------------------
        # Conflict Check (Only for multiple rows with same VALID top date)
        # ------------------------------------------------------------------
        if len(student_rows) > 1 and date_col:
            
            # Check for conflicts: multiple entries with the same LATEST date
            top_date = student_rows.iloc[0][date_col]
            
            # Compare strictly? NaT should not be here as we filtered them, but check anyway.
            # Assuming top_date is valid.
            
            # Find all rows with this same date
            latest_rows = student_rows[student_rows[date_col] == top_date]
            
            if len(latest_rows) > 1:
                # Check if they have different selections
                selections = set()
                if selection_col:
                     for idx, row in latest_rows.iterrows():
                         s_val = str(row[selection_col]).lower().strip()
                         selections.add(s_val)
                
                if len(selections) > 1:
                    print(f"Conflict usage for student {student_id}: Multiple selections {selections} on {top_date}")
                    # Add all conflicting rows to invalid list
                    for idx, row in latest_rows.iterrows():
                        # Make a copy to add error reason?
                        # Or just append. The user will see duplicate rows with same date in the error file.
                        # Maybe we can add a 'Error Reason' column if we want, but keeping it simple for now.
                        invalid_rows.append(row)
                    
                    # Skip this student
                    continue

        # Take top
        top_row = student_rows.iloc[0]
        
        # Get selection
        selection_val = 'd' 
        if selection_col:
            raw_val = str(top_row[selection_col]).lower().strip()
            if raw_val in ['a', 'b', 'c', 'd']:
                selection_val = raw_val
            else:
                print(f"Invalid selection '{raw_val}' for student {student_id}. Skipping.")
                invalid_rows.append(top_row)
                continue
        
        entry = {
            'id': str(top_row[student_id_col]),
            'last_name': str(top_row[last_name_col]).strip() if last_name_col else "",
            'selection': str(selection_val).lower().strip(), 
        }
        processed_data.append(entry)
        
    return processed_data