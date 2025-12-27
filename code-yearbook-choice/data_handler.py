import pandas as pd
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
    
    if date_col:
        # Identify invalid dates
        # coerce errors -> NaT
        temp_dates = pd.to_datetime(df[date_col], errors='coerce')
        
        # Find indices where date is NaT but original was not empty
        
        mask_invalid = temp_dates.isna() & df[date_col].notna()
        # also check empty strings if they are not considered NaN by pandas read
        
        # Get the invalid rows
        bad_rows_df = df[mask_invalid]
        
        if not bad_rows_df.empty:
            print(f"Found {len(bad_rows_df)} rows with invalid dates.")
            # Add to invalid list
            for idx, row in bad_rows_df.iterrows():
                invalid_rows.append(row)
                
            # Filter main DF to exclude these rows for automation
            df = df[~mask_invalid].copy()
            # Update date column in clean DF to the datetime objects
            df[date_col] = temp_dates[~mask_invalid]
    else:
        # If no date column, we can't sort by date, but we don't error out.
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

        if date_col and date_col in student_rows.columns:
            # Sort desc
            student_rows = student_rows.sort_values(by=date_col, ascending=False)
            
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
        
    # SAVE ERROR REPORT (Again, in case we added new invalid rows inside the loop)
    # We might have saved some earlier. We should probably do one save at the end?
    # Or append? 
    # Current code saved `invalid_rows` BEFORE the loop. 
    # Let's move the save logic to AFTER the loop to include conflicts.
    
    if invalid_rows:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        reports_dir = "reports"
        if not os.path.exists(reports_dir):
            os.makedirs(reports_dir)
            
        error_filename = os.path.join(reports_dir, f"unknown_errors_{timestamp}.csv")
        
        try:
            pd.DataFrame(invalid_rows).to_csv(error_filename, index=False)
            print(f"Reported {len(invalid_rows)} invalid/conflicting entries to: {error_filename}")
        except Exception as e:
            print(f"Failed to save error report: {e}")
            
    return processed_data
        

