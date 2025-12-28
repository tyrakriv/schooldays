import pandas as pd
from excel_utils import find_column_robust, get_excel_path
import os
import sys
from datetime import datetime

def validate_data():
    print("--- Starting Data Validation ---")
    
    # 0. Define Session Log Path (Shared with main.py)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    reports_dir = "reports"
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
    
    # We use a single file for both Setup and Run errors
    report_file = os.path.join(reports_dir, f"session-errors-{timestamp}.csv")
    
    # Save this path to a temp file so main.py knows where to log
    session_info_path = os.path.join(os.path.dirname(__file__), "current_session.txt")
    try:
        with open(session_info_path, "w") as f:
            f.write(report_file)
    except Exception as e:
        print(f"Warning: Could not save session info: {e}")

    # 1. Find Excel File
    excel_path = get_excel_path()
    
    if not excel_path:
        print("Error: No Excel file (.xlsx) found in this folder.")
        sys.exit(1)
    print(f"Checking file: {excel_path}")

    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Critical Error: Could not read Excel file. {e}")
        sys.exit(1)

    # 2. Check Columns
    student_id_col = find_column_robust(df, "student id")
    selection_col = find_column_robust(df, ["yearbook photo", "selection"])
    date_col = find_column_robust(df, "yearbook date")
    last_name_col = find_column_robust(df, "student last name")

    missing_cols = []
    if not student_id_col: missing_cols.append("Student ID")
    if not selection_col: missing_cols.append("Yearbook Selection")
    if not date_col: missing_cols.append("Yearbook Date")
    if not last_name_col: missing_cols.append("Student Last Name")

    if missing_cols:
        print(f"Error: Missing required columns: {', '.join(missing_cols)}")
        print(f"Found columns: {list(df.columns)}")
        sys.exit(1)

    print("Columns identified successfully.")

    # 3. Process Data (Clean, Dedup, Sort)
    cleaned_rows = []
    error_rows = []
    
    unique_ids = df[student_id_col].unique()
    print(f"Processing {len(unique_ids)} unique Student IDs...")
    
    for student_id in unique_ids:
        # Filter rows for this student
        student_rows = df[df[student_id_col] == student_id].copy()
        
        if student_rows.empty:
            continue

        row_count = len(student_rows)
        
        # --- Date Logic ---
        if date_col:
             # Convert dates (Strict + Fallback)
             timestamp_fmt = '%m/%d/%Y %I:%M:%S %p'
             
             # Attempt parse
             row_dates = pd.to_datetime(student_rows[date_col], format=timestamp_fmt, errors='coerce')
             
             # Fallback for manual entries
             original_vals = student_rows[date_col]
             failed_parse_mask = row_dates.isna() & original_vals.notna()
             
             if failed_parse_mask.any():
                 import warnings
                 with warnings.catch_warnings():
                     warnings.simplefilter("ignore")
                     fallback_dates = pd.to_datetime(original_vals[failed_parse_mask], errors='coerce')
                 
                 row_dates.loc[failed_parse_mask] = fallback_dates
             
             # Apply parsed dates
             student_rows[date_col] = row_dates
             
             # Check Validity
             invalid_date_mask = row_dates.isna() & original_vals.notna()
             
             # Rule: Multiple rows require valid dates for sorting
             if row_count > 1 and invalid_date_mask.any():
                 # Reject all rows for this student
                 for idx, row in student_rows.iterrows():
                     err_row = row.copy()
                     err_row['error_reason'] = "Multiple rows with invalid dates"
                     error_rows.append(err_row)
                 continue
                 
             # Sort Descending (Newest First)
             student_rows = student_rows.sort_values(by=date_col, ascending=False)
        
        elif row_count > 1:
            # Multiple rows but no date column at all
            print(f"Error: Student {student_id} has duplicates but no Date column.")
            for idx, row in student_rows.iterrows():
                 err_row = row.copy()
                 err_row['error_reason'] = "Duplicate rows without Date column"
                 error_rows.append(err_row)
            continue

        # --- Top Row Selection ---
        top_row = student_rows.iloc[0]
        
        # --- Conflict Check (Optional: Same Date) ---
        if row_count > 1 and date_col:
            top_date = top_row[date_col]
            same_date_rows = student_rows[student_rows[date_col] == top_date]
            
            if len(same_date_rows) > 1:
                 # Check selection conflict
                 selections = set()
                 for idx, r in same_date_rows.iterrows():
                     s_val = str(r[selection_col]).lower().strip() if selection_col else ""
                     selections.add(s_val)
                 
                 if len(selections) > 1:
                     print(f"Error: Student {student_id} has conflicting selections on same date {top_date}.")
                     for idx, row in same_date_rows.iterrows():
                         err_row = row.copy()
                         err_row['error_reason'] = f"Conflicting selections {selections} on same date"
                         error_rows.append(err_row)
                     continue

        # --- Validate Selection ---
        if selection_col:
            sel = top_row[selection_col]
            if pd.isna(sel) or str(sel).strip().lower() not in ['a', 'b', 'c', 'd']:
                err_row = top_row.copy()
                err_row['error_reason'] = f"Invalid Selection: '{sel}'"
                error_rows.append(err_row)
                continue
        
        # --- Add to Clean List ---
        cleaned_rows.append(top_row)

    # 4. Save Outputs
    print(f"\nProcessing Complete.")
    print(f"Cleaned Rows: {len(cleaned_rows)}")
    print(f"Error Rows: {len(error_rows)}")

    # Save Cleaned Data
    # Use code folder path
    clean_file = os.path.join(os.path.dirname(__file__), "cleaned_data.xlsx")
    if cleaned_rows:
        try:
            clean_df = pd.DataFrame(cleaned_rows)
            # Reorder columns to put ID first, if possible, but keeping original structure is fine
            clean_df.to_excel(clean_file, index=False)
        except Exception as e:
            print(f"Critical Error saving cleaned data: {e}")
            sys.exit(1)
    else:
        print("-> Warning: No valid data found to save.")
        if os.path.exists(clean_file):
            os.remove(clean_file) 

    # Save Errors
    if error_rows:
        # report_file is already defined at top
        
        error_df = pd.DataFrame(error_rows)
        error_df.to_csv(report_file, index=False)
        print(f"-> Error report started: {report_file}")
        
    if not cleaned_rows:
         sys.exit(1) # Fail if nothing to run
    else:
         print("-> Ready for automation.")

if __name__ == "__main__":
    validate_data()
