import pandas as pd
from excel_utils import find_column_robust, get_excel_path
import os
import sys
from datetime import datetime

def validate_data():
    print("--- Starting Data Validation ---")
    
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

    # 3. Validate Rows
    error_rows = []
    valid_count = 0

    # Calculate counts per student to handle single-row exception
    student_counts = df[student_id_col].value_counts()

    # Iterate through rows
    for index, row in df.iterrows():
        errors = []
        
        # Check Student ID (Basic check: not empty)
        sid = row[student_id_col]
        if pd.isna(sid) or str(sid).strip() == "":
            errors.append("Missing Student ID")
            row_count = 0
        else:
            row_count = student_counts.get(sid, 0)

        # Check Selection
        sel = row[selection_col]
        if pd.isna(sel) or str(sel).strip().lower() not in ['a', 'b', 'c', 'd']:
            errors.append(f"Invalid Selection: '{sel}' (Must be a, b, c, or d)")

        # Check Date
        date_val = row[date_col]
        try:
             # Attempt to convert to datetime to verify it's a date
             if pd.isna(date_val):
                 # Missing date is acceptable IF it's a single row
                 if row_count > 1:
                     errors.append("Missing Date (Required for sorting multiple entries)")
             else:
                 pd.to_datetime(date_val)
        except Exception:
            # Date is invalid (format issue)
            if row_count > 1:
                errors.append(f"Invalid Date Format: '{date_val}' (Required for sorting multiple entries)")
            else:
                # Single row: we tolerate the invalid date
                pass

        if errors:
            # Create error entry
            error_entry = row.copy()
            # Combine multiple errors if any
            error_entry['error_reason'] = "; ".join(errors)
            error_rows.append(error_entry)
        else:
            valid_count += 1

    # 4. Report
    print(f"\nValidation Complete.")
    print(f"Valid Rows: {valid_count}")
    print(f"Invalid Rows: {len(error_rows)}")

    if error_rows:
        reports_dir = "reports"
        if not os.path.exists(reports_dir):
            os.makedirs(reports_dir)
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_file = os.path.join(reports_dir, f"data-validation-{timestamp}.csv")
        
        # Create DataFrame
        error_df = pd.DataFrame(error_rows)
        
        error_df.to_csv(report_file, index=False)
        print(f"-> Error report generated: {report_file}")
        print("Note: These rows will be skipped during automation. Proceeding with valid data...")
        # sys.exit(1)  <-- Removed to allow continuation
    else:
        print("-> No errors found. Data is ready for automation.")

if __name__ == "__main__":
    validate_data()
