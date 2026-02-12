import pandas as pd
from excel_utils import find_column_robust, get_excel_path
import os
import sys
from datetime import datetime

def validate_data_from_path(excel_path):
    """
    Run validation on the given Excel path. Returns (cleaned_rows, error_rows).
    cleaned_rows and error_rows are lists of row-like dicts/Series with original
    columns; error_rows entries also have 'error_reason'. No files are written.
    """
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        raise RuntimeError(f"Could not read Excel file: {e}") from e

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
        raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")

    cleaned_rows, error_rows = _process_validation(df, student_id_col, selection_col, date_col, last_name_col)
    cleaned = [row.to_dict() if hasattr(row, "to_dict") else row for row in cleaned_rows]
    errors = [row.to_dict() if hasattr(row, "to_dict") else row for row in error_rows]
    cols = {"student_id": student_id_col, "last_name": last_name_col, "selection": selection_col}
    return cleaned, errors, cols


def write_error_report(error_rows, filepath):
    """Write error rows to Excel (same format as validation report). Used by validate_data and tests."""
    df = pd.DataFrame(error_rows) if error_rows else pd.DataFrame()
    df.to_excel(filepath, index=False)


def _process_validation(df, student_id_col, selection_col, date_col, last_name_col):
    """Core validation loop: returns (cleaned_rows, error_rows) as lists of Series."""
    cleaned_rows = []
    error_rows = []

    unique_ids = df[student_id_col].unique()

    for student_id in unique_ids:
        # Flag invalid or empty student ID before processing
        sid_str = str(student_id).strip().lower() if pd.notna(student_id) else ""
        if not sid_str or sid_str == "invalid" or sid_str == "nan":
            student_rows = df[df[student_id_col] == student_id].copy()
            if not student_rows.empty:
                err_row = student_rows.iloc[0].copy()
                err_row['error_reason'] = f"Invalid Student ID: '{student_id}'"
                error_rows.append(err_row)
            continue

        student_rows = df[df[student_id_col] == student_id].copy()
        if student_rows.empty:
            continue

        row_count = len(student_rows)

        if date_col:
            timestamp_fmt = '%m/%d/%Y %I:%M:%S %p'
            row_dates = pd.to_datetime(student_rows[date_col], format=timestamp_fmt, errors='coerce')
            original_vals = student_rows[date_col]
            failed_parse_mask = row_dates.isna() & original_vals.notna()
            if failed_parse_mask.any():
                import warnings
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    fallback_dates = pd.to_datetime(original_vals[failed_parse_mask], errors='coerce')
                row_dates.loc[failed_parse_mask] = fallback_dates
            student_rows[date_col] = row_dates
            invalid_date_mask = row_dates.isna() & original_vals.notna()
            if invalid_date_mask.any():
                err_row = student_rows.iloc[0].copy()
                err_row['error_reason'] = "Multiple rows with invalid dates" if row_count > 1 else "Invalid date"
                error_rows.append(err_row)
                continue
            student_rows = student_rows.sort_values(by=date_col, ascending=False)
        elif row_count > 1:
            err_row = student_rows.iloc[0].copy()
            err_row['error_reason'] = "Duplicate rows without Date column"
            error_rows.append(err_row)
            continue

        top_row = student_rows.iloc[0]
        if row_count > 1 and date_col:
            top_date = top_row[date_col]
            same_date_rows = student_rows[student_rows[date_col] == top_date]
            if len(same_date_rows) > 1:
                selections = set()
                for idx, r in same_date_rows.iterrows():
                    s_val = str(r[selection_col]).lower().strip() if selection_col else ""
                    selections.add(s_val)
                if len(selections) > 1:
                    err_row = same_date_rows.iloc[0].copy()
                    err_row['error_reason'] = f"Conflicting selections {selections} on same date"
                    error_rows.append(err_row)
                    continue

        if selection_col:
            sel = top_row[selection_col]
            if pd.isna(sel) or str(sel).strip().lower() not in ['a', 'b', 'c', 'd']:
                err_row = top_row.copy()
                err_row['error_reason'] = f"Invalid Selection: '{sel}'"
                error_rows.append(err_row)
                continue

        cleaned_rows.append(top_row)

    return cleaned_rows, error_rows


def validate_data():
    print("--- Starting Data Validation ---")
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    reports_dir = "reports"
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
    report_file = os.path.join(reports_dir, f"yearbook-choice-errors-{timestamp}.xlsx")
    session_info_path = os.path.join(os.path.dirname(__file__), "current_session.txt")
    try:
        with open(session_info_path, "w") as f:
            f.write(report_file)
    except Exception as e:
        print(f"Warning: Could not save session info: {e}")

    excel_path = get_excel_path()
    if not excel_path:
        sys.exit(1)
    print(f"Checking file: {excel_path}")

    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Critical Error: Could not read Excel file. {e}")
        sys.exit(1)

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
    print(f"Processing {len(df[student_id_col].unique())} unique Student IDs...")

    cleaned_rows, error_rows = _process_validation(df, student_id_col, selection_col, date_col, last_name_col)

    # 4. Save Outputs
    print(f"\nProcessing Complete.")

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
    write_error_report(error_rows, report_file)
    if error_rows:
        print(f"-> Error report started: {report_file}")
        
    if not cleaned_rows:
         sys.exit(1) # Fail if nothing to run
    else:
         print("-> Ready for automation.")

if __name__ == "__main__":
    validate_data()
