import os
import sys
import pandas as pd
from excel_utils import get_excel_path, find_column_robust

def validate():
    # Look in the PARENT directory (../) relative to this script
    # This script is in /.../package-choice/code-package-choice/
    # We want /.../package-choice/
    
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    print(f"Validating input files in: {base_dir}")
    
    excel_path = get_excel_path(base_dir)
    
    import glob
    if not excel_path:
        # get_excel_path prints an error if > 1 file is found, but silent if 0.
        # We only want to print our error if there are genuinely 0 files.
        if not glob.glob(os.path.join(base_dir, "*.xlsx")):
            print("\n[ERROR] No Excel file found in the package-choice folder!")
            print("Please place your Input Excel file in the 'package-choice' folder.")
        sys.exit(1)
        
    print(f"Found Input File: {os.path.basename(excel_path)}")
    
    # helper for multiple file check is inside get_excel_path already, 
    # but let's double check if get_excel_path returns None if multiple are found?
    # Yes, my implementation returns None and prints error if > 1.
    
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"\n[ERROR] Could not read Excel file: {e}")
        sys.exit(1)
        
    # Check Columns
    # Required: ID, Product
    id_col = find_column_robust(df, "student id")
    product_col = find_column_robust(df, ["product name", "package choice"])
    
    missing = []
    if not id_col: missing.append("Student ID")
    if not product_col: missing.append("Package Choice / Product Name")
    
    if missing:
        print(f"\n[ERROR] Missing required columns: {', '.join(missing)}")
        print("Please check your Excel file headers.")
        sys.exit(1)

    print("\n[SUCCESS] Input file is valid and ready for processing.")
    sys.exit(0)

if __name__ == "__main__":
    validate()
