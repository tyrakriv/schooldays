import glob
import os

def find_column_robust(dataframe, keywords):
    """
    Helper to find a column by keyword (case-insensitive).
    """
    if isinstance(keywords, str):
        keywords = [keywords]
        
    for col in dataframe.columns:
        col_lower = str(col).strip().lower()
        for kw in keywords:
            if kw.lower() in col_lower:
                 return col
    return None

def get_excel_path(directory="."):
    """
    Finds the first .xlsx file in the specified directory.
    """
    search_path = os.path.join(directory, "*.xlsx")
    excel_files = glob.glob(search_path)
    
    if not excel_files:
        return None
        
    return excel_files[0]
