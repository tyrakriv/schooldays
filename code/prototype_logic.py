import pandas as pd
from datetime import datetime

# Mock data reflecting the screenshot
data = {
    'Student ID': [187259, 187259, 187259, 180883, 175738, 175738],
    'First Name': ['Stephani', 'Stephani', 'Stephani', 'Leon', 'Habiba', 'Habiba'],
    'Last Name': ['Abrokwa', 'Abrokwa', 'Abrokwa', 'Alba', 'Arafat', 'Arafat'],
    'Yearbook Selection': ['a', 'a', 'a', 'a', 'c', 'c'],
    'Yearbook Date': [
        '10/6/2025 18:57',
        '10/6/2025 22:21', # Latest for 187259
        '10/6/2025 21:00', # Hypothetical middle
        '10/9/2025 17:39',
        '10/7/2025 20:06',
        '10/7/2025 20:07'  # Latest for 175738
    ]
}

def get_latest_student_selection(df, student_id):
    """
    Filters for student_id, sorts by Yearbook Date, and returns the selection char.
    """
    # 1. Filter by Student ID
    student_rows = df[df['Student ID'] == student_id].copy()
    
    if student_rows.empty:
        return None, "Student not found"
    
    # 2. Convert Date to datetime if not already
    student_rows['Yearbook Date'] = pd.to_datetime(student_rows['Yearbook Date'])
    
    # 3. Sort by Date Descending
    student_rows = student_rows.sort_values(by='Yearbook Date', ascending=False)
    
    # 4. Pick top row
    latest_row = student_rows.iloc[0]
    selection = latest_row['Yearbook Selection']
    
    return selection, latest_row

def map_selection_to_index(selection_char):
    """
    Maps 'a' -> 0, 'b' -> 1, 'c' -> 2, etc.
    """
    char = selection_char.lower()
    return ord(char) - ord('a')

# --- Execution ---
df = pd.DataFrame(data)
target_id = 187259

print(f"Looking up Student ID: {target_id}")
selection, details = get_latest_student_selection(df, target_id)

if selection:
    index = map_selection_to_index(selection)
    print(f"Found latest entry from: {details['Yearbook Date']}")
    print(f"Selection char: '{selection}'")
    print(f"Computed List Index to Select: {index} (Where 0 is the first item)")
    
    # Explanation for User
    print("-" * 30)
    print(f"Logic: Char '{selection}' corresponds to index {index}.")
    print(f"Automation would select the item at position {index} in the list.")
else:
    print("Student not found.")
