
import pandas as pd
import glob

files = glob.glob("*.xlsx")
if not files:
    print("No Excel files found")
else:
    f = files[0]
    print(f"Reading {f}...")
    df = pd.read_excel(f)
    print("Columns found:")
    for c in df.columns:
        print(f" - '{c}'")
