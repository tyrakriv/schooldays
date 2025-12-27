
import data_handler
import pandas as pd
import os

print("--- Verifying Data Loading Logic ---")

# Ensure we can load
data = data_handler.load_and_process_data(None)

print(f"Loaded {len(data)} students.")

# Check specific ID 187259 -> Should be 'b'
found_187259 = False
for s in data:
    if s['id'] == '187259':
        found_187259 = True
        print(f"Student 187259 Selection: '{s['selection']}'")
        if s['selection'] == 'b':
            print("PASS: Correct selection for 187259")
        else:
            print(f"FAIL: Expected 'b', got '{s['selection']}'")

if not found_187259:
    print("FAIL: Student 187259 not found in processed data.")
    
# Check for 'd' default overload
d_count = sum(1 for s in data if s['selection'] == 'd')
print(f"Total 'd' selections: {d_count}/{len(data)}")

if d_count == len(data) and len(data) > 0:
    print("WARNING: All selections are 'd'. Logic might still be broken if source has variation.")

print("--- verification complete ---")
