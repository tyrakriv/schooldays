import pandas as pd
import os
from excel_utils import find_column_robust, get_excel_path
from collections import defaultdict

def normalize_text(text):
    if pd.isna(text):
        return ""
    return str(text).strip().lower()

def map_product_to_code(product_name):
    """
    Maps a product name to (code, type, raw_name).
    Type can be: 'standard', 'group', 'cd', 'touchup', 'unknown', 'ignore'
    """
    p = normalize_text(product_name)
    
    # Ignore list
    if p in ["no photo package wanted", ""]:
        return None, 'ignore', p

    # Standard Packages (Quick Package Entry)
    # We use 'in' to handle messy inputs like "3x5â€™s Package"
    if "mini wallet" in p: return "m", 'standard', p # Check mini before wallet
    if "wallets" in p or "wallet prints" in p: return "w", 'standard', p
    
    # Group Prints Checking First to avoid confusion with standards
    if "group print" in p:
        if "5" in p and "7" in p: # Matches "5â€ x 7â€ ... Group Print"
            return "m", 'group', p
        elif "8" in p and "10" in p: # Matches "8â€ x 10â€ ... Group Print"
            return "l", 'group', p
        else:
            return None, 'unknown', p # 3x5 group print dropped

    # Now valid standards that might overlap numbers
    if "3x5" in p or "3 x 5" in p: return "f", 'standard', p
    if "5x7" in p or "5 x 7" in p: return "s", 'standard', p
    if "8x10" in p or "8 x 10" in p: return "t", 'standard', p
    
    if "basic" in p: return "b", 'standard', p
    if "classic" in p: return "c", 'standard', p
    if "deluxe" in p: return "d", 'standard', p
    if "economy" in p: return "e", 'standard', p
    if "ultimate" in p: return "u", 'standard', p

    # CD
    if "digital" in p and "portraits" in p: return "CD", 'cd', p # "All 4 digital portraits..."
    if "cd" in p: return "CD", 'cd', p

    # Touchup
    if "touch up" in p:
        return "Pending", 'touchup', p

    return None, 'unknown', p


def load_and_process_data(excel_path=None):
    """
    Reads the Excel file and processes students and their packages.
    """
    if not excel_path:
        # Look in the PARENT directory package-choice/
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        excel_path = get_excel_path(base_dir)
    
    if not excel_path or not os.path.exists(excel_path):
        print("Error: No Input Excel file found (or multiple found).")
        return []

    print(f"Loading data from: {excel_path}")
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return []

    processed_data = []

    # Identify columns
    id_col = find_column_robust(df, "student id")
    choice_col = find_column_robust(df, ["photo choice", "yearbook choice"])
    product_col = find_column_robust(df, ["product name", "package choice", "description"])
    qty_col = find_column_robust(df, ["quantity", "qty"])
    last_name_col = find_column_robust(df, ["last name", "student last name"])

    if not id_col or not product_col:
        print(f"Error: Missing required columns. Found ID: {id_col}, Product: {product_col}")
        return []
        
    if not choice_col:
        print("Warning: 'Photo Choice' column not found. Defaulting to 'a' if needed?")
        
    # Group by Student ID
    df['normalized_id'] = df[id_col].astype(str).str.strip()
    unique_ids = df['normalized_id'].unique()
    
    for sid in unique_ids:
        student_rows = df[df['normalized_id'] == sid]
        
        last_name = ""
        if last_name_col and not student_rows.empty:
             val = student_rows.iloc[0][last_name_col]
             if pd.notna(val):
                 last_name = str(val).strip()

        # Track duplicates: (choice, raw_product) -> count
        seen_entries = defaultdict(int)
        
        # We need to process grouping by CHOICES
        # Structure: choice -> { 'standard_string': "", 'others': [], 'has_personal': False }
        choices_map = {} 
        
        # List of errors for this student
        student_errors = []

        for idx, row in student_rows.iterrows():
            raw_product = row[product_col]
            
            # 1. Get Quantity
            qty = 1
            if qty_col and pd.notna(row[qty_col]):
                try:
                    qty = int(float(row[qty_col]))
                except:
                    qty = 1 # Default or Error? Assuming default 1 if parse fails, or maybe log?
            
            # 2. Get Photo Choice
            photo_choice = None # No longer defaulting to 'a'
            if choice_col and pd.notna(row[choice_col]):
                photo_choice = str(row[choice_col]).strip().lower()
            
            # Use a placeholder for grouping if None
            # logic: if photo_choice is None, we still process it (likely a Group Print only order)
            group_key = photo_choice if photo_choice else "NO_SELECTION" 

            # Duplicate Check Key
            dup_key = (group_key, str(raw_product).strip().lower())
            seen_entries[dup_key] += 1
            if seen_entries[dup_key] > 1:
                student_errors.append({
                    'raw_product': raw_product,
                    'reason': "Duplicate Line detected in Excel"
                })
                continue

            # Process Product
            code, p_type, raw_name = map_product_to_code(raw_product)
            
            if p_type == 'ignore':
                continue
            
            # Initialize choice group if new
            if group_key not in choices_map:
                choices_map[group_key] = {
                    'real_choice': photo_choice, # processing key
                    'standard_string': "", 
                    'others': [], 
                    'has_personal': False,
                    'group_print_types': set(),
                    'group_codes': ""
                }
            
            grp = choices_map[group_key]

            if p_type == 'standard':
                # Handle Quantity for standard -> Repeat string
                # e.g. code='f', qty=2 -> 'ff'
                if qty < 1: qty = 1 # Safety
                grp['standard_string'] += (code * qty)
                grp['has_personal'] = True
            
            elif p_type in ['group', 'cd', 'touchup']:
                # Quantity Check: Must be 1 for CD/Touchup usually? 
                # User specifically asked for Group Print quantity support (e.g. 2x 5x7 -> mm).
                
                if p_type != 'group' and qty > 1:
                     student_errors.append({
                        'raw_product': raw_product,
                        'reason': f"Quantity {qty} not allowed for {p_type}"
                    })
                     continue
                
                # Check Group Print Limit (>2 different types)
                if p_type == 'group':
                    grp['group_print_types'].add(code)
                    if len(grp['group_print_types']) > 2:
                        student_errors.append({
                           'raw_product': raw_product,
                           'reason': "More than 2 different Group Print types selected"
                        })
                        continue
                    
                    # Accumulate group code
                    # Handle Quantity (e.g. 2x -> 'mm')
                    grp['group_codes'] += (code * qty)
                    continue # Do NOT add to 'others' yet, we will add combined at end

                # Add to 'others' list (CD, Touchup)
                # Check for duplicates (ONLY 1 CD or Touchup allowed per choice group)
                existing_type = next((x for x in grp['others'] if x['type'] == p_type), None)
                if existing_type:
                     student_errors.append({
                        'raw_product': raw_product,
                        'reason': f"Multiple {p_type.upper()} items selected (Only 1 allowed)"
                    })
                     continue

                grp['others'].append({
                    'code': code,
                    'type': p_type,
                    'raw_product': raw_product
                })
                
            elif p_type == 'unknown':
                student_errors.append({
                    'raw_product': raw_product,
                    'reason': "Unknown Product Code (Not 5x7 or 8x10 Group, or recognized pkg)"
                })

        # Post-process: Assign target boxes for 'others' (Group logic)
        final_choices = []
        for key, data in choices_map.items():
            
            # If we have accumulated group codes, add them as a single item now
            if data.get('group_codes'):
                data['others'].append({
                    'code': data['group_codes'],
                    'type': 'group',
                    'raw_product': 'Combined Group Prints'
                })

            # Process 'others' to resolve Group Print box location
            processed_others = []
            for item in data['others']:
                p_type = item['type']
                target_box = None
                
                if p_type == 'cd':
                    target_box = 'cd_box'
                elif p_type == 'touchup':
                    target_box = 'touchup'
                elif p_type == 'group':
                    if data['has_personal']:
                        target_box = 'class_pkg_box'
                    else:
                        target_box = 'class_pix_no_pkg_box'
                
                item['target_box'] = target_box
                processed_others.append(item)
            
            final_choices.append({
                'photo_choice': data['real_choice'], # Can be None
                'standard_string': data['standard_string'],
                'others': processed_others
            })

        processed_data.append({
            'id': sid,
            'last_name': last_name,
            'choices_groups': final_choices,
            'errors': student_errors
        })
        
    return processed_data
