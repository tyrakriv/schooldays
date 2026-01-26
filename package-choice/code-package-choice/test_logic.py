import pandas as pd
from data_handler_package import map_product_to_code

def test_mappings():
    print("--- TESTING PRODUCT MAPPING ---")
    
    test_inputs = [
        "3x5â€™s Package",
        "5x7â€™s Package",
        "8x10 Package",
        "Economy Package",
        "Deluxe Package",
        "Ultimate Package",
        "Classic package",
        "Basic Package",
        "Mini Wallets Package",
        "Wallets Package",
        "5â€ x 7â€ (127 x 178 mm) Group Print",
        "8â€ x 10â€ (203 x 254 mm) Group Print",
        "All 4 digital portraits in Hi-Resolution jpg format",
        "Touch Up Photos",
        "Lost Order Form",
        "No Photo Package Wanted"
    ]
    
    print(f"{'INPUT':<55} | {'CODE':<5} | {'TYPE':<10}")
    print("-" * 80)
    
    for inp in test_inputs:
        code, p_type, _ = map_product_to_code(inp)
        print(f"{inp:<55} | {str(code):<5} | {p_type:<10}")

if __name__ == "__main__":
    test_mappings()
