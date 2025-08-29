import pandas as pd
import openpyxl

def debug_lc_matching():
    """Debug why the same LC numbers are getting different Match IDs"""
    
    print("=== DEBUGGING LC MATCHING ===")
    
    # Read input files to see what LC numbers exist
    try:
        # Read the input files (not the output files)
        df1_input = pd.read_excel('Input Files/Interunit GeoTex.xlsx', header=8)
        df2_input = pd.read_excel('Input Files/Interunit Steel.xlsx', header=8)
        print("✓ Input files loaded successfully")
    except Exception as e:
        print(f"✗ Error loading input files: {e}")
        return
    
    print(f"\n=== INPUT FILE OVERVIEW ===")
    print(f"File 1 (GeoTex): {df1_input.shape[0]} rows, {df1_input.shape[1]} columns")
    print(f"File 2 (Steel): {df2_input.shape[0]} rows, {df2_input.shape[1]} columns")
    
    # Check if there are multiple transactions with the same LC number
    print(f"\n=== CHECKING FOR DUPLICATE LC NUMBERS ===")
    
    # Extract LC numbers from description column (assuming it's column 2 or 3)
    # Let me check the column structure first
    print(f"File 1 columns: {list(df1_input.columns)}")
    print(f"File 2 columns: {list(df2_input.columns)}")
    
    # Look for LC numbers in the description column
    desc_col1 = df1_input.iloc[:, 2]  # Column C (Description)
    desc_col2 = df2_input.iloc[:, 2]  # Column C (Description)
    
    # Simple LC number extraction (L/C- followed by numbers)
    import re
    
    def extract_lc_numbers(description_series):
        lc_numbers = []
        for desc in description_series:
            if pd.notna(desc):
                # Look for L/C- pattern
                matches = re.findall(r'L/C-\d+', str(desc))
                lc_numbers.extend(matches)
        return lc_numbers
    
    lc_numbers1 = extract_lc_numbers(desc_col1)
    lc_numbers2 = extract_lc_numbers(desc_col2)
    
    print(f"\nFile 1 - Found LC numbers: {sorted(set(lc_numbers1))}")
    print(f"File 2 - Found LC numbers: {sorted(set(lc_numbers2))}")
    
    # Check for duplicates within each file
    from collections import Counter
    
    lc_count1 = Counter(lc_numbers1)
    lc_count2 = Counter(lc_numbers2)
    
    print(f"\n=== DUPLICATE LC NUMBERS ANALYSIS ===")
    
    duplicates1 = {lc: count for lc, count in lc_count1.items() if count > 1}
    duplicates2 = {lc: count for lc, count in lc_count2.items() if count > 1}
    
    if duplicates1:
        print(f"File 1 has duplicate LC numbers:")
        for lc, count in duplicates1.items():
            print(f"  {lc}: {count} times")
    else:
        print("File 1: No duplicate LC numbers found")
    
    if duplicates2:
        print(f"File 2 has duplicate LC numbers:")
        for lc, count in duplicates2.items():
            print(f"  {lc}: {count} times")
    else:
        print("File 2: No duplicate LC numbers found")
    
    # Check which LC numbers exist in both files
    common_lcs = set(lc_numbers1) & set(lc_numbers2)
    only_in_file1 = set(lc_numbers1) - set(lc_numbers2)
    only_in_file2 = set(lc_numbers2) - set(lc_numbers1)
    
    print(f"\n=== LC NUMBER OVERLAP ===")
    print(f"LC numbers in both files: {sorted(common_lcs)}")
    print(f"LC numbers only in File 1: {sorted(only_in_file1)}")
    print(f"LC numbers only in File 2: {sorted(only_in_file2)}")
    
    # Now check the output files to see what Match IDs were assigned
    print(f"\n=== CHECKING OUTPUT FILES ===")
    
    try:
        df1_output = pd.read_excel('Output/Interunit GeoTex_MATCHED.xlsx', header=8)
        df2_output = pd.read_excel('Output/Interunit Steel_MATCHED.xlsx', header=8)
        
        # Get Match IDs and LC numbers from output files
        mid_col1 = df1_output.iloc[:, 0]  # Match ID column
        mid_col2 = df2_output.iloc[:, 0]  # Match ID column
        
        desc_col1_out = df1_output.iloc[:, 3]  # Description column
        desc_col2_out = df2_output.iloc[:, 3]  # Description column
        
        # Find rows with Match IDs and extract LC numbers
        print(f"\n=== MATCH ID ANALYSIS ===")
        
        # File 1 analysis
        mid_rows1 = mid_col1.dropna()
        print(f"File 1 - Rows with Match IDs: {len(mid_rows1)}")
        
        for idx, match_id in mid_rows1.head(10).items():
            desc = desc_col1_out.iloc[idx] if idx < len(desc_col1_out) else "N/A"
            lc_matches = re.findall(r'L/C-\d+', str(desc))
            lc_num = lc_matches[0] if lc_matches else "No LC found"
            print(f"  Row {idx}: Match ID '{match_id}' -> LC: {lc_num}")
        
        # File 2 analysis
        mid_rows2 = mid_col2.dropna()
        print(f"\nFile 2 - Rows with Match IDs: {len(mid_rows2)}")
        
        for idx, match_id in mid_rows2.head(10).items():
            desc = desc_col2_out.iloc[idx] if idx < len(desc_col2_out) else "N/A"
            lc_matches = re.findall(r'L/C-\d+', str(desc))
            lc_num = lc_matches[0] if lc_matches else "No LC found"
            print(f"  Row {idx}: Match ID '{match_id}' -> LC: {lc_num}")
        
    except Exception as e:
        print(f"Error reading output files: {e}")

if __name__ == "__main__":
    debug_lc_matching()
