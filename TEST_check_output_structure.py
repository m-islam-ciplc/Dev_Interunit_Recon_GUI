import pandas as pd

def check_output_structure():
    """Check the exact structure of output files to find where LC numbers are stored"""
    
    print("=== CHECKING OUTPUT FILE STRUCTURE ===")
    
    try:
        df1 = pd.read_excel('Output/Interunit GeoTex_MATCHED.xlsx', header=8)
        df2 = pd.read_excel('Output/Interunit Steel_MATCHED.xlsx', header=8)
        print("✓ Output files loaded successfully")
    except Exception as e:
        print(f"✗ Error loading output files: {e}")
        return
    
    print(f"\n=== COLUMN STRUCTURE ===")
    print(f"File 1 columns: {list(df1.columns)}")
    print(f"File 2 columns: {list(df2.columns)}")
    
    print(f"\n=== SAMPLE DATA FROM FIRST FEW ROWS ===")
    
    # Show first 5 rows with Match IDs
    mid_col1 = df1.iloc[:, 0]  # Match ID column
    mid_rows1 = mid_col1.dropna()
    
    print(f"\nFile 1 - First 5 rows with Match IDs:")
    for i, (idx, match_id) in enumerate(mid_rows1.head().items()):
        if i < 5:
            print(f"\nRow {idx} - Match ID: '{match_id}'")
            print(f"  All columns:")
            for col_idx, col_name in enumerate(df1.columns):
                value = df1.iloc[idx, col_idx]
                if pd.notna(value) and str(value).strip():
                    print(f"    {col_name}: '{value}'")
    
    # Also check the original input files to see where LC numbers are
    print(f"\n=== CHECKING INPUT FILES FOR LC NUMBERS ===")
    
    try:
        df1_input = pd.read_excel('Input Files/Interunit GeoTex.xlsx', header=8)
        df2_input = pd.read_excel('Input Files/Interunit Steel.xlsx', header=8)
        
        print(f"Input File 1 columns: {list(df1_input.columns)}")
        print(f"Input File 2 columns: {list(df2_input.columns)}")
        
        # Look for LC numbers in different columns
        import re
        
        def find_lc_in_column(df, col_idx, col_name):
            lc_found = []
            for row_idx in range(min(10, len(df))):  # Check first 10 rows
                value = df.iloc[row_idx, col_idx]
                if pd.notna(value):
                    lc_matches = re.findall(r'L/C-\d+', str(value))
                    if lc_matches:
                        lc_found.append((row_idx, lc_matches[0]))
            return lc_found
        
        print(f"\nSearching for LC numbers in input files:")
        
        # Check each column in input files
        for col_idx, col_name in enumerate(df1_input.columns):
            lc_found = find_lc_in_column(df1_input, col_idx, col_name)
            if lc_found:
                print(f"  File 1 - Column '{col_name}' (idx {col_idx}): Found LC numbers in rows {lc_found[:3]}")
        
        for col_idx, col_name in enumerate(df2_input.columns):
            lc_found = find_lc_in_column(df2_input, col_idx, col_name)
            if lc_found:
                print(f"  File 2 - Column '{col_name}' (idx {col_idx}): Found LC numbers in rows {lc_found[:3]}")
        
    except Exception as e:
        print(f"Error reading input files: {e}")

if __name__ == "__main__":
    check_output_structure()
