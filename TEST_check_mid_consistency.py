import pandas as pd
import numpy as np

def check_mid_consistency():
    """Check if Match IDs are consistent across both output files"""
    
    print("=== CHECKING MATCH ID CONSISTENCY ===")
    
    # Read both output files
    try:
        df1 = pd.read_excel('Output/Interunit GeoTex_MATCHED.xlsx', header=8)
        df2 = pd.read_excel('Output/Interunit Steel_MATCHED.xlsx', header=8)
        print("✓ Both files loaded successfully")
    except Exception as e:
        print(f"✗ Error loading files: {e}")
        return
    
    print(f"\n=== FILE OVERVIEW ===")
    print(f"File 1 (GeoTex): {df1.shape[0]} rows, {df1.shape[1]} columns")
    print(f"File 2 (Steel): {df2.shape[0]} rows, {df2.shape[1]} columns")
    
    # Check column names
    print(f"\n=== COLUMN NAMES ===")
    print(f"File 1 columns: {list(df1.columns)}")
    print(f"File 2 columns: {list(df2.columns)}")
    
    # Check Match ID column (first column)
    mid_col1 = df1.iloc[:, 0]
    mid_col2 = df2.iloc[:, 0]
    
    print(f"\n=== MATCH ID ANALYSIS ===")
    print(f"File 1 - Match ID column name: '{df1.columns[0]}'")
    print(f"File 2 - Match ID column name: '{df2.columns[0]}'")
    
    # Count non-empty Match IDs
    non_empty_mids1 = mid_col1.dropna()
    non_empty_mids2 = mid_col2.dropna()
    
    print(f"File 1 - Rows with Match IDs: {len(non_empty_mids1)}")
    print(f"File 2 - Rows with Match IDs: {len(non_empty_mids2)}")
    
    # Show unique Match IDs in each file
    unique_mids1 = non_empty_mids1.unique()
    unique_mids2 = non_empty_mids2.unique()
    
    print(f"\nFile 1 - Unique Match IDs: {sorted(unique_mids1)}")
    print(f"File 2 - Unique Match IDs: {sorted(unique_mids2)}")
    
    # Check if Match IDs are consistent (same unique values)
    mids_consistent = set(unique_mids1) == set(unique_mids2)
    print(f"\n✓ Match IDs are consistent across files: {mids_consistent}")
    
    if not mids_consistent:
        print(f"✗ File 1 only: {set(unique_mids1) - set(unique_mids2)}")
        print(f"✗ File 2 only: {set(unique_mids2) - set(unique_mids1)}")
    
    # Check for empty strings vs NaN
    empty_strings1 = (mid_col1 == '').sum()
    empty_strings2 = (mid_col2 == '').sum()
    nan_count1 = mid_col1.isna().sum()
    nan_count2 = mid_col2.isna().sum()
    
    print(f"\n=== EMPTY VALUES ANALYSIS ===")
    print(f"File 1 - Empty strings: {empty_strings1}, NaN values: {nan_count1}")
    print(f"File 2 - Empty strings: {empty_strings2}, NaN values: {nan_count2}")
    
    # Show first few rows with actual Match IDs
    print(f"\n=== SAMPLE ROWS WITH MATCH IDS ===")
    
    # File 1 samples
    print(f"\nFile 1 - First 5 rows with Match IDs:")
    for i, (idx, value) in enumerate(non_empty_mids1.head().items()):
        if i < 5:
            print(f"  Row {idx}: '{value}'")
    
    # File 2 samples  
    print(f"\nFile 2 - First 5 rows with Match IDs:")
    for i, (idx, value) in enumerate(non_empty_mids2.head().items()):
        if i < 5:
            print(f"  Row {idx}: '{value}'")
    
    # Check if Match IDs follow expected pattern (M001, M002, etc.)
    print(f"\n=== MATCH ID PATTERN CHECK ===")
    pattern_matches1 = [mid for mid in unique_mids1 if str(mid).startswith('M')]
    pattern_matches2 = [mid for mid in unique_mids2 if str(mid).startswith('M')]
    
    print(f"File 1 - Match IDs following 'M' pattern: {len(pattern_matches1)}/{len(unique_mids1)}")
    print(f"File 2 - Match IDs following 'M' pattern: {len(pattern_matches2)}/{len(unique_mids2)}")
    
    if pattern_matches1:
        print(f"  File 1 pattern examples: {sorted(pattern_matches1)[:5]}")
    if pattern_matches2:
        print(f"  File 2 pattern examples: {sorted(pattern_matches2)[:5]}")

if __name__ == "__main__":
    check_mid_consistency()
