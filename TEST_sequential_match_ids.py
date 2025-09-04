#!/usr/bin/env python3
"""Test script to verify sequential Match IDs."""

import pandas as pd
import re
import os

def check_sequential_match_ids(file_path):
    """Check if Match IDs in the output file are sequential."""
    print(f"\nChecking Match IDs in: {file_path}")
    
    # Read the Excel file (skip metadata rows)
    df = pd.read_excel(file_path, header=8)
    
    # Get all unique Match IDs (excluding None/NaN)
    match_ids = df.iloc[:, 0].dropna().unique()
    
    if len(match_ids) == 0:
        print("No Match IDs found in the file.")
        return True
    
    # Extract the numeric part of Match IDs
    match_numbers = []
    for match_id in match_ids:
        if isinstance(match_id, str) and match_id.startswith('M'):
            try:
                num = int(match_id[1:])
                match_numbers.append((num, match_id))
            except ValueError:
                print(f"Warning: Invalid Match ID format: {match_id}")
    
    # Sort by numeric value
    match_numbers.sort(key=lambda x: x[0])
    
    # Check for sequential order
    is_sequential = True
    expected = 1
    
    print(f"\nFound {len(match_numbers)} unique Match IDs:")
    for num, match_id in match_numbers:
        print(f"  {match_id} (numeric: {num})")
        if num != expected:
            print(f"    ❌ Expected M{expected:03d}, but got {match_id}")
            is_sequential = False
        expected = num + 1
    
    if is_sequential:
        print("\n✅ Match IDs are sequential!")
    else:
        print("\n❌ Match IDs are NOT sequential!")
    
    # Show Match Type distribution
    if 'Match Type' in df.columns or df.shape[1] > 12:  # Last column should be Match Type
        match_types = df.iloc[:, -1].dropna()
        if len(match_types) > 0:
            print("\nMatch Type distribution:")
            for match_type, count in match_types.value_counts().items():
                print(f"  {match_type}: {count} rows")
    
    return is_sequential

def main():
    """Main function to check all output files."""
    output_folder = "Output"
    
    # Find all matched Excel files in the output folder
    matched_files = []
    if os.path.exists(output_folder):
        for file in os.listdir(output_folder):
            if file.endswith("_MATCHED.xlsx"):
                matched_files.append(os.path.join(output_folder, file))
    
    if not matched_files:
        print(f"No matched files found in {output_folder}")
        return
    
    print(f"Found {len(matched_files)} matched files to check:")
    
    all_sequential = True
    for file_path in matched_files:
        if not check_sequential_match_ids(file_path):
            all_sequential = False
    
    print("\n" + "="*60)
    if all_sequential:
        print("✅ ALL FILES HAVE SEQUENTIAL MATCH IDs!")
    else:
        print("❌ SOME FILES HAVE NON-SEQUENTIAL MATCH IDs!")
    print("="*60)

if __name__ == "__main__":
    main()