import pandas as pd

# Check the output file
df = pd.read_excel('Output/Interunit Steel_MATCHED.xlsx', header=8)
print(f"File shape: {df.shape}")
print(f"Columns: {list(df.columns)}")

# Find rows with Match IDs
match_rows = df[df.iloc[:, 0].notna()]
print(f"\nFound {len(match_rows)} rows with Match IDs")

# Check if transaction blocks are expanded
print("\n=== CHECKING TRANSACTION BLOCK EXPANSION ===")
for i, (idx, row) in enumerate(match_rows.head(5).iterrows()):
    match_id = row.iloc[0]
    print(f"\nMatch ID {match_id} at row {idx}:")
    print(f"  Date: {row.iloc[2]}")
    print(f"  Particulars: {row.iloc[3]}")
    print(f"  Audit Info: {str(row.iloc[1])[:100]}...")
    
    # Check if this Match ID appears in multiple consecutive rows
    consecutive_count = 0
    current_idx = idx
    while current_idx < len(df) and df.iloc[current_idx, 0] == match_id:
        consecutive_count += 1
        current_idx += 1
    
    print(f"  This Match ID appears in {consecutive_count} consecutive rows")
    
    if consecutive_count > 1:
        print(f"  Transaction block spans rows {idx} to {idx + consecutive_count - 1}")
    else:
        print(f"  Only 1 row - transaction block NOT expanded!")
