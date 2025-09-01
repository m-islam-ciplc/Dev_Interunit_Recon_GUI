#!/usr/bin/env python3
"""
Test the specific aggregated PO scenarios found
"""

import pandas as pd
import re

# Read the input files
print("Reading input files...")
df1 = pd.read_excel('Input Files/Pole Book STEEL.xlsx', header=None)  # File 1
df2 = pd.read_excel('Input Files/Steel Book POLE.xlsx', header=None)  # File 2

print(f"File 1 (Pole Book STEEL): {df1.shape}")
print(f"File 2 (Steel Book POLE): {df2.shape}")

# PO Pattern for extraction
PO_PATTERN = r'(?:^|\s)([A-Z0-9/]+/PO/[A-Z0-9/]+)(?:\s|$|[,\.])'

print(f"\n=== TESTING SPECIFIC AGGREGATED PO SCENARIOS ===")

# Function to find transaction block header
def find_transaction_block_header(description_row_idx, df):
    """Find the transaction block header row for a given description row."""
    for row_idx in range(description_row_idx, -1, -1):
        row = df.iloc[row_idx]
        
        has_date = pd.notna(row.iloc[0]) and str(row.iloc[0]).strip() != ''
        has_debit = pd.notna(row.iloc[7]) and row.iloc[7] != 0
        has_credit = pd.notna(row.iloc[8]) and row.iloc[8] != 0
        
        if has_date and (has_debit or has_credit):
            return row_idx
    
    return description_row_idx

# Test Scenario 1: File 1 Row 35 (2 POs) -> File 2 (borrower transactions)
print(f"\n=== SCENARIO 1: File 1 Row 35 (2 POs) ===")
narration1 = str(df1.iloc[35, 2])
pos1 = re.findall(PO_PATTERN, narration1.upper())
print(f"Lender Narration: {narration1[:100]}...")
print(f"POs found: {pos1}")
print(f"PO Count: {len(pos1)}")

# Find transaction block header
header1 = find_transaction_block_header(35, df1)
lender_amount = df1.iloc[header1, 7]
print(f"Lender Transaction Header: Row {header1}")
print(f"Lender Amount (Debit): {lender_amount}")

# Now search for these POs in File 2 as borrower transactions
print(f"\nSearching for borrower transactions in File 2...")
borrower_matches = []
total_borrower_amount = 0

for i in range(8, len(df2)):
    narration2 = str(df2.iloc[i, 2])
    po_matches = re.findall(PO_PATTERN, narration2.upper())
    
    for po in pos1:
        if po in po_matches:
            # Found a matching PO, check if it's a borrower transaction
            header2 = find_transaction_block_header(i, df2)
            debit2 = df2.iloc[header2, 7] if pd.notna(df2.iloc[header2, 7]) else 0
            credit2 = df2.iloc[header2, 8] if pd.notna(df2.iloc[header2, 8]) else 0
            
            if credit2 > 0:  # Borrower transaction (credit > 0)
                borrower_matches.append({
                    'row': i,
                    'header_row': header2,
                    'po': po,
                    'amount': credit2,
                    'narration': narration2[:100]
                })
                total_borrower_amount += credit2
                print(f"  ✅ Found borrower PO {po}: Amount {credit2} (Row {header2})")
            elif debit2 > 0:
                print(f"  ⚠️  Found lender PO {po}: Amount {debit2} (Row {header2}) - Not a borrower")

print(f"\nBorrower Matches Found: {len(borrower_matches)}")
print(f"Total Borrower Amount: {total_borrower_amount}")
print(f"Lender Amount: {lender_amount}")
print(f"Amount Match: {'✅ YES' if abs(lender_amount - total_borrower_amount) < 0.01 else '❌ NO'}")

# Check if all POs are found
found_pos = [match['po'] for match in borrower_matches]
missing_pos = [po for po in pos1 if po not in found_pos]
print(f"PO Coverage: {'✅ COMPLETE' if not missing_pos else f'❌ INCOMPLETE - Missing: {missing_pos}'}")

if not missing_pos and abs(lender_amount - total_borrower_amount) < 0.01:
    print(f"\n🎉 AGGREGATED PO MATCH FOUND!")
    print(f"Match Type: Aggregated PO")
    print(f"PO Count: {len(pos1)}")
    print(f"Lender Amount: {lender_amount}")
    print(f"Total Borrower Amount: {total_borrower_amount}")
    print(f"Borrower Transactions: {len(borrower_matches)}")
else:
    print(f"\n❌ No valid aggregated PO match found")

# Test Scenario 2: File 2 Row 26 (3 POs) -> File 1 (borrower transactions)
print(f"\n" + "="*60)
print(f"=== SCENARIO 2: File 2 Row 26 (3 POs) ===")
narration2 = str(df2.iloc[26, 2])
pos2 = re.findall(PO_PATTERN, narration2.upper())
print(f"Lender Narration: {narration2[:100]}...")
print(f"POs found: {pos2}")
print(f"PO Count: {len(pos2)}")

# Find transaction block header
header2 = find_transaction_block_header(26, df2)
lender_amount2 = df2.iloc[header2, 7]
print(f"Lender Transaction Header: Row {header2}")
print(f"Lender Amount (Debit): {lender_amount2}")

# Search for these POs in File 1 as borrower transactions
print(f"\nSearching for borrower transactions in File 1...")
borrower_matches2 = []
total_borrower_amount2 = 0

for i in range(8, len(df1)):
    narration1 = str(df1.iloc[i, 2])
    po_matches = re.findall(PO_PATTERN, narration1.upper())
    
    for po in pos2:
        if po in po_matches:
            # Found a matching PO, check if it's a borrower transaction
            header1 = find_transaction_block_header(i, df1)
            debit1 = df1.iloc[header1, 7] if pd.notna(df1.iloc[header1, 7]) else 0
            credit1 = df1.iloc[header1, 8] if pd.notna(df1.iloc[header1, 8]) else 0
            
            if credit1 > 0:  # Borrower transaction (credit > 0)
                borrower_matches2.append({
                    'row': i,
                    'header_row': header1,
                    'po': po,
                    'amount': credit1,
                    'narration': narration1[:100]
                })
                total_borrower_amount2 += credit1
                print(f"  ✅ Found borrower PO {po}: Amount {credit1} (Row {header1})")
            elif debit1 > 0:
                print(f"  ⚠️  Found lender PO {po}: Amount {debit1} (Row {header1}) - Not a borrower")

print(f"\nBorrower Matches Found: {len(borrower_matches2)}")
print(f"Total Borrower Amount: {total_borrower_amount2}")
print(f"Lender Amount: {lender_amount2}")
print(f"Amount Match: {'✅ YES' if abs(lender_amount2 - total_borrower_amount2) < 0.01 else '❌ NO'}")

# Check if all POs are found
found_pos2 = [match['po'] for match in borrower_matches2]
missing_pos2 = [po for po in pos2 if po not in found_pos2]
print(f"PO Coverage: {'✅ COMPLETE' if not missing_pos2 else f'❌ INCOMPLETE - Missing: {missing_pos2}'}")

if not missing_pos2 and abs(lender_amount2 - total_borrower_amount2) < 0.01:
    print(f"\n🎉 AGGREGATED PO MATCH FOUND!")
    print(f"Match Type: Aggregated PO")
    print(f"PO Count: {len(pos2)}")
    print(f"Lender Amount: {lender_amount2}")
    print(f"Total Borrower Amount: {total_borrower_amount2}")
    print(f"Borrower Transactions: {len(borrower_matches2)}")
else:
    print(f"\n❌ No valid aggregated PO match found")
