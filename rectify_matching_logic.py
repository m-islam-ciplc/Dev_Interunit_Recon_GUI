#!/usr/bin/env python3
"""
Rectify Matching Logic Module

This module implements the logic for finding rectify entry matches between two files based on:
1. Amounts must match exactly (no tolerance)
2. "Entered By" names must match
3. Both narrations must contain "Rectify entry"
4. Follows the same core logic and format as other matchers
"""

import pandas as pd

class RectifyMatchingLogic:
    """Handles the logic for finding rectify entry matches between two files."""
    
    def __init__(self):
        pass
    
    def find_potential_matches(self, transactions1, transactions2, existing_matches=None, match_counter=0):
        """Find potential rectify entry matches between the two files."""
        print(f"\nFile 1: {len(transactions1)} transactions")
        print(f"File 2: {len(transactions2)} transactions")
        
        # Find matches - RECTIFY LOGIC: Amount → Entered By → "Rectify entry" in narration
        matches = []
        
        # Use shared state if provided, otherwise create new
        if existing_matches is None:
            existing_matches = {}
        if match_counter is None:
            match_counter = 0
        
        print(f"\n=== RECTIFY MATCHING LOGIC ===")
        print(f"1. Check if amounts are EXACTLY the same")
        print(f"2. Check if 'Entered By' names are the same")
        print(f"3. Check if both narrations contain 'Rectify entry'")
        print(f"4. Only if all three match: Assign same Match ID")
        
        # Use shared state for tracking which combinations have already been matched
        # Key: (Amount, Entered_By), Value: match_id
        
        # FIRST: Pre-filter transactions that contain "Rectify entry" to avoid expensive processing
        print("Pre-filtering transactions with 'Rectify entry'...")
        
        # Get indices of transactions with "Rectify entry" in File 1
        file1_rectify_indices = []
        for idx1 in range(len(transactions1)):
            if self.check_rectify_in_narration(idx1, transactions1):
                file1_rectify_indices.append(idx1)
        
        # Get indices of transactions with "Rectify entry" in File 2  
        file2_rectify_indices = []
        for idx2 in range(len(transactions2)):
            if self.check_rectify_in_narration(idx2, transactions2):
                file2_rectify_indices.append(idx2)
        
        print(f"File 1: {len(file1_rectify_indices)} transactions with 'Rectify entry'")
        print(f"File 2: {len(file2_rectify_indices)} transactions with 'Rectify entry'")
        
        # Only process transactions that actually contain "Rectify entry"
        for idx1 in file1_rectify_indices:
            # Find the transaction block header row for this transaction in File 1
            block_header1 = self.find_transaction_block_header(idx1, transactions1)
            header_row1 = transactions1.iloc[block_header1]
            
            # Extract amounts and determine transaction type for File 1
            # Based on investigation: amounts are in columns 8 and 9 (iloc[7] and iloc[8])
            file1_debit = header_row1.iloc[7] if pd.notna(header_row1.iloc[7]) else 0
            file1_credit = header_row1.iloc[8] if pd.notna(header_row1.iloc[8]) else 0
            
            file1_is_lender = file1_debit > 0
            file1_is_borrower = file1_credit > 0
            file1_amount = file1_debit if file1_is_lender else file1_credit
            
            # Find "Entered By" name for this transaction block in File 1
            file1_entered_by = self.find_entered_by_name(block_header1, transactions1)
            
            # File 1 already has "Rectify entry" (we pre-filtered for this)
            file1_has_rectify = True
            
            # Now look for matches in File 2 (only those with "Rectify entry")
            for idx2 in file2_rectify_indices:
                # Find the transaction block header row for this transaction in File 2
                block_header2 = self.find_transaction_block_header(idx2, transactions2)
                header_row2 = transactions2.iloc[block_header2]
                
                # Extract amounts and determine transaction type for File 2
                # Based on investigation: amounts are in columns 8 and 9 (iloc[7] and iloc[8])
                file2_debit = header_row2.iloc[7] if pd.notna(header_row2.iloc[7]) else 0
                file2_credit = header_row2.iloc[8] if pd.notna(header_row2.iloc[8]) else 0
                
                file2_is_lender = file2_debit > 0
                file2_is_borrower = file2_credit > 0
                file2_amount = file2_debit if file2_is_lender else file2_credit
                
                # STEP 1: Check if amounts are EXACTLY the same
                if file1_amount != file2_amount:
                    continue
                
                # STEP 2: Check if transaction types are opposite (one lender, one borrower)
                if not ((file1_is_lender and file2_is_borrower) or (file1_is_borrower and file2_is_lender)):
                    continue
                
                # Find "Entered By" name for this transaction block in File 2
                file2_entered_by = self.find_entered_by_name(block_header2, transactions2)
                
                # STEP 3: Check if "Entered By" names are the same
                if file1_entered_by != file2_entered_by:
                    continue
                
                # STEP 4: File 2 already has "Rectify entry" (we pre-filtered for this)
                file2_has_rectify = True
                
                # STEP 5: Check if we already have a match for this combination
                match_key = (file1_amount, file1_entered_by)
                
                if match_key in existing_matches:
                    # Use existing Match ID for consistency
                    match_id = existing_matches[match_key]
                else:
                    # Create new Match ID
                    match_counter += 1
                    match_id = f"M{match_counter:03d}"
                    existing_matches[match_key] = match_id
                
                # Create the match
                matches.append({
                    'match_id': match_id,
                    'File1_Index': block_header1,
                    'File2_Index': block_header2,
                    'File1_Date': header_row1.iloc[0],
                    'File1_Description': header_row1.iloc[2],
                    'File1_Debit': header_row1.iloc[7],
                    'File1_Credit': header_row1.iloc[8],
                    'File2_Date': header_row2.iloc[0],
                    'File2_Description': header_row2.iloc[2],
                    'File2_Debit': header_row2.iloc[7],
                    'File2_Credit': header_row2.iloc[8],
                    'File1_Amount': file1_amount,
                    'File2_Amount': file2_amount,
                    'File1_Type': 'Lender' if file1_is_lender else 'Borrower',
                    'File2_Type': 'Lender' if file2_is_lender else 'Borrower',
                    'Entered_By': file1_entered_by,
                    'Match_Type': 'Rectify Entry'
                })
                
                # BREAK: Exit inner loop after finding a match for this File1 transaction
                # This prevents checking remaining File2 transactions for the same File1 transaction
                break
        
        print(f"\n=== RECTIFY MATCHING RESULTS ===")
        print(f"Found {len(matches)} valid rectify matches across {len(existing_matches)} unique Match ID combinations!")
        
        return matches
    
    def find_transaction_block_header(self, description_row_idx, transactions_df):
        """Find the transaction block header row for a given description row."""
        # Start from the description row and go backwards to find the block header
        # Block header is the row with date and particulars (Dr/Cr)
        for row_idx in range(description_row_idx, -1, -1):
            row = transactions_df.iloc[row_idx]
            
            # Check if this row has a date and particulars
            has_date = pd.notna(row.iloc[0]) and str(row.iloc[0]).strip() != ''
            has_particulars = pd.notna(row.iloc[1]) and str(row.iloc[1]).strip() != ''
            
            # Check if this row has either Debit or Credit amount (not both nan)
            # Based on investigation: amounts are in columns 8 and 9 (iloc[7] and iloc[8])
            has_debit = pd.notna(row.iloc[7]) and row.iloc[7] != 0
            has_credit = pd.notna(row.iloc[8]) and row.iloc[8] != 0
            
            # Transaction block header: has date, particulars, and either debit or credit
            if has_date and (has_debit or has_credit):
                return row_idx
        
        # If no header found, return the description row itself
        return description_row_idx
    
    def find_entered_by_name(self, block_header_idx, transactions_df):
        """Find the 'Entered By' name for a transaction block."""
        # Look forward from the block header to find "Entered By :" row
        for row_idx in range(block_header_idx, len(transactions_df)):
            row = transactions_df.iloc[row_idx]
            particulars = row.iloc[1]  # Column B (Particulars)
            
            if pd.notna(particulars) and str(particulars).strip() == 'Entered By :':
                # Found "Entered By :", get the name from the SAME row, column 2
                name = row.iloc[2]  # Column C (Description) - same row as "Entered By :"
                return str(name).strip() if pd.notna(name) else "Unknown"
        
        return "Unknown"
    
    def check_rectify_in_narration(self, block_header_idx, transactions_df):
        """Check if the transaction block contains 'Rectify entry' in narration."""
        # Look through all rows in the transaction block for "Rectify entry"
        for row_idx in range(block_header_idx, len(transactions_df)):
            row = transactions_df.iloc[row_idx]
            description = row.iloc[2]  # Column C (Description)
            
            if pd.notna(description) and isinstance(description, str):
                if "Rectify entry" in description:
                    return True
        
        return False
