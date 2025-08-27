import pandas as pd

# Configuration
AMOUNT_TOLERANCE = 0.01  # Amount matching tolerance for rounding differences

class LCMatchingLogic:
    """Handles the logic for finding LC number matches between two files."""
    
    def __init__(self):
        self.amount_tolerance = AMOUNT_TOLERANCE
    
    def find_potential_matches(self, transactions1, transactions2, lc_numbers1, lc_numbers2):
        """Find potential LC number matches between the two files."""
        # Filter rows with LC numbers
        lc_transactions1 = transactions1[lc_numbers1.notna()].copy()
        lc_transactions2 = transactions2[lc_numbers2.notna()].copy()
        
        print(f"\nFile 1: {len(lc_transactions1)} transactions with LC numbers")
        print(f"File 2: {len(lc_transactions2)} transactions with LC numbers")
        
        # Find matches
        matches = []
        match_counter = 0
        
        for idx1, lc1 in enumerate(lc_numbers1):
            if not lc1:
                continue
            for idx2, lc2 in enumerate(lc_numbers2):
                if not lc2:
                    continue
                if lc1 == lc2:
                    # Find the transaction block header row for each LC
                    # This is the row with date and particulars (Dr/Cr)
                    block_header1 = self.find_transaction_block_header(idx1, transactions1)
                    block_header2 = self.find_transaction_block_header(idx2, transactions2)
                    
                    # Get the transaction block header rows
                    header_row1 = transactions1.iloc[block_header1]
                    header_row2 = transactions2.iloc[block_header2]
                    
                    # Extract amounts from both files
                    file1_debit = header_row1.iloc[7] if pd.notna(header_row1.iloc[7]) else 0
                    file1_credit = header_row1.iloc[8] if pd.notna(header_row1.iloc[8]) else 0
                    file2_debit = header_row2.iloc[7] if pd.notna(header_row2.iloc[7]) else 0
                    file2_credit = header_row2.iloc[8] if pd.notna(header_row2.iloc[8]) else 0
                    
                    # Determine transaction types and amounts
                    # Lender: Has Debit amount (Dr), Borrower: Has Credit amount (Cr)
                    file1_is_lender = file1_debit > 0
                    file1_is_borrower = file1_credit > 0
                    file2_is_lender = file2_debit > 0
                    file2_is_borrower = file2_credit > 0
                    
                    # Get the actual amounts
                    file1_amount = file1_debit if file1_is_lender else file1_credit
                    file2_amount = file2_debit if file2_is_lender else file2_credit
                    
                    # CRITICAL: Only create a match if:
                    # 1. One file is lender (Dr) and other is borrower (Cr)
                    # 2. The amounts are the same (within tolerance)
                    if ((file1_is_lender and file2_is_borrower) or (file1_is_borrower and file2_is_lender)):
                        # Check if amounts match (within configured tolerance for rounding)
                        if abs(file1_amount - file2_amount) < self.amount_tolerance:
                            match_counter += 1
                            
                            # Debug: Print the actual row data to see what we're working with   
                            print(f"\nDEBUG: LC {lc1} VALID match found (amounts match):")
                            print(f"  File1 Description Row {idx1} → Block Header Row {block_header1}: Date={header_row1.iloc[0]}, Particulars={header_row1.iloc[1]}")
                            print(f"  File2 Description Row {idx2} → Block Header Row {block_header2}: Date={header_row2.iloc[0]}, Particulars={header_row2.iloc[1]}")
                            print(f"  Amounts: File1={file1_amount} ({'Lender' if file1_is_lender else 'Borrower'}), File2={file2_amount} ({'Lender' if file2_is_lender else 'Borrower'})")
                            
                            matches.append({
                                'match_id': f"M{match_counter:03d}",  # Unique match ID
                                'File1_Index': block_header1,  # This is the transaction block header row
                                'File2_Index': block_header2,  # This is the transaction block header row
                                'LC_Number': lc1,
                                'File1_Date': header_row1.iloc[0],  # First column (Date)
                                'File1_Description': header_row1.iloc[2],  # Third column (Description)
                                'File1_Debit': header_row1.iloc[7],  # Eighth column (Debit)
                                'File1_Credit': header_row1.iloc[8],  # Ninth column (Credit)
                                'File2_Date': header_row2.iloc[0],  # First column (Date)
                                'File2_Description': header_row2.iloc[2],  # Third column (Description)
                                'File2_Debit': header_row2.iloc[7],  # Eighth column (Debit)
                                'File2_Credit': header_row2.iloc[8],  # Ninth column (Credit)
                                'File1_Amount': file1_amount,
                                'File2_Amount': file2_amount,
                                'File1_Type': 'Lender' if file1_is_lender else 'Borrower',
                                'File2_Type': 'Lender' if file2_is_lender else 'Borrower'
                            })
                        else:
                            print(f"\nDEBUG: LC {lc1} REJECTED - amounts don't match:")
                            print(f"  File1: {file1_amount} ({'Lender' if file1_is_lender else 'Borrower'})")
                            print(f"  File2: {file2_amount} ({'Lender' if file2_is_lender else 'Borrower'})")
                            print(f"  Difference: {abs(file1_amount - file2_amount)}")
                    else:
                        print(f"\nDEBUG: LC {lc1} REJECTED - transaction types don't match:")
                        print(f"  File1: {'Lender' if file1_is_lender else 'Borrower' if file1_is_borrower else 'Neither'}")
                        print(f"  File2: {'Lender' if file2_is_lender else 'Borrower' if file2_is_borrower else 'Neither'}")
        
        print(f"\nFound {len(matches)} potential LC matches!")
        
        # Show some examples
        if matches:
            print("\n=== SAMPLE MATCHES ===")
            for i, match in enumerate(matches[:5]):  # Show first 5 matches
                print(f"\nMatch {i+1}:")
                print(f"LC Number: {match['LC_Number']}")
                print(f"File 1: {match['File1_Date']} - {str(match['File1_Description'])[:50]}...")
                print(f"  Debit: {match['File1_Debit']}, Credit: {match['File1_Credit']}")
                print(f"File 2: {match['File2_Date']} - {str(match['File2_Description'])[:50]}...")
                print(f"  Debit: {match['File2_Debit']}, Credit: {match['File2_Credit']}")
        
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
            has_debit = pd.notna(row.iloc[7]) and row.iloc[7] != 0
            has_credit = pd.notna(row.iloc[8]) and row.iloc[8] != 0
            
            # Transaction block header: has date, particulars, and either debit or credit
            if has_date and (has_debit or has_credit):
                return row_idx
        
        # If no header found, return the description row itself
        return description_row_idx
    
    def set_amount_tolerance(self, tolerance):
        """Set the amount tolerance for matching."""
        self.amount_tolerance = tolerance
        print(f"Amount tolerance set to: {self.amount_tolerance}")
