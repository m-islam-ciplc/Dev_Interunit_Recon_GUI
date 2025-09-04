import pandas as pd

# LC Number extraction pattern
LC_PATTERN = r'\b(?:L/C|LC)[-\s]?\d+[/\s]?\d*\b'

# Configuration


class LCMatchingLogic:
    """Handles the logic for finding LC number matches between two files."""
    
    def __init__(self, block_identifier):
        """
        Initialize with a shared TransactionBlockIdentifier instance.
        
        Args:
            block_identifier: Shared instance of TransactionBlockIdentifier for consistent transaction block logic
        """
        self.block_identifier = block_identifier
    
    def find_potential_matches(self, transactions1, transactions2, lc_numbers1, lc_numbers2, existing_matches=None, match_id_manager=None):
        """Find potential LC number matches between the two files."""
        # Filter rows with LC numbers
        lc_transactions1 = transactions1[lc_numbers1.notna()].copy()
        lc_transactions2 = transactions2[lc_numbers2.notna()].copy()
        
        print(f"\nFile 1: {len(lc_transactions1)} transactions with LC numbers")
        print(f"File 2: {len(lc_transactions2)} transactions with LC numbers")
        
        # Find matches - NEW LOGIC: Amount → Entered By → LC Number
        matches = []
        
        # Use shared state if provided, otherwise create new
        if existing_matches is None:
            existing_matches = {}
        if match_id_manager is None:
            try:
                from ..match_id_manager import get_match_id_manager
            except ImportError:
                from match_id_manager import get_match_id_manager
            match_id_manager = get_match_id_manager()
        
        print(f"\n=== NEW MATCHING LOGIC ===")
        print(f"1. Check if lender debit and borrower credit amounts are EXACTLY the same")
        print(f"2. Check if LC numbers match between them")
        print(f"3. Only if both match: Assign same Match ID")
        print(f"4. IMPORTANT: All transactions with same LC+Amount get SAME Match ID")
        
        # Use shared state for tracking which combinations have already been matched
        # Key: (LC_Number, Amount), Value: match_id
        
        # Process each transaction in File 1 to find matches in File 2
        for idx1, lc1 in enumerate(lc_numbers1):
            if not lc1:
                continue
                
            print(f"\n--- Processing File 1 Row {idx1} with LC: {lc1} ---")
            
            # Find the transaction block header row for this LC in File 1
            block_header1 = self.block_identifier.find_transaction_block_header(idx1, transactions1)
            header_row1 = transactions1.iloc[block_header1]
            
            # Extract amounts and determine transaction type for File 1
            # Based on investigation: amounts are in columns 8 and 9 (iloc[7] and iloc[8])
            file1_debit = header_row1.iloc[7] if pd.notna(header_row1.iloc[7]) else 0
            file1_credit = header_row1.iloc[8] if pd.notna(header_row1.iloc[8]) else 0
            
            file1_is_lender = file1_debit > 0
            file1_is_borrower = file1_credit > 0
            file1_amount = file1_debit if file1_is_lender else file1_credit
            
            print(f"  File 1: Amount={file1_amount}, Type={'Lender' if file1_is_lender else 'Borrower'}")
            
            # Now look for matches in File 2
            for idx2, lc2 in enumerate(lc_numbers2):
                if not lc2:
                    continue
                    
                print(f"    Checking File 2 Row {idx2} with LC: {lc2}")
                
                # Find the transaction block header row for this LC in File 2
                block_header2 = self.block_identifier.find_transaction_block_header(idx2, transactions2)
                header_row2 = transactions2.iloc[block_header2]
                
                # Extract amounts and determine transaction type for File 2
                # Based on investigation: amounts are in columns 8 and 9 (iloc[7] and iloc[8])
                file2_debit = header_row2.iloc[7] if pd.notna(header_row2.iloc[7]) else 0
                file2_credit = header_row2.iloc[8] if pd.notna(header_row2.iloc[8]) else 0
                
                file2_is_lender = file2_debit > 0
                file2_is_borrower = file2_credit > 0
                file2_amount = file2_debit if file2_is_lender else file2_credit
                
                print(f"      File 2: Amount={file2_amount}, Type={'Lender' if file2_is_lender else 'Borrower'}")
                
                # STEP 1: Check if amounts are EXACTLY the same
                if file1_amount != file2_amount:
                    print(f"      REJECTED: Amounts don't match ({file1_amount} vs {file2_amount})")
                    continue
                
                print(f"      STEP 1 PASSED: Amounts match exactly")
                
                # STEP 2: Check if transaction types are opposite (one lender, one borrower)
                if not ((file1_is_lender and file2_is_borrower) or (file1_is_borrower and file2_is_lender)):
                    print(f"      REJECTED: Transaction types don't match (both same type)")
                    continue
                
                print(f"      STEP 2 PASSED: Transaction types are opposite")
                
                # STEP 3: Check if LC numbers match
                if lc1 != lc2:
                    print(f"      REJECTED: LC numbers don't match ('{lc1}' vs '{lc2}')")
                    continue
                
                print(f"      STEP 3 PASSED: LC numbers match")
                
                # STEP 4: Check if we already have a match for this combination
                # Create new sequential Match ID (simplified approach)
                context = f"LC_{lc1}_File1_Row_{idx1}_File2_Row_{idx2}"
                # Match ID will be assigned later in post-processing
                match_id = None
                print(f"      CREATING new Match ID: {match_id}")
                
                print(f"      ALL CRITERIA MET - MATCH FOUND!")
                
                # Create the match
                matches.append({
                    'match_id': match_id,
                    'Match_Type': 'LC',  # Add explicit match type
                    'File1_Index': block_header1,
                    'File2_Index': block_header2,
                    'LC_Number': lc1,
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
                    'File2_Type': 'Lender' if file2_is_lender else 'Borrower'
                })
        
        print(f"\n=== MATCHING RESULTS ===")
        print(f"Found {len(matches)} valid matches across {len(existing_matches)} unique Match ID combinations!")
        
        # Show some examples
        if matches:
            print("\n=== SAMPLE MATCHES ===")
            for i, match in enumerate(matches[:5]):
                print(f"\nMatch {i+1}:")
                print(f"Match ID: {match['match_id']}")
                print(f"LC Number: {match['LC_Number']}")
                print(f"Amount: {match['File1_Amount']}")
                print(f"File 1: {match['File1_Date']} - {str(match['File1_Description'])[:50]}...")
                print(f"  Type: {match['File1_Type']}, Debit: {match['File1_Debit']}, Credit: {match['File1_Credit']}")
                print(f"File 2: {match['File2_Date']} - {str(match['File2_Description'])[:50]}...")
                print(f"  Type: {match['File2_Type']}, Debit: {match['File2_Debit']}, Credit: {match['File2_Credit']}")
        
        return matches
    
    # Transaction block identification methods are now provided by the shared TransactionBlockIdentifier instance
    # This ensures consistent behavior across all matching modules
    

    

