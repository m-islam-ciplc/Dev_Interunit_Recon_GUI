import pandas as pd

class NarrationMatchingLogic:
    """Handles the logic for finding identical narration matches between two files."""
    
    def __init__(self, block_identifier):
        """
        Initialize with a shared TransactionBlockIdentifier instance.
        
        Args:
            block_identifier: Shared instance of TransactionBlockIdentifier for consistent transaction block logic
        """
        self.block_identifier = block_identifier
    
    def find_potential_matches(self, transactions1, transactions2, existing_matches=None, match_counter=0):
        """Find potential identical narration matches between the two files."""
        
        print(f"\nFile 1: {len(transactions1)} transactions")
        print(f"File 2: {len(transactions2)} transactions")
        
        # Find matches - Identical Narration Logic
        matches = []
        
        # Use shared state if provided, otherwise create new
        if existing_matches is None:
            existing_matches = {}
        if match_counter is None:
            match_counter = 0
        
        print(f"\n=== NARRATION MATCHING LOGIC (STEP 1 - HIGHEST PRIORITY) ===")
        print(f"1. Find narrations with EXACTLY identical text between files")
        print(f"2. Validate: Lender Debit == Borrower Credit")
        print(f"3. Most reliable match type - runs first")
        
        # Process each transaction in File 1 to find identical narrations
        processed_narrations = set()  # Track which narrations we've already processed
        
        for idx1 in range(len(transactions1)):
            # Skip if we've already processed this narration
            if idx1 in processed_narrations:
                continue
                
            # Find the transaction block header row for this index
            block_header1 = self.block_identifier.find_transaction_block_header(idx1, transactions1)
            header_row1 = transactions1.iloc[block_header1]
            
            # Find the description row within this transaction block (narration is in description rows)
            description_row1 = self.block_identifier.find_description_row_in_block(idx1, transactions1)
            if description_row1 is None:
                continue
                
            # Extract narration from the DESCRIPTION row (not header row)
            narration1 = str(transactions1.iloc[description_row1, 2]).strip()
            
            # Skip empty or very short narrations
            if len(narration1) < 10 or narration1.lower() in ['nan', 'none', '']:
                continue
                
            # Extract amounts from the HEADER row (amounts are in header rows)
            file1_debit = header_row1.iloc[7] if pd.notna(header_row1.iloc[7]) else 0
            file1_credit = header_row1.iloc[8] if pd.notna(header_row1.iloc[8]) else 0
            
            file1_is_lender = file1_debit > 0
            file1_is_borrower = file1_credit > 0
            file1_amount = file1_debit if file1_is_lender else file1_credit
            
            # Only process if there's a valid amount
            if file1_amount <= 0:
                continue
            
            # Mark this narration as processed
            processed_narrations.add(block_header1)
            processed_narrations.add(description_row1)
            
            # Also mark all rows in the same transaction block as processed
            # to avoid processing the same narration multiple times
            for i in range(len(transactions1)):
                if i != block_header1 and i != description_row1:  # Don't mark header or description
                    other_row = transactions1.iloc[i]
                    other_narration = str(other_row.iloc[2]).strip()
                    if other_narration == narration1:
                        processed_narrations.add(i)
            
            print(f"\n--- Processing File 1 Row {block_header1} (Header) / {description_row1} (Description) ---")
            print(f"  Narration: {narration1[:80]}...")
            print(f"  Amount: {file1_amount}, Type: {'Lender' if file1_is_lender else 'Borrower'}")
            
            # Now search for identical narrations in File 2
            matching_file2_transactions = []
            processed_file2_narrations = set()  # Track processed File 2 narrations to avoid duplicates
            
            for idx2 in range(len(transactions2)):
                # Skip if we've already processed this narration in File 2
                if idx2 in processed_file2_narrations:
                    continue
                    
                # Find the transaction block header row for this index in File 2
                block_header2 = self.block_identifier.find_transaction_block_header(idx2, transactions2)
                header_row2 = transactions2.iloc[block_header2]
                
                # Find the description row within this transaction block in File 2
                description_row2 = self.block_identifier.find_description_row_in_block(idx2, transactions2)
                if description_row2 is None:
                    continue
                
                # Extract narration from the DESCRIPTION row in File 2
                narration2 = str(transactions2.iloc[description_row2, 2]).strip()
                
                # Check for EXACT narration match
                if narration1 == narration2:
                    # Found identical narration, check amounts and transaction type from header row
                    file2_debit = header_row2.iloc[7] if pd.notna(header_row2.iloc[7]) else 0
                    file2_credit = header_row2.iloc[8] if pd.notna(header_row2.iloc[8]) else 0
                    
                    file2_is_lender = file2_debit > 0
                    file2_is_borrower = file2_credit > 0
                    file2_amount = file2_debit if file2_is_lender else file2_credit
                    
                    # Only process if there's a valid amount
                    if file2_amount <= 0:
                        continue
                    
                    # Check if this creates a valid lender-borrower pair
                    if (file1_is_lender and file2_is_borrower) or (file1_is_borrower and file2_is_lender):
                        # Ensure amounts match (lender debit = borrower credit)
                        lender_amount = file1_debit if file1_is_lender else file2_debit
                        borrower_amount = file1_credit if file1_is_borrower else file2_credit
                        
                        if abs(lender_amount - borrower_amount) < 0.01:  # Allow small rounding differences
                            matching_file2_transactions.append({
                                'row': idx2,
                                'header_row': block_header2,
                                'description_row': description_row2,
                                'amount': file2_amount,
                                'type': 'Lender' if file2_is_lender else 'Borrower',
                                'date': header_row2.iloc[0]
                            })
                            print(f"    ✅ Found identical narration in File 2 Row {block_header2} (Header) / {description_row2} (Description)")
                            print(f"      Amount: {file2_amount}, Type: {'Lender' if file2_is_lender else 'Borrower'}")
                            
                            # Mark this narration and all similar ones in File 2 as processed
                            for i in range(len(transactions2)):
                                if i != block_header2 and i != description_row2:  # Don't mark header or description
                                    other_row = transactions2.iloc[i]
                                    other_narration = str(other_row.iloc[2]).strip()
                                    if other_narration == narration2:
                                        processed_file2_narrations.add(i)
                            processed_file2_narrations.add(block_header2)
                            processed_file2_narrations.add(description_row2)
            
            # If we found matching transactions, create the match
            if matching_file2_transactions:
                print(f"  🎯 Found {len(matching_file2_transactions)} matching transactions")
                
                # Create only ONE match per unique narration combination
                # Take the first matching transaction (they should all be equivalent)
                match_data = matching_file2_transactions[0]
                
                # Check if we already have a match for this combination
                match_key = (narration1, file1_amount, match_data['amount'])
                
                if match_key in existing_matches:
                    # Use existing Match ID for consistency
                    match_id = existing_matches[match_key]
                    print(f"    🔄 REUSING existing Match ID: {match_id}")
                else:
                    # Create new Match ID
                    match_counter += 1
                    match_id = f"M{match_counter:03d}"
                    existing_matches[match_key] = match_id
                    print(f"    🆕 CREATING new Match ID: {match_id}")
                
                # Determine which file is lender and which is borrower
                if file1_is_lender:
                    lender_file = 1
                    lender_index = block_header1
                    borrower_file = 2
                    borrower_index = match_data['header_row']
                    lender_amount = file1_amount
                    borrower_amount = match_data['amount']
                else:
                    lender_file = 2
                    lender_index = match_data['header_row']
                    borrower_file = 1
                    borrower_index = block_header1
                    lender_amount = match_data['amount']
                    borrower_amount = file1_amount
                
                print(f"    🎉 NARRATION MATCH FOUND!")
                
                # Create the match
                matches.append({
                    'match_id': match_id,
                    'Match_Type': 'Narration',
                    'File1_Index': block_header1,
                    'File2_Index': match_data['header_row'],
                    'Narration': narration1,
                    'File1_Amount': file1_amount,
                    'File2_Amount': match_data['amount'],
                    'File1_Date': header_row1.iloc[0],
                    'File2_Date': match_data['date'],
                    'Lender_File': lender_file,
                    'Lender_Index': lender_index,
                    'Borrower_File': borrower_file,
                    'Borrower_Index': borrower_index,
                    'Lender_Amount': lender_amount,
                    'Borrower_Amount': borrower_amount,
                    # Add the missing fields required by output creation
                    'File1_Debit': file1_debit,
                    'File1_Credit': file1_credit,
                    'File2_Debit': match_data['amount'] if match_data['type'] == 'Lender' else 0,
                    'File2_Credit': match_data['amount'] if match_data['type'] == 'Borrower' else 0
                })
            else:
                print(f"  ❌ No identical narrations found in File 2")
        
        print(f"\n=== NARRATION MATCHING RESULTS ===")
        print(f"Found {len(matches)} valid narration matches across {len(existing_matches)} unique Match ID combinations!")
        
        return matches
    
    # Transaction block identification methods are now provided by the shared TransactionBlockIdentifier instance
    # This ensures consistent behavior across all matching modules
