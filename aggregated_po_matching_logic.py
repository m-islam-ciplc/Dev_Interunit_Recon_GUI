import pandas as pd
import re

# PO Number extraction pattern - same as existing PO matching logic
PO_PATTERN = r'(?:^|\s)([A-Z0-9/]+/PO/[A-Z0-9/]+)(?:\s|$|[,\.])'

class AggregatedPOMatchingLogic:
    """Handles the logic for finding aggregated PO matches between two files."""
    
    def __init__(self):
        pass
    
    def find_potential_matches(self, transactions1, transactions2, po_numbers1, po_numbers2, existing_matches=None, match_counter=0):
        """Find potential aggregated PO matches between the two files."""
        # Filter rows with PO numbers
        po_transactions1 = transactions1[po_numbers1.notna()].copy()
        po_transactions2 = transactions2[po_numbers2.notna()].copy()
        
        print(f"\nFile 1: {len(po_transactions1)} transactions with PO numbers")
        print(f"File 2: {len(po_transactions2)} transactions with PO numbers")
        
        # Find matches - Aggregated PO Logic
        matches = []
        
        # Use shared state if provided, otherwise create new
        if existing_matches is None:
            existing_matches = {}
        if match_counter is None:
            match_counter = 0
        
        print(f"\n=== AGGREGATED PO MATCHING LOGIC ===")
        print(f"1. Find lender narrations with multiple PO numbers")
        print(f"2. Find all borrower transactions containing these POs")
        print(f"3. Validate: Lender Debit == Sum(All Borrower Credits)")
        print(f"4. Ensure ALL lender POs are present in borrower transactions")
        print(f"5. No tolerance - exact amount match required")
        
        # Process each transaction in File 1 to find multi-PO narrations
        # We need to scan ALL narrations to find those with multiple POs, not just the indexed ones
        processed_narrations = set()  # Track which narrations we've already processed
        
        print(f"\nScanning {len(transactions1)} transactions for multi-PO narrations...")
        
        for idx1 in range(len(transactions1)):
            # Skip if we've already processed this narration
            if idx1 in processed_narrations:
                continue
                
            # Find the transaction block header row for this index
            block_header1 = self.find_transaction_block_header(idx1, transactions1)
            header_row1 = transactions1.iloc[block_header1]
            
            # Extract narration to check for multiple POs
            narration1 = str(header_row1.iloc[2])
            all_pos_in_narration = re.findall(PO_PATTERN, narration1.upper())
            
            # Debug: Show some narrations and PO counts
            if idx1 < 50:  # Only show first 50 for debugging
                print(f"  Row {idx1}: {len(all_pos_in_narration)} POs found in narration")
                if len(all_pos_in_narration) > 0:
                    print(f"    Sample POs: {all_pos_in_narration[:3]}")
            
            # Only process if there are multiple POs
            if len(all_pos_in_narration) < 2:
                continue
                
            print(f"\n--- Processing File 1 Row {block_header1} with {len(all_pos_in_narration)} POs ---")
            print(f"  POs found: {all_pos_in_narration}")
            
            # Extract amounts and determine transaction type for File 1
            file1_debit = header_row1.iloc[7] if pd.notna(header_row1.iloc[7]) else 0
            file1_credit = header_row1.iloc[8] if pd.notna(header_row1.iloc[8]) else 0
            
            file1_is_lender = file1_debit > 0
            file1_is_borrower = file1_credit > 0
            file1_amount = file1_debit if file1_is_lender else file1_credit
            
            print(f"  File 1: Amount={file1_amount}, Type={'Lender' if file1_is_lender else 'Borrower'}")
            
            # Only process lender transactions (debit > 0)
            if not file1_is_lender:
                print(f"  ❌ SKIP: Not a lender transaction (credit > 0)")
                continue
            
            # Mark this narration as processed
            processed_narrations.add(block_header1)
            
            print(f"  ✅ MULTI-PO NARRATION FOUND: {len(all_pos_in_narration)} POs")
            
            # Now look for matching borrower transactions
            matching_borrower_data = []
            total_borrower_amount = 0
            
            for idx2, po2 in enumerate(po_numbers2):
                if not po2:
                    continue
                
                # Check if this PO is in our lender's PO list
                if po2 not in all_pos_in_narration:
                    continue
                
                print(f"    Found matching PO {po2} in File 2 Row {idx2}")
                
                # Find the transaction block header row for this PO in File 2
                block_header2 = self.find_transaction_block_header(idx2, transactions2)
                header_row2 = transactions2.iloc[block_header2]
                
                # Extract amounts and determine transaction type for File 2
                file2_debit = header_row2.iloc[7] if pd.notna(header_row2.iloc[7]) else 0
                file2_credit = header_row2.iloc[8] if pd.notna(header_row2.iloc[8]) else 0
                
                file2_is_lender = file2_debit > 0
                file2_is_borrower = file2_credit > 0
                file2_amount = file2_credit if file2_is_borrower else file2_debit
                
                print(f"      File 2: Amount={file2_amount}, Type={'Lender' if file2_is_lender else 'Borrower'}")
                
                # Only process borrower transactions (credit > 0)
                if not file2_is_borrower:
                    print(f"      ❌ SKIP: Not a borrower transaction (debit > 0)")
                    continue
                
                # Add to matching data
                matching_borrower_data.append({
                    'row': idx2,
                    'po': po2,
                    'amount': file2_amount,
                    'block_header': block_header2
                })
                total_borrower_amount += file2_amount
                
                print(f"      ✅ ADDED: PO {po2}, Amount {file2_amount}, Total so far: {total_borrower_amount}")
            
            # Check if we have matches for all POs
            matched_pos = [data['po'] for data in matching_borrower_data]
            missing_pos = [po for po in all_pos_in_narration if po not in matched_pos]
            
            if missing_pos:
                print(f"  ❌ REJECTED: Missing POs in borrower transactions: {missing_pos}")
                continue
            
            print(f"  ✅ ALL POs FOUND: {len(matched_pos)}/{len(all_pos_in_narration)}")
            
            # Check if amounts match exactly
            if abs(file1_amount - total_borrower_amount) > 0.01:  # Allow small rounding differences
                print(f"  ❌ REJECTED: Amounts don't match (Lender: {file1_amount}, Borrower Sum: {total_borrower_amount})")
                continue
            
            print(f"  ✅ AMOUNTS MATCH: Lender {file1_amount} == Borrower Sum {total_borrower_amount}")
            
            # Check if we already have a match for this combination
            match_key = (tuple(sorted(all_pos_in_narration)), file1_amount)
            
            if match_key in existing_matches:
                # Use existing Match ID for consistency
                match_id = existing_matches[match_key]
                print(f"  🔄 REUSING existing Match ID: {match_id}")
            else:
                # Create new Match ID
                match_counter += 1
                match_id = f"M{match_counter:03d}"
                existing_matches[match_key] = match_id
                print(f"  🆕 CREATING new Match ID: {match_id}")
            
            print(f"  🎉 ALL CRITERIA MET - AGGREGATED PO MATCH FOUND!")
            
            # Create the match
            matches.append({
                'match_id': match_id,
                'Match_Type': 'Aggregated_PO',
                'File1_Index': block_header1,
                'File2_Indices': [data['block_header'] for data in matching_borrower_data],
                'All_POs': all_pos_in_narration,
                'PO_Count': len(all_pos_in_narration),
                'File1_Date': header_row1.iloc[0],
                'File1_Description': header_row1.iloc[2],
                'File1_Debit': header_row1.iloc[7],
                'File1_Credit': header_row1.iloc[8],
                'File1_Amount': file1_amount,
                'File1_Type': 'Lender',
                'Borrower_Details': matching_borrower_data,
                'Total_Borrower_Amount': total_borrower_amount
            })
        
        print(f"\n=== AGGREGATED PO MATCHING RESULTS ===")
        print(f"Found {len(matches)} valid aggregated PO matches across {len(existing_matches)} unique Match ID combinations!")
        
        # Show some examples
        if matches:
            print("\n=== SAMPLE AGGREGATED PO MATCHES ===")
            for i, match in enumerate(matches[:5]):
                print(f"\nAggregated PO Match {i+1}:")
                print(f"Match ID: {match['match_id']}")
                print(f"PO Count: {match['PO_Count']}")
                print(f"All POs: {match['All_POs']}")
                print(f"Lender Amount: {match['File1_Amount']}")
                print(f"Total Borrower Amount: {match['Total_Borrower_Amount']}")
                print(f"File 1: {match['File1_Date']} - {str(match['File1_Description'])[:50]}...")
                print(f"  Type: {match['File1_Type']}, Debit: {match['File1_Debit']}, Credit: {match['File1_Credit']}")
                print(f"Borrower Transactions: {len(match['Borrower_Details'])}")
                for j, borrower in enumerate(match['Borrower_Details'][:3]):  # Show first 3
                    print(f"  {j+1}. PO {borrower['po']}: Amount {borrower['amount']}")
                if len(match['Borrower_Details']) > 3:
                    print(f"  ... and {len(match['Borrower_Details']) - 3} more")
        
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
