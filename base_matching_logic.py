"""
Base Matching Logic Module

Provides common infrastructure for all transaction matching algorithms.
Reduces code duplication and ensures consistent behavior across all matchers.
"""

import pandas as pd
from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Any, Tuple


class BaseMatchingLogic(ABC):
    """
    Abstract base class for all transaction matching logic modules.
    
    Provides common functionality:
    - Amount validation
    - Transaction type determination
    - Match creation workflow
    - Shared state management
    """
    
    def __init__(self, block_identifier):
        """
        Initialize with a shared TransactionBlockIdentifier instance.
        
        Args:
            block_identifier: Shared instance of TransactionBlockIdentifier
        """
        self.block_identifier = block_identifier
    
    def find_potential_matches(
        self, 
        transactions1: pd.DataFrame, 
        transactions2: pd.DataFrame, 
        data1: pd.Series,
        data2: pd.Series,
        existing_matches: Dict = None, 
        match_id_manager = None
    ) -> List[Dict]:
        """
        Template method for finding matches. Subclasses implement specific logic.
        
        Args:
            transactions1: DataFrame of transactions from first file
            transactions2: DataFrame of transactions from second file  
            data1: Series of extracted data from first file (LC numbers, PO numbers, etc.)
            data2: Series of extracted data from second file
            existing_matches: Dictionary of existing matches (shared state)
            match_id_manager: Manager for generating sequential match IDs
            
        Returns:
            List of match dictionaries
        """
        # Initialize shared state
        if existing_matches is None:
            existing_matches = {}
        if match_id_manager is None:
            from match_id_manager import get_match_id_manager
            match_id_manager = get_match_id_manager()
        
        print(f"\n=== {self.get_match_type().upper()} MATCHING LOGIC ===")
        self.print_matching_criteria()
        
        matches = []
        
        # Filter to rows with relevant data
        filtered_data1 = data1.dropna()
        filtered_data2 = data2.dropna()
        
        print(f"File 1: {len(filtered_data1)} transactions with {self.get_match_type()}")
        print(f"File 2: {len(filtered_data2)} transactions with {self.get_match_type()}")
        
        # Process each item in file 1
        for idx1, item1 in filtered_data1.items():
            if not self.is_valid_item(item1):
                continue
                
            print(f"\n--- Processing File 1 Row {idx1} with {self.get_match_type()}: {item1} ---")
            
            # Get transaction data for file 1
            transaction_data1 = self.get_transaction_data(idx1, transactions1)
            if not transaction_data1:
                continue
                
            print(f"  File 1: Amount={transaction_data1['amount']}, Type={transaction_data1['type']}")
            
            # Look for matches in file 2
            for idx2, item2 in filtered_data2.items():
                if not self.is_valid_item(item2):
                    continue
                    
                print(f"    Checking File 2 Row {idx2} with {self.get_match_type()}: {item2}")
                
                # Get transaction data for file 2
                transaction_data2 = self.get_transaction_data(idx2, transactions2)
                if not transaction_data2:
                    continue
                    
                print(f"      File 2: Amount={transaction_data2['amount']}, Type={transaction_data2['type']}")
                
                # Validate match criteria
                if self.validate_match(transaction_data1, transaction_data2, item1, item2):
                    # Create match
                    match = self.create_match(
                        transaction_data1, transaction_data2, 
                        item1, item2, idx1, idx2
                    )
                    matches.append(match)
                    print(f"      ALL CRITERIA MET - {self.get_match_type().upper()} MATCH FOUND!")
        
        print(f"\n=== {self.get_match_type().upper()} MATCHING RESULTS ===")
        print(f"Found {len(matches)} valid {self.get_match_type()} matches!")
        
        self.print_sample_matches(matches)
        
        return matches
    
    def get_transaction_data(self, row_idx: int, transactions_df: pd.DataFrame) -> Optional[Dict]:
        """Extract transaction data (amounts, type, etc.) for a given row."""
        try:
            # Find the transaction block header row
            block_header = self.block_identifier.find_transaction_block_header(row_idx, transactions_df)
            header_row = transactions_df.iloc[block_header]
            
            # Extract amounts (columns 7 and 8 are debit/credit)
            debit = header_row.iloc[7] if pd.notna(header_row.iloc[7]) else 0
            credit = header_row.iloc[8] if pd.notna(header_row.iloc[8]) else 0
            
            is_lender = debit > 0
            is_borrower = credit > 0
            amount = debit if is_lender else credit
            
            return {
                'block_header': block_header,
                'header_row': header_row,
                'debit': debit,
                'credit': credit,
                'amount': amount,
                'is_lender': is_lender,
                'is_borrower': is_borrower,
                'type': 'Lender' if is_lender else 'Borrower',
                'date': header_row.iloc[0],
                'description': header_row.iloc[2]
            }
        except Exception as e:
            print(f"Error extracting transaction data for row {row_idx}: {e}")
            return None
    
    def validate_match(
        self, 
        transaction_data1: Dict, 
        transaction_data2: Dict, 
        item1: Any, 
        item2: Any
    ) -> bool:
        """
        Validate if two transactions can be matched.
        Common validation: amounts match and transaction types are opposite.
        """
        # Step 1: Check if amounts are exactly the same
        if transaction_data1['amount'] != transaction_data2['amount']:
            print(f"       REJECTED: Amounts don't match ({transaction_data1['amount']} vs {transaction_data2['amount']})")
            return False
        
        print(f"       STEP 1 PASSED: Amounts match exactly")
        
        # Step 2: Check if transaction types are opposite
        if not ((transaction_data1['is_lender'] and transaction_data2['is_borrower']) or 
                (transaction_data1['is_borrower'] and transaction_data2['is_lender'])):
            print(f"       REJECTED: Transaction types don't match (both same type)")
            return False
        
        print(f"       STEP 2 PASSED: Transaction types are opposite")
        
        # Step 3: Specific validation (implemented by subclasses)
        return self.validate_specific_criteria(transaction_data1, transaction_data2, item1, item2)
    
    def create_match(
        self,
        transaction_data1: Dict,
        transaction_data2: Dict, 
        item1: Any,
        item2: Any,
        idx1: int,
        idx2: int
    ) -> Dict:
        """Create a match dictionary with common fields."""
        # Base match structure
        match = {
            'match_id': None,  # Will be assigned later in post-processing
            'Match_Type': self.get_match_type(),
            'File1_Index': transaction_data1['block_header'],
            'File2_Index': transaction_data2['block_header'],
            'File1_Date': transaction_data1['date'],
            'File1_Description': transaction_data1['description'],
            'File1_Debit': transaction_data1['debit'],
            'File1_Credit': transaction_data1['credit'],
            'File2_Date': transaction_data2['date'],
            'File2_Description': transaction_data2['description'],
            'File2_Debit': transaction_data2['debit'],
            'File2_Credit': transaction_data2['credit'],
            'File1_Amount': transaction_data1['amount'],
            'File2_Amount': transaction_data2['amount'],
            'File1_Type': transaction_data1['type'],
            'File2_Type': transaction_data2['type']
        }
        
        # Add specific fields (implemented by subclasses)
        specific_fields = self.get_specific_match_fields(item1, item2)
        match.update(specific_fields)
        
        return match
    
    def print_sample_matches(self, matches: List[Dict], max_samples: int = 3):
        """Print sample matches for debugging."""
        if not matches:
            return
            
        print(f"\n=== SAMPLE {self.get_match_type().upper()} MATCHES ===")
        for i, match in enumerate(matches[:max_samples]):
            print(f"\n{self.get_match_type()} Match {i+1}:")
            print(f"Match ID: {match['match_id']}")
            print(f"Amount: {match['File1_Amount']}")
            print(f"File 1: {match['File1_Date']} - {str(match['File1_Description'])[:50]}...")
            print(f"  Type: {match['File1_Type']}, Debit: {match['File1_Debit']}, Credit: {match['File1_Credit']}")
            print(f"File 2: {match['File2_Date']} - {str(match['File2_Description'])[:50]}...")
            print(f"  Type: {match['File2_Type']}, Debit: {match['File2_Debit']}, Credit: {match['File2_Credit']}")
    
    # Abstract methods that subclasses must implement
    
    @abstractmethod
    def get_match_type(self) -> str:
        """Return the type of match (e.g., 'LC', 'PO', 'USD')."""
        pass
    
    @abstractmethod
    def print_matching_criteria(self):
        """Print the specific matching criteria for this match type."""
        pass
    
    @abstractmethod
    def is_valid_item(self, item: Any) -> bool:
        """Check if an item is valid for matching."""
        pass
    
    @abstractmethod
    def validate_specific_criteria(
        self, 
        transaction_data1: Dict, 
        transaction_data2: Dict, 
        item1: Any, 
        item2: Any
    ) -> bool:
        """Validate specific criteria for this match type."""
        pass
    
    @abstractmethod
    def get_specific_match_fields(self, item1: Any, item2: Any) -> Dict:
        """Return specific fields for this match type."""
        pass
