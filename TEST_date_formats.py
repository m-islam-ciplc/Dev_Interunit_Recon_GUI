import pandas as pd
import os

def test_date_formats():
    """Test script to examine actual date formats in input vs output files."""
    
    print("=== TESTING ACTUAL DATE FORMATS ===\n")
    
    # Test input file
    input_file = "Input Files/Interunit GeoTex.xlsx"
    print(f"1. READING INPUT FILE: {input_file}")
    
    try:
        # Read the input file to see original format
        input_df = pd.read_excel(input_file, header=None)
        
        # Extract transaction data (rows 8+, which are Excel rows 9+)
        transactions = input_df.iloc[8:, :]
        
        # Set first row as headers and remove it from data
        transactions.columns = transactions.iloc[0]
        transactions = transactions.iloc[1:].reset_index(drop=True)
        
        print(f"   Input file shape: {transactions.shape}")
        print(f"   Input file columns: {list(transactions.columns)}")
        
        # Show first 5 date values from input
        if len(transactions.columns) > 0:
            date_col = transactions.iloc[:, 0]  # First column should be date
            print(f"   Input file - First 5 date values:")
            for i, date_val in enumerate(date_col.head()):
                print(f"     Row {i}: {date_val} (Type: {type(date_val).__name__})")
        
    except Exception as e:
        print(f"   ERROR reading input file: {e}")
    
    print("\n" + "="*50 + "\n")
    
    # Test output file
    output_file = "Output/Interunit GeoTex_MATCHED.xlsx"
    print(f"2. READING OUTPUT FILE: {output_file}")
    
    try:
        # Read the output file to see final format
        output_df = pd.read_excel(output_file, header=None)
        
        # Extract transaction data (rows 8+, which are Excel rows 9+)
        output_transactions = output_df.iloc[8:, :]
        
        # Set first row as headers and remove it from data
        output_transactions.columns = output_transactions.iloc[0]
        output_transactions = output_transactions.iloc[1:].reset_index(drop=True)
        
        print(f"   Output file shape: {output_transactions.shape}")
        print(f"   Output file columns: {list(output_transactions.columns)}")
        
        # Show first 5 date values from output
        if len(output_transactions.columns) > 2:  # After adding Match ID and Audit Info
            date_col = output_transactions.iloc[:, 2]  # Date column is now at index 2
            print(f"   Output file - First 5 date values:")
            for i, date_val in enumerate(date_col.head()):
                print(f"     Row {i}: {date_val} (Type: {type(date_val).__name__})")
        
    except Exception as e:
        print(f"   ERROR reading output file: {e}")
    
    print("\n" + "="*50 + "\n")
    
    # Test the actual Tally format conversion
    print("3. TESTING TALLY DATE FORMAT CONVERSION")
    
    try:
        # Read a small sample to test the conversion
        sample_df = pd.read_excel(input_file, header=None)
        sample_transactions = sample_df.iloc[8:, :]
        sample_transactions.columns = sample_transactions.iloc[0]
        sample_transactions = sample_transactions.iloc[1:].reset_index(drop=True)
        
        if len(sample_transactions.columns) > 0:
            date_col = sample_transactions.iloc[:, 0]
            
            # Test the conversion function
            def format_tally_date(date_val):
                if pd.isna(date_val):
                    return date_val
                if isinstance(date_val, str):
                    return date_val
                if hasattr(date_val, 'strftime'):
                    # Convert datetime to Tally format: '01/Jul/2024'
                    return date_val.strftime('%d/%b/%Y')
                return date_val
            
            print("   Testing conversion from pandas datetime to Tally format:")
            for i, date_val in enumerate(date_col.head(3)):
                converted = format_tally_date(date_val)
                print(f"     Original: {date_val} ({type(date_val).__name__})")
                print(f"     Converted: {converted} ({type(converted).__name__})")
                print()
                
    except Exception as e:
        print(f"   ERROR testing conversion: {e}")

if __name__ == "__main__":
    test_date_formats()
