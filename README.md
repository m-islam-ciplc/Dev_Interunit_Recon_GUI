# Excel Transaction Matcher

A Python script for matching LC (Letter of Credit) transactions between two Excel files and generating matched output files with audit information.

## Features

- **Automatic LC Number Extraction**: Uses regex patterns to find LC numbers in transaction descriptions
- **Transaction Block Identification**: Identifies logical transaction blocks within Excel files
- **Smart Matching**: Matches transactions based on LC numbers and amount validation
- **Audit Trail**: Generates detailed audit information for each match
- **Flexible Configuration**: Easy-to-modify configuration section at the top of the script
- **Command Line Options**: Override configuration via command line arguments

## Quick Start

1. **Update Configuration**: Modify the variables at the top of `excel_transaction_matcher.py`
2. **Run the Script**: `python excel_transaction_matcher.py`
3. **Check Output**: Find matched files in the specified output folder

## Configuration Options

### File Paths
```python
# Input file paths - Update these to point to your Excel files
INPUT_FILE1_PATH = "Short Data/LC Interunit Steel.xlsx"
INPUT_FILE2_PATH = "Short Data/LC Interunit GeoTex.xlsx"

# Output folder - Where to save the matched files
OUTPUT_FOLDER = "Short Data"
```

### File Naming
```python
# File naming patterns for output files
OUTPUT_SUFFIX = "_MATCHED.xlsx"    # Suffix for matched files
SIMPLE_SUFFIX = "_SIMPLE.xlsx"     # Suffix for simple test files
```

### Processing Options
```python
# Processing options
CREATE_SIMPLE_FILES = True   # Set to False if you don't want simple test files
CREATE_ALT_FILES = False     # Set to False if you don't want alternative files
VERBOSE_DEBUG = True         # Set to False to reduce debug output
```

### LC Pattern
```python
# LC Number extraction pattern (modify if your LC numbers have different format)
LC_PATTERN = r'\b(?:L/C|LC)[-\s]?\d+[/\s]?\d*\b'
```

## Command Line Usage

### Show Current Configuration
```bash
python excel_transaction_matcher.py --config
```

### Show Configuration Help
```bash
python excel_transaction_matcher.py --help-config
```

### Override Configuration
```bash
# Use different input files
python excel_transaction_matcher.py --file1 "path/to/file1.xlsx" --file2 "path/to/file2.xlsx"

# Use different output folder
python excel_transaction_matcher.py --output "path/to/output/folder"

# Combine options
python excel_transaction_matcher.py --file1 "file1.xlsx" --file2 "file2.xlsx" --output "output"
```

## Input File Structure

The script expects Excel files with the following structure:

- **Rows 1-8**: Metadata (company info, dates, etc.)
- **Row 9**: Column headers (Date, Particulars, Description, Vch Type, Vch No., Debit, Credit)
- **Row 10**: Opening balance
- **Row 11+**: Transaction data

### Transaction Block Format
- **Block Start**: Row with date and Dr/Cr in Particulars column
- **Block End**: Row with "Entered By :" in Particulars column
- **LC Numbers**: Found in Description column (Column C)

## Output Files

### Main Output Files
- `{filename}_MATCHED.xlsx`: Contains original data plus Match ID and Audit Info columns
- Preserves original file structure with metadata

### Test Files (Optional)
- `{filename}_SIMPLE.xlsx`: Simple version without metadata (for testing)
- `{filename}_ALT_SIMPLE.xlsx`: Alternative DataFrame version (if enabled)

## Output Columns

### Match ID Column
- Sequential IDs: M001, M002, M003, etc.
- Same ID appears in both files for corresponding matches

### Audit Info Column
JSON-formatted string containing:
```json
{
  "match_type": "LC",
  "match_method": "reference_match",
  "lc_number": "LC-123456789/25",
  "lender_amount": "1000000.00",
  "borrower_amount": "1000000.00"
}
```

## Example Usage

### Basic Usage
```bash
# Update configuration in script, then run
python excel_transaction_matcher.py
```

### Command Line Override
```bash
# Process different files without changing script
python excel_transaction_matcher.py --file1 "Data/File1.xlsx" --file2 "Data/File2.xlsx" --output "Results"
```

### Check Configuration
```bash
# See current settings
python excel_transaction_matcher.py --config

# Get help on configuration options
python excel_transaction_matcher.py --help-config
```

## Troubleshooting

### Common Issues

1. **"Input file not found"**: Check file paths in configuration
2. **"No matches found"**: Verify LC numbers exist in Description column
3. **"Permission denied"**: Close Excel files before running script
4. **"Index out of bounds"**: Check if input files have expected structure

### Debug Mode
- Set `VERBOSE_DEBUG = True` for detailed output
- Set `VERBOSE_DEBUG = False` for minimal output

### File Structure Issues
- Ensure input files follow the expected structure
- Check that LC numbers are in the Description column
- Verify transaction blocks are properly formatted

## Requirements

- Python 3.7+
- pandas
- openpyxl (for Excel file handling)

## Installation

```bash
pip install pandas openpyxl
```

## License

This script is provided as-is for educational and business use.
