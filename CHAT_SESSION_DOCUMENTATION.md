# Excel Transaction Matcher - Chat Session Documentation

## 📋 Project Overview

This project involves creating a Python script to match transactions between two Excel files based on LC (Letter of Credit) numbers. The script identifies matching transactions between two files and assigns unique Match IDs, ensuring both files have consistent Match IDs for the same transactions.

## 🎯 Core Requirements

1. **Match Transactions**: Find transactions with matching LC numbers between two Excel files
2. **Consistent Match IDs**: Both files must have the exact same Match IDs for the same matches
3. **Sequential Numbering**: Match IDs should start from M001 and be sequential
4. **Output Files**: Generate matched Excel files with Match ID and Audit Info columns

## 🔧 Technical Architecture

### Main Script: `excel_transaction_matcher.py`
- **Configuration Section**: Centralized settings for file paths, patterns, and options
- **ExcelTransactionMatcher Class**: Main logic for processing files and finding matches
- **LC Matching Logic**: Extracts and matches LC numbers between files

### Supporting Files
- **`lc_matching_logic.py`**: Handles LC-specific extraction and matching rules
- **Input Files**: Excel files containing transaction data with LC numbers
- **Output Files**: Generated matched files with Match IDs and audit information

## 📁 File Structure

```
Project/
├── excel_transaction_matcher.py      # Main script
├── lc_matching_logic.py             # LC matching logic
├── Input Files/                     # Input Excel files
│   ├── Interunit GeoTex.xlsx
│   └── Interunit Steel.xlsx
└── Output/                          # Generated output files
    ├── Interunit GeoTex_MATCHED.xlsx
    └── Interunit Steel_MATCHED.xlsx
```

## 🔍 Key Features Implemented

### 1. LC Number Extraction
- **Pattern Matching**: Uses regex pattern `r'\b(?:L/C|LC)[-\s]?\d+[/\s]?\d*\b'`
- **Format Support**: Handles LC-123/456, L/C-123/456, LC 123/456 formats
- **Narration Row Detection**: Identifies LC numbers in non-bold description rows

### 2. Transaction Block Identification
- **Format-Based Detection**: Uses openpyxl to check bold formatting in Column C
- **Header Recognition**: Identifies transaction blocks with Date + Dr/Cr + Bold Description
- **Parent-Child Linking**: Links narration rows to their parent transaction rows

### 3. Match ID Assignment
- **Sequential Counter**: Starts from M001 and increments sequentially
- **Consistency Check**: Ensures both files receive the same Match ID for each match
- **Audit Information**: Stores match details in JSON format

### 4. Excel Output Formatting
- **Column Widths**: Sets specific widths for all columns
- **Text Wrapping**: Enables wrapping for Audit Info and Description columns
- **Date Format Preservation**: Maintains Tally date format (%d/%b/%Y)
- **Number Formatting**: Removes scientific notation from amount columns

## 🚨 Issues Encountered and Resolved

### Issue 1: Blank Output Files
**Problem**: Output Excel files were being created but with blank columns
**Root Cause**: Incorrect DataFrame indexing and column assignment
**Solution**: Fixed DataFrame access patterns and added debug verification

### Issue 2: Missing Match IDs (M001-M005)
**Problem**: LC matches were not being found, causing gaps in Match ID sequence
**Root Cause**: LC matching logic was too restrictive and failing silently
**Solution**: Simplified LC extraction logic and added comprehensive debugging

### Issue 3: Inconsistent Match IDs Between Files
**Problem**: Same transaction getting different Match IDs in different files
**Root Cause**: Duplicate matching by both LC and PO logic
**Solution**: Implemented duplicate prevention by excluding already-matched transactions

### Issue 4: LC Pattern Matching
**Problem**: LC numbers not being extracted from input files
**Root Cause**: LC extraction was looking in wrong columns/rows
**Solution**: Implemented format-based detection using openpyxl to identify narration rows

## 📊 Configuration Options

```python
# File Paths
INPUT_FILE1_PATH = "Input Files/Interunit GeoTex.xlsx"
INPUT_FILE2_PATH = "Input Files/Interunit Steel.xlsx"
OUTPUT_FOLDER = "Output"
OUTPUT_SUFFIX = "_MATCHED.xlsx"

# Processing Options
CREATE_SIMPLE_FILES = False
CREATE_ALT_FILES = False
VERBOSE_DEBUG = True

# Matching Parameters
LC_PATTERN = r'\b(?:L/C|LC)[-\s]?\d+[/\s]?\d*\b'
AMOUNT_TOLERANCE = 0.01
```

## 🔧 Command Line Usage

```bash
# Run with default configuration
python excel_transaction_matcher.py

# Show current configuration
python excel_transaction_matcher.py --config

# Override file paths
python excel_transaction_matcher.py --file1 "path/to/file1.xlsx" --file2 "path/to/file2.xlsx"

# Override output folder
python excel_transaction_matcher.py --output "CustomOutput"
```

## 📈 Processing Flow

1. **File Loading**: Read Excel files using pandas with openpyxl engine
2. **LC Extraction**: Identify narration rows and extract LC numbers
3. **Block Identification**: Find transaction blocks using formatting
4. **Matching Logic**: Compare LC numbers between files
5. **Match ID Assignment**: Assign sequential Match IDs to valid matches
6. **Output Generation**: Create matched files with consistent Match IDs
7. **Formatting**: Apply column widths, text wrapping, and date preservation

## 🧪 Testing and Validation

### Debug Output
- **LC Extraction**: Shows found LC numbers and parent row links
- **Match Creation**: Displays each match with details
- **Consistency Check**: Verifies Match IDs are identical in both files
- **DataFrame State**: Shows final state before saving

### Validation Steps
1. **Match ID Consistency**: Both files must have identical Match IDs
2. **Sequential Order**: Match IDs should be M001, M002, M003, etc.
3. **Audit Info**: Each match should have complete audit information
4. **File Integrity**: Output files should contain all expected data

## 🔍 Debugging Features

### Comprehensive Logging
- **LC Extraction Debug**: Shows which LC numbers are found and linked
- **Match Creation Debug**: Displays each match as it's created
- **File State Debug**: Shows DataFrame contents before and after operations
- **Error Tracking**: Identifies and reports any inconsistencies

### Verification Steps
- **Before/After Checks**: Verifies data state at each step
- **Consistency Validation**: Ensures both files have matching data
- **Row Index Tracking**: Monitors which rows receive which Match IDs
- **Data Integrity**: Checks for missing or corrupted data

## 📝 Key Learnings

### 1. Excel Processing
- **Format Detection**: Use openpyxl for formatting-based logic
- **Row Linking**: Properly link narration rows to transaction headers
- **Data Validation**: Verify data integrity at each processing step

### 2. Match ID Management
- **Sequential Assignment**: Ensure counter starts from correct value
- **Consistency Check**: Verify both files receive identical Match IDs
- **Duplicate Prevention**: Avoid matching same transaction multiple times

### 3. Debugging Strategy
- **Comprehensive Logging**: Add debug output at each critical step
- **State Verification**: Check data state before and after operations
- **Incremental Testing**: Test each component separately before integration

## 🚀 Future Enhancements

### Potential Improvements
1. **PO Matching**: Add Purchase Order number matching capability
2. **Configurable Patterns**: Allow user-defined regex patterns
3. **Batch Processing**: Support for multiple file pairs
4. **Advanced Validation**: Additional business rule validation
5. **Performance Optimization**: Optimize for large file processing

### Code Quality
1. **Error Handling**: Add comprehensive exception handling
2. **Unit Tests**: Implement automated testing
3. **Documentation**: Add inline code documentation
4. **Logging**: Implement proper logging framework

## 📊 Current Status

**✅ Completed Features:**
- LC number extraction and matching
- Transaction block identification
- Match ID assignment and consistency
- Excel output formatting
- Comprehensive debugging

**🔄 In Progress:**
- Final testing and validation
- Performance optimization
- Error handling improvements

**📋 Next Steps:**
- Complete testing with actual input files
- Validate Match ID consistency
- Optimize performance for large files
- Add comprehensive error handling

## 🔗 Related Files

- **Main Script**: `excel_transaction_matcher.py`
- **LC Logic**: `lc_matching_logic.py`
- **Input Files**: Excel files in `Input Files/` directory
- **Output Files**: Generated files in `Output/` directory
- **Configuration**: Settings at top of main script

## 📞 Support and Maintenance

This script was developed to solve specific Excel transaction matching requirements. For maintenance or modifications:

1. **Review Configuration**: Check settings at top of main script
2. **Test Changes**: Always test with sample data before production use
3. **Debug Output**: Use verbose debug mode to troubleshoot issues
4. **Backup Data**: Keep backups of input files before processing

---

*Documentation created from chat session on Excel Transaction Matcher project*
*Last updated: [Current Date]*
