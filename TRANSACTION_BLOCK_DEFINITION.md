# Transaction Block Identification Logic

## Overview
A transaction block is a group of consecutive rows that represent a single financial transaction in the Excel file.

## Transaction Block Structure

### 1. **Block Start (First Row)**
- **Column A**: Contains a Date Value
- **Column B**: Contains "Dr" or "Cr"
- **Column C**: Contains the first ledger entry in **BOLD TEXT**

### 2. **Block Content (Middle Rows)**
- **Column A**: Date values
- **Column B**: Dr/Cr indicators
- **Column C**: 
  - **Ledger Info**: Multiple rows with **BOLD TEXT** (stacked one after another)
  - **Narration**: Exactly 1 row with **NORMAL TEXT** (not bold, not italic)
  - **Person Name**: 1 row with **BOLD TEXT** (the person who entered the transaction)

### 3. **Block End (Last Row)**
- **Column A**: Date value
- **Column B**: Contains "Entered By :"
- **Column C**: Contains the person's name in **BOLD TEXT**

## Example Structure

```
Row 1: [01/Jan/2025] | [Cr] | [LEDGER INFO 1 - BOLD]     ← Block starts
Row 2: [01/Jan/2025] | [Cr] | [LEDGER INFO 2 - BOLD]     ← More ledger info
Row 3: [01/Jan/2025] | [Cr] | [LEDGER INFO 3 - BOLD]     ← More ledger info
Row 4: [01/Jan/2025] | [Cr] | [NARRATION - NORMAL TEXT]  ← Narration
Row 5: [01/Jan/2025] | [Entered By :] | [PERSON NAME]     ← Block ends
```

## Key Rules

1. **Block Start**: Date + Dr/Cr in Columns A & B
2. **Block End**: "Entered By :" in Column B
3. **Ledger Info**: Always in **BOLD TEXT** in Column C
4. **Narration**: Always in **NORMAL TEXT** in Column C
5. **Person Name**: Always in **BOLD TEXT** in Column C
6. **Block Span**: Multiple rows from start to end

## Purpose
This logic ensures that each transaction block gets its own unique Match ID and Audit Info, preventing multiple unrelated transactions from being grouped together under the same identifier.
