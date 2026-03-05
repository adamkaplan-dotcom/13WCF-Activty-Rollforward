# USDx Balances Weekly Update Script

## Overview
This script automates the process of updating your master balance tracking file with weekly data from forecast files.

## What It Does

### Step 1: Read Source File
- Opens the weekly forecast Excel file
- Finds the "USDx Balances" tab
- Extracts the balance date from row 12 (header row)
- Collects all balances: reference number (column B) → current week balance (column K)
- Skips section headers (like "Bullish, CoinDesk Fiat Balances...")

### Step 2: Read Target File
- Opens your master Excel file
- Identifies reference number column (column B)
- Finds the last populated column
- Checks for existing dates to avoid duplicates

### Step 3: Match and Insert
- Adds a new column with the week's date
- Matches each ref number and inserts the balance
- Flags missing refs (exist in source but not target)
- Sets missing values to 0.0

### Step 4: Save
- Preserves all existing formatting
- Saves the updated file

### Step 5: Validation
- Shows summary statistics
- Lists any mismatches or new refs

## Installation

```bash
pip3 install openpyxl
```

## Usage

### Basic Usage (overwrites target file):
```bash
python3 update_balances.py "Weekly_Forecast.xlsx" "Master_Balances.xlsx"
```

### Save to a new file:
```bash
python3 update_balances.py "Weekly_Forecast.xlsx" "Master_Balances.xlsx" "Master_Balances_Updated.xlsx"
```

## File Requirements

### Source File (Weekly Forecast):
- Must have a sheet named "USDx Balances" (or containing "USDx")
- Row 12 should contain headers
- Column B should contain reference numbers
- Column K (or marked "Current Week") should contain balances
- Balance date should be in the header row

### Target File (Master):
- Column B should contain reference numbers
- Row 1 should contain date headers
- Each subsequent column represents a week

## Example Output

```
============================================================
STEP 1: Reading source file
============================================================
✓ Found sheet: USDx Balances
✓ Found balance date: 2026-02-20 in column K
✓ Extracted 45 balance records
✓ Skipped 3 section header rows

============================================================
STEP 2: Reading target file
============================================================
✓ Loaded workbook: Master_Balances.xlsx
✓ Reference column: B
✓ Last populated column: M
✓ Found 12 existing date columns
  Most recent: 2026-02-13

============================================================
STEP 3: Updating target file with new data
============================================================
✓ Added new column N with date: 2026-02-20
✓ Matched and updated: 42 records
⚠ Missing in source: 3 records (set to 0.0)
⚠ New refs found in source: 2 (not in target)

============================================================
STEP 4: Saving updated file
============================================================
✓ Saved to: Master_Balances_Updated.xlsx

============================================================
STEP 5: SUMMARY
============================================================
Balance date: 2026-02-20
Records matched: 42
Records missing from source: 3
New refs in source: 2

✓ Update complete!
============================================================
```

## Troubleshooting

**Error: "Could not find 'USDx Balances' sheet"**
- Check that your source file has a sheet with "USDx" in the name

**Error: "No permission"**
- Make sure the target file is not open in Excel

**Duplicate date warning:**
- The script will warn you if the date already exists
- You can choose to overwrite or abort

## Notes

- The script preserves all existing formatting in the target file
- New reference numbers found in the source are logged but not automatically added
- Missing balances are set to 0.0
- Always review the summary output before using the file
