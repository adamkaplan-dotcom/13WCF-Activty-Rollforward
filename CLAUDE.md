# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Balance Updater is a **Python Flask web application** that processes Excel files to update USDx balance data. It reads from a "Weekly Balances" file and updates an "Activity Rollforward" file with balance data and formulas.

**Key Operations:**
1. Reads Current Week balances (column K) from USDx Balances tab (K10 header)
2. Finds last date in row 14 of Beginning Balances tab
3. Pastes column K data to the next column
4. Copies and adjusts formula from row 10 for date header
5. Copies formulas for rows 12-20 and 148-end from previous column

## Development Commands

### Start the application
```bash
# Using the launcher script (recommended)
./START_PYTHON_APP.command

# Or manually
python3 -m venv python_venv
source python_venv/bin/activate
pip install -r requirements.txt
python3 13WCF-Activity Rollforward App v1.1.py
```

### Install/update dependencies
```bash
source python_venv/bin/activate
pip install -r requirements.txt
```

### Run in development mode
The app runs with `debug=True` by default on port 5000. Access at http://localhost:5000

## Architecture

### Technology Stack
- **Python 3.9+**: Core runtime
- **Flask 3.0.0**: Web framework
- **openpyxl 3.1.2**: Excel file processing
- **Werkzeug 3.0.1**: WSGI utilities

### File Structure
```
BalanceUpdater/
├── 13WCF-Activity Rollforward App v1.1.py  # Main Flask application
├── requirements.txt               # Python dependencies
├── START_PYTHON_APP.command       # Launch script (double-click this!)
├── templates/
│   └── index.html                # Web interface
├── static/                       # Static assets (empty)
├── uploads/                      # Temporary file uploads (auto-deleted)
├── outputs/                      # Processed files saved here
├── python_venv/                  # Python virtual environment
├── README.md                     # User documentation
├── CLAUDE.md                     # This file
│
└── Archive/                      # Old/archived versions
    ├── React/                    # Previous React browser-based version
    ├── Python_Old/               # Previous Python versions
    ├── Old_Test_Files/           # Test files and backups
    └── Open Python App.html      # HTML launcher (archived)
```

### Processing Flow
1. User uploads two .xlsx files via web interface
2. Files saved to uploads/ with secure filenames
3. `process_files()` function performs all Excel operations using openpyxl
4. Output saved to outputs/ with timestamp
5. Uploaded files deleted immediately
6. User downloads processed file

## Excel Processing Logic (Critical Details)

### Source File: "Weekly Balances"
- **Tab**: "USDx Balances"
- **Header row**: 10 (0-indexed in code: 9)
- **Data starts**: Row 13 (0-indexed: 12)
- **Current Week balance column**: K (0-indexed: 10)
- **Reference column**: B (0-indexed: 1)
- **Section headers to skip**: Rows containing "bullish", "coindesk", "fiat", "balance", "total" (case-insensitive)

### Target File: "Activity Rollforward"
- **Tab**: "Beginning Balances" (created if doesn't exist)
- **Date detection row**: 14 (0-indexed: 13)
- **Formula header row**: 10 (0-indexed: 9)
- **Data starts**: Row 15 (0-indexed: 14)
- **Matching column**: B (0-indexed: 1)
- **Formula copy ranges**:
  - Rows 12-20 (0-indexed: 11-19)
  - Rows 148-end (0-indexed: 147-end)

### Key Processing Steps in `process_files()`

Located in `13WCF-Activity Rollforward App v1.1.py`, lines ~25-223:

1. **Read source data** (Weekly Balances file)
   - Load with `openpyxl.load_workbook(data_only=True)` to read calculated values
   - Extract column K (Current Week balances) from USDx Balances tab
   - Build dictionary of balances keyed by Column B reference numbers

2. **Open target file** (Activity Rollforward)
   - Load with `openpyxl.load_workbook()` (preserves formulas)
   - Find or create "Beginning Balances" sheet

3. **Find last date** in row 14
   - Scan columns left to right to find last non-None value
   - New column = last date column + 1

4. **Copy formatting** from previous column to new column
   - Uses `copy.copy()` to copy style attributes
   - Includes font, border, fill, number_format, alignment, protection

5. **Copy formula** from row 10 for date header
   - Checks if source cell has formula
   - Copies formula structure to new column

6. **Paste balances** starting row 15
   - Match by Column B reference numbers
   - Set to 0.0 if reference not found in source

7. **Copy formulas** for specific row ranges
   - Rows 12-20: Header/summary calculations
   - Rows 148-end: Footer calculations
   - Copies both formulas and values

8. **Save output** with timestamp
   - Format: `Activity_Rollforward_Updated_YYYYMMDD_HHMMSS.xlsx`

### openpyxl Cell References

All use 0-indexed row/column numbers (subtract 1 from Excel row/column):
- Row 10 in Excel = row 9 in code
- Column K in Excel = column 10 in code

Key variables in code:
```python
header_row_source = 9        # K10 in Excel
data_start_row = 12          # Row 13 in Excel
balance_col = 10             # Column K in Excel
ref_col = 1                  # Column B in Excel
dateRow = 13                 # Row 14 in Excel
formulaRow = 9               # Row 10 in Excel
```

## File Handling

### Upload Constraints
- Max file size: 100MB (configurable in `13WCF-Activity Rollforward App v1.1.py` line 16)
- Allowed extensions: .xlsx, .xls
- Files must have required tabs (validation during processing)

### Security
- Uses `secure_filename()` for all uploaded files
- Files deleted immediately after processing
- No external network calls
- Runs locally only

## API Endpoints

Located in `13WCF-Activity Rollforward App v1.1.py`:

- `GET /` - Main upload interface (renders templates/index.html)
- `POST /upload` - Process files (expects multipart form with weekly_file and rollforward_file)
- `GET /download/<filename>` - Download processed file from outputs/
- `GET /outputs` - List all output files with metadata

## Common Modifications

### Changing Excel Row/Column References

All hardcoded references are in `process_files()` function (13WCF-Activity Rollforward App v1.1.py):
- `header_row_source = 9` (line ~43) - Excel row 10
- `data_start_row = 12` (line ~44) - Excel row 13
- `balance_col = 10` (line ~68) - Excel column K
- `ref_col = 1` (line ~69) - Excel column B
- `dateRow = 13` (line ~275) - Excel row 14
- `formulaRow = 9` (line ~302) - Excel row 10

Formula copy ranges:
- `range(11, 20)` (line ~369) - Excel rows 12-20
- `range(147, endRow)` (line ~391) - Excel rows 148-end

### Modifying Section Header Skip Logic

Keywords checked in line ~79:
```python
if not any(keyword in ref_str.lower() for keyword in ['bullish', 'coindesk', 'fiat', 'balance', 'total']):
```

### Changing Output Naming

Timestamp format at line ~179:
```python
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
```

### Adding New File Validations

Modify `allowed_file()` function (line ~21) or add validation in `process_files()` before opening workbooks.

## Debugging Tips

### Flask Debug Mode
Debug mode is enabled by default (13WCF-Activity Rollforward App v1.1.py line ~316):
```python
app.run(debug=True, host='0.0.0.0', port=5000)
```

This provides:
- Detailed error pages in browser
- Auto-reload on code changes
- Interactive debugger

### Logging
Processing logs are sent to browser via `log.append()` messages.
Check terminal output for Flask server logs.

### Common Errors

**"Sheet not found" errors:**
- Check tab names exactly match: "USDx Balances", "Beginning Balances"
- Tab names are case-sensitive

**"No dates found in row 14":**
- Verify row 14 has date values in target file
- Check that row numbering hasn't changed

**Large file crashes:**
- Python handles files up to 100MB+ well
- If issues persist, increase system memory or use file streaming

## Archived Versions

### Archive/React/
Previous React/JSX browser-based version with:
- ZIP-level XLSX manipulation for perfect formatting preservation
- Browser-only processing (no server required)
- **Issue**: Crashes on files >35 MB due to browser memory limits

### Archive/Python_Old/
Previous iterations of the Python version with different processing logic.

## Performance Notes

- Python/Flask handles large files (50-100 MB) better than browser-based solutions
- openpyxl is memory-efficient for typical Excel files
- Processing time scales linearly with row count
- Typical processing time: 5-15 seconds for files with 1000-5000 rows

## Security Considerations

- Runs on localhost only (0.0.0.0:5000)
- No data sent to external servers
- Files auto-deleted after processing
- Uses secure filename sanitization
- No authentication (local use only)

## Versioning Protocol (IMPORTANT)

### Current Version
**v1.1** - `13WCF-Activity Rollforward App v1.1.py`

### When Saving/Updating the App

**ALWAYS follow this versioning protocol:**

1. **Read current version** from `VERSION` file
2. **Determine new version number:**
   - Increment MINOR (1.0 → 1.1) for: bug fixes, small features, minor improvements
   - Increment MAJOR (1.9 → 2.0) for: significant features, breaking changes, major rewrites
3. **Archive old version:**
   ```bash
   mv "13WCF-Activity Rollforward App v1.1.py" "Archive/Old_Versions_YYYY-MM-DD/"
   ```
4. **Save new version:**
   ```bash
   # Save as: 13WCF-Activity Rollforward App v1.1.py
   ```
5. **Update references:**
   - `START_PYTHON_APP.command` - update python3 command
   - `VERSION` file - add entry with changes
   - `CLAUDE.md` - update filename references
6. **Commit to git:**
   ```bash
   git add .
   git commit -m "Version 1.1: [description of changes]"
   git tag -a v1.1 -m "Version 1.1"
   git push origin main --tags
   ```

### File Naming Convention
```
13WCF-Activity Rollforward App v{MAJOR}.{MINOR}.py
```

**Examples:**
- `13WCF-Activity Rollforward App v1.1.py` - Initial release
- `13WCF-Activity Rollforward App v1.1.py` - Bug fixes
- `13WCF-Activity Rollforward App v2.0.py` - Major update

### Why This Matters
- Maintains version history in Archive
- Tracks changes over time
- Allows rollback to previous versions if needed
- Ensures consistency across all launcher files
