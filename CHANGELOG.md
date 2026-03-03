# Balance Updater - Changelog

## v1.2 (March 3, 2026) - Bug Fix: Cockpit Column Sync

### Bug Fixes
🐛 **Fixed Cockpit Tab Column Synchronization**
- **Issue**: Cockpit tab was searching for its own last column instead of using the same column as Beginning Balances
- **Fix**: Cockpit now uses the exact same column number as Beginning Balances (e.g., if Beginning Balances uses BK, Cockpit also uses BK)
- **Impact**: Ensures both tabs stay in sync and use matching columns for the same week's data

### Technical Changes
- Changed Cockpit logic to use `new_col` (from Beginning Balances) instead of searching independently
- Copies from `new_col - 1` to `new_col` (e.g., BJ → BK)
- Removed redundant column search logic in Cockpit processing
- BTH integration now correctly uses `cockpit_new_col` variable

### Result
✅ Beginning Balances column BK ← Weekly Balances data
✅ Cockpit column BK ← Formulas copied from BJ
✅ Cockpit cell BK20 ← BTH Financing total (if BTH file provided)

## v1.1 (March 3, 2026) - Cockpit Tab & BTH Integration

### New Features
✅ **Cockpit Tab Processing**
- Finds last column with data in Cockpit tab
- Copies formulas from last column to next column (just like Beginning Balances)
- Preserves formatting from previous column
- Example: Column BJ formulas → Column BK

✅ **BTH Investments Financing Schedule Integration**
- Added optional third file upload
- Reads Summary tab from BTH file
- Matches date in row 30 (BTH Financing dates)
- Extracts Total BTH Financing from row 42
- Pastes to Cockpit tab, row 20, in new column
- Example: 564,519,892 pasted to BK20

### Technical Updates
- Updated `process_files()` to accept optional `bth_file_path` parameter
- Added STEP 8: Processing Cockpit Tab
- Added STEP 9: Processing BTH Investments File
- Updated upload endpoint to handle 3 files (2 required, 1 optional)
- Updated UI with 3-column grid layout
- Enhanced JavaScript to manage BTH file upload

### UI Changes
- Changed upload grid from 2 columns to 3 columns
- Added responsive breakpoint for mobile (stacks vertically)
- BTH file field marked as optional with lighter opacity
- Maintains all existing functionality for the two required files

## v1.0 (March 3, 2026) - Initial Versioned Release

### ✅ Formula Adjustment Function - FIXED
- **Problem**: Previous formula adjustment was modifying function names (INDEX became INDEY, SUM became SUN)
- **Solution**: Implemented two-pass regex approach:
  - First pass: Matches cell references with row numbers (CC14, $A$1, etc.)
  - Second pass: Matches column-only ranges ($B:$B, A:Z, etc.)
- **Result**: All test cases now pass correctly
  - ✓ `=CC14+7` → `=CD14+7`
  - ✓ `=INDEX($B:$B,MATCH(CC14,$A:$A,0))` → `=INDEX($B:$B,MATCH(CD14,$A:$A,0))`
  - ✓ `=$CC$14+7` → `=$CC$14+7` (absolute columns stay fixed)
  - ✓ `=SUM(CC14:CC20)` → `=SUM(CD14:CD20)`

### ✅ Special Reference Handling - NEW FEATURE
- **References 11, 60, 76, and 91** now copy values from previous week instead of Weekly Balances file
- **Applies to**: Rows 15-148 in Beginning Balances tab
- **Example**: Cell CD31 (ref 11) will copy the value from CC31
- **Logging**: Shows which references were copied from previous week

### Files Renamed
- `app.py` → `13WCF- Balance Updater App.py`
- `OPEN IN CHROME.command` → `13WCF - Balance Updater.command`

### Files Restored/Created
- `START_PYTHON_APP.command` - Launcher for Flask server
- `SHOW QR CODE.command` - Shows mobile access QR code
- `Mobile Access - Scan QR Code.html` - QR code page for network access

### Processing Logic Updates
**STEP 6 - Pasting Balances:**
- Regular references: Paste from Weekly Balances file
- Special references (11, 60, 76, 91) in rows 15-148: Copy from previous week's column
- All others not found: Set to 0.0

**STEP 7 - Formula Copying:**
- Rows 12-20: Copy and adjust formulas from previous column
- Rows 148-239: Copy and adjust formulas from previous column
- Formula adjustment preserves absolute references ($A, $B:$B)
- Formula adjustment updates relative references (CC → CD)

### Network Access
- Local IP: 10.5.41.228
- Access from other devices: http://10.5.41.228:5000
- QR code available for mobile devices

## File Structure
```
BalanceUpdater/
├── 13WCF-Activity Rollforward App v1.0.py  ⭐ Main application (versioned)
├── START_PYTHON_APP.command                ⭐ Start server
├── 13WCF - Balance Updater.command         ⭐ Open in Chrome
├── SHOW QR CODE.command
├── Mobile Access - Scan QR Code.html
├── VERSION                                 📋 Version tracking
├── CHANGELOG.md                            📋 Change history
├── requirements.txt
├── templates/
├── static/
├── uploads/
├── outputs/
├── python_venv/
├── README.md
├── CLAUDE.md
└── Archive/
    ├── Old_Versions_2026-03-03/
    ├── Old_Test_Files/
    ├── Python_Old/
    └── React/
```

## Quick Start
1. Double-click `START_PYTHON_APP.command` to start the server
2. Double-click `13WCF - Balance Updater.command` to open in Chrome
3. Upload your files and process!

## Mobile Access
1. Make sure server is running
2. Double-click `SHOW QR CODE.command`
3. Scan QR code with mobile device
4. Access app from phone/tablet
