# Balance Updater React App

A browser-based React application for processing USDx balance files and updating Activity Rollforward files with **full formatting preservation**.

## 🚀 Quick Start

### Prerequisites

- **Node.js 18+** (Download from https://nodejs.org/)

### Option 1: Double-Click to Launch (Easiest)

1. **Double-click** `start_react_app.command`
2. Wait for the server to start (first time will install dependencies)
3. Your browser will open automatically to: **http://localhost:5173**
4. Drag & drop your files and click "Process Files"

### Option 2: Manual Start

```bash
cd ~/Desktop/Claude/13WCF/BalanceUpdater
npm install
npm run dev
```

Then open your browser to: **http://localhost:5173**

## 📋 What It Does

The app automatically:

1. **Reads** Column K (Current Week) from USDx Balances tab (header in K10)
2. **Finds** the last date in row 14 of Beginning Balances tab
3. **Pastes** Column K data into the next column to the right
4. **Copies** the formula from row 10 to generate a new date header
5. **Copies** formulas from previous column for:
   - Rows 12-20 (header/summary rows)
   - Rows 148-end (footer calculations)
6. **Preserves** ALL formatting, styles, colors, borders, and number formats
7. **Saves** the updated file with a timestamp

## 🎯 How to Use

### Step 1: Start the Application

Double-click `start_react_app.command` or run `npm run dev`

### Step 2: Upload Files

1. **Click** "Choose File" for **Weekly Balances file**
   - Example: `13WCF Weekly Balances - 02.27.2026 Forecast.xlsx`
   - Must have "USDx Balances" tab

2. **Click** "Choose File" for **Activity Rollforward file**
   - Example: `Bullish - Activity Rollforward_ 02.20.2026.xlsx`
   - Must have "Beginning Balances" tab

### Step 3: Process

1. Click the **"Process Files"** button
2. Wait for processing (usually 5-10 seconds)
3. View the results and statistics in real-time

### Step 4: Download

Click the **"Download Updated File"** button to get your processed file

## 📁 File Structure

```
BalanceUpdater/
├── src/
│   ├── main.jsx              # React entry point
│   └── BalanceUpdater.jsx    # Main app component
├── index.html                # HTML template
├── package.json              # Dependencies
├── vite.config.js            # Build configuration
├── start_react_app.command   # Launch script (double-click this!)
├── README_REACT.md           # This file
│
├── Archive/                  # Old Python version
│   ├── app.py
│   ├── start_app.command
│   └── ...
```

## 🔍 Technical Features

### Advanced Formatting Preservation

Unlike standard Excel libraries, this app uses:

- **ZIP-level manipulation** of XLSX files
- **Style transplantation** from original file
- **Cell-by-cell style index restoration**
- **Formula reference adjustment** when copying columns

This ensures 100% formatting preservation including:
- ✅ Cell colors and fills
- ✅ Borders and gridlines
- ✅ Number formats
- ✅ Fonts and text styles
- ✅ Column widths
- ✅ Conditional formatting

### Processing Details

**Source File: Weekly Balances**
- Tab: "USDx Balances"
- Header: K10 ("Current Week")
- Data: Column K, rows 13+
- Reference: Column B

**Target File: Activity Rollforward**
- Tab: "Beginning Balances"
- Find last date: Row 14
- Paste to: Next column right
- Match by: Column B reference numbers
- Formula ranges copied: Rows 12-20, 148-end

## 📊 What Gets Updated

### Input Files:
- **Weekly Balances File**: Contains USDx Balances tab with Current Week data in column K
- **Activity Rollforward File**: Your master tracking file with Beginning Balances tab

### Output:
- **Updated Rollforward File**: Downloaded with timestamp
  - New column added with formula-generated date header
  - Current Week balances pasted and matched by reference
  - All formulas copied and adjusted from previous week
  - **All original formatting perfectly preserved**

## ⚙️ Technical Stack

- **React 18**: Modern UI framework
- **Vite**: Fast build tool and dev server
- **SheetJS (XLSX)**: Excel file reading/writing
- **Browser APIs**: Native compression/decompression
- **ZIP manipulation**: Low-level XLSX format handling

## 🛑 Stopping the Server

Press **CTRL+C** in the terminal window to stop the development server

## 📝 Notes

- Runs entirely in your browser - no data sent to external servers
- No files are stored - processing happens in memory
- Works offline after initial load
- Cross-platform (Mac, Windows, Linux)
- Modern browsers required (Chrome, Firefox, Safari, Edge)

## 🐛 Troubleshooting

**"Node.js is not installed" error:**
- Download and install from https://nodejs.org/
- Restart your terminal after installation

**"Dependencies not installed" error:**
- Run: `npm install` in the BalanceUpdater directory
- Or just double-click `start_react_app.command` again

**Browser doesn't open automatically:**
- Manually open: http://localhost:5173

**Files not processing:**
- Check that "USDx Balances" tab exists in Weekly file
- Check that "Beginning Balances" tab exists in Rollforward file
- Ensure files are `.xlsx` format (not `.xls`)
- Check browser console (F12) for detailed errors

**Formatting not preserved:**
- This shouldn't happen! The app uses advanced ZIP manipulation
- If you see this, check the browser console for errors

## 🔒 Security

- Runs completely offline in your local browser
- No data is uploaded to external servers
- All processing happens in browser memory
- No files are saved to disk except your download
- Only you have access to the application

## 🆚 Comparison with Python Version

### React App (NEW)
✅ **Better formatting preservation** (ZIP-level manipulation)
✅ **No Python/server required** (runs in browser)
✅ **More portable** (works anywhere with Node.js)
✅ **Modern UI** (responsive, real-time updates)

### Python Flask App (OLD)
✅ **Simpler codebase** (easier to understand)
✅ **Familiar technology** (if you know Python)

Both versions produce the same business logic results, but the React version preserves formatting better.

## 📧 Support

For issues or questions:
- Check the "Processing Log" in the app for detailed error messages
- Review the browser console (F12 → Console tab)
- Ensure input files have the required tabs

---

**Version:** 2.0 (React)
**Last Updated:** March 3, 2026
