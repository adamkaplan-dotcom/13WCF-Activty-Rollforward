# Balance Updater Web Application

A local web application for processing USDx balance files and updating Activity Rollforward files.

## 🚀 Quick Start

### Option 1: Double-Click to Launch (Easiest)

1. **Double-click** `start_app.command`
2. Wait for the server to start
3. Open your browser to: **http://localhost:5000**
4. Drag & drop your files and click "Process Files"

### Option 2: Manual Start

```bash
cd ~/Desktop/BalanceUpdater
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

Then open your browser to: **http://localhost:5000**

## 📋 What It Does

The app automatically:

1. **Reads** the USDx Balances tab from your Weekly Balances file
2. **Pastes** all USDx Balance data into the "Beginning Balances" tab of the Activity Rollforward file
3. **Extracts** Current Week balances from Column K
4. **Adds** a new column on the right with the current week's balances
5. **Matches** data by reference numbers in Column B
6. **Saves** the updated file with a timestamp

## 🎯 How to Use

### Step 1: Start the Application

Double-click `start_app.command` or run manually (see Quick Start above)

### Step 2: Upload Files

1. **Drop or click** the left box to upload your **Weekly Balances file**
   - Example: `13WCF Weekly Balances - 02.27.2026 Forecast.xlsx`

2. **Drop or click** the right box to upload your **Activity Rollforward file**
   - Example: `Bullish - Activity Rollforward_ 02.20.2026.xlsx`

### Step 3: Process

1. Click the **"Process Files"** button
2. Wait for processing (usually 5-10 seconds)
3. View the results and statistics

### Step 4: Download

Click the **"Download Updated File"** button to get your processed file

## 📁 File Structure

```
BalanceUpdater/
├── app.py                  # Flask application
├── requirements.txt        # Python dependencies
├── start_app.command       # Launch script (double-click this!)
├── README.md              # This file
├── templates/
│   └── index.html         # Web interface
├── uploads/               # Temporary uploads (auto-deleted)
└── outputs/               # Processed files saved here
```

## 📊 What Gets Updated

### Input Files:
- **Weekly Balances File**: Contains USDx Balances tab with current week data
- **Activity Rollforward File**: Your master tracking file

### Output:
- **Updated Rollforward File**: Saved in `outputs/` folder with timestamp
  - Beginning Balances tab populated with USDx data
  - New column added with Current Week balances
  - All matched by Column B references

## 🔍 Features

✅ **Drag & Drop Interface** - Easy file upload
✅ **Automatic Processing** - No manual steps needed
✅ **Real-time Feedback** - See processing logs live
✅ **Statistics Dashboard** - View match rates and counts
✅ **Timestamped Outputs** - Never overwrite previous files
✅ **Error Handling** - Clear error messages if something goes wrong

## ⚙️ Technical Details

### Requirements:
- Python 3.9 or higher
- Flask 3.0.0
- openpyxl 3.1.2

### Processing Logic:
1. Reads from row 12 (headers) and row 13+ (data) in USDx Balances
2. Skips section header rows (contains "Bullish", "CoinDesk", etc.)
3. Matches on Column B (reference numbers)
4. Sets unmatched records to 0.0
5. Preserves all existing data and formatting

## 🛑 Stopping the Server

Press **CTRL+C** in the terminal window to stop the server

## 📝 Notes

- Files are automatically deleted after processing
- Outputs are saved in the `outputs/` folder
- Maximum file size: 100MB per file
- Only `.xlsx` and `.xls` files are supported
- The app runs locally - no internet connection needed
- No data is sent to external servers

## 🐛 Troubleshooting

**"Port already in use" error:**
- Another instance is running. Close it first or change the port in `app.py`

**"Module not found" error:**
- Run: `pip install -r requirements.txt`

**Files not processing:**
- Check that your files have the "USDx Balances" tab
- Ensure Column B has reference numbers
- Verify the files are `.xlsx` format

## 🔒 Security

- Runs completely offline on your local machine
- No data is uploaded to external servers
- Uploaded files are deleted immediately after processing
- Only you have access to the web interface

## 📧 Support

For issues or questions, check the processing log in the web interface for detailed error messages.

---

**Version:** 1.0
**Last Updated:** March 3, 2026
