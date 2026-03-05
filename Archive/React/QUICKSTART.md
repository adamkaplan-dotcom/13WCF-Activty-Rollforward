# ⚡ Quick Start Guide

## 1️⃣ Install Node.js (First Time Only)

**If you already have Node.js installed, skip to step 2.**

1. Visit https://nodejs.org/
2. Download the **LTS** version (recommended)
3. Run the installer
4. Restart your terminal/computer

To verify installation, open Terminal and run:
```bash
node --version
```
You should see something like `v18.x.x` or higher.

## 2️⃣ Start the App

### Easiest Way (Double-Click)

1. **Double-click** `start_react_app.command`
2. Wait 30-60 seconds (first time will install dependencies)
3. Your browser will automatically open to **http://localhost:5173**

### Terminal Way

```bash
cd ~/Desktop/Claude/13WCF/BalanceUpdater
npm install  # First time only
npm run dev
```

## 3️⃣ Use the App

1. **Upload Weekly Balances file** (click "Choose File")
   - Must have "USDx Balances" tab

2. **Upload Activity Rollforward file** (click "Choose File")
   - Must have "Beginning Balances" tab

3. **Click "Process Files"** button

4. **Download your file** when processing completes

## 4️⃣ Stop the Server

Press `CTRL+C` in the terminal window

---

## 🆘 Troubleshooting

### "Node.js is not installed"
→ Install from https://nodejs.org/, then restart terminal

### "Port 5173 already in use"
→ Another instance is running. Close it first, or change port in `vite.config.js`

### Files won't process
→ Check:
- Files are `.xlsx` format (not `.xls`)
- "USDx Balances" tab exists in Weekly file
- "Beginning Balances" tab exists in Rollforward file

### Browser doesn't open
→ Manually navigate to: **http://localhost:5173**

---

## 📚 More Information

- **Full documentation**: See `README_REACT.md`
- **Developer guide**: See `CLAUDE.md`
- **Old Python version**: Archived in `Archive/` folder
