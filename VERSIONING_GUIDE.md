# Versioning Guide for 13WCF-Activity Rollforward App

## ✅ Current Version: v1.0

**Main File:** `13WCF-Activity Rollforward App v1.0.py`

---

## 📋 How to Update the App (Step-by-Step)

### Step 1: Determine New Version Number

Check `VERSION` file for current version, then decide:

- **Bug fix or small change?** → Increment MINOR (1.0 → 1.1)
- **Major feature or breaking change?** → Increment MAJOR (1.9 → 2.0)

### Step 2: Archive Current Version

```bash
# Create dated archive folder if it doesn't exist
mkdir -p "Archive/Old_Versions_$(date +%Y-%m-%d)"

# Move current version to archive
mv "13WCF-Activity Rollforward App v1.0.py" "Archive/Old_Versions_$(date +%Y-%m-%d)/"
```

### Step 3: Save New Version

Save your updated code as:
```
13WCF-Activity Rollforward App v1.1.py
```

### Step 4: Update References

**A. Update START_PYTHON_APP.command:**
```bash
python3 "13WCF-Activity Rollforward App v1.1.py"
```

**B. Update VERSION file:**
```
1.1

# Version History

## v1.1 (YYYY-MM-DD)
- [List your changes here]
- Bug fix: ...
- Enhancement: ...

## v1.0 (2026-03-03)
- Initial versioned release
...
```

**C. Update CLAUDE.md:**
Replace all instances of `v1.0` with `v1.1`

### Step 5: Commit to Git

```bash
git add .
git commit -m "Version 1.1: [description of changes]

Changes:
- [List main changes]
- [List other changes]

Co-Authored-By: Claude Sonnet 4.5 <noreply@anthropic.com>
"

git tag -a v1.1 -m "Version 1.1: [brief description]"
git push origin main --tags
```

---

## 🎯 Quick Reference

### File Naming Convention
```
13WCF-Activity Rollforward App v{MAJOR}.{MINOR}.py
```

### Files to Update on Version Change
1. ✅ Main app file (new version)
2. ✅ `START_PYTHON_APP.command`
3. ✅ `VERSION` file
4. ✅ `CLAUDE.md`
5. ✅ Git commit + tag

### Archive Structure
```
Archive/
├── Old_Versions_YYYY-MM-DD/  ← Previous app versions
├── Old_Test_Files/            ← Test scripts
├── Python_Old/                ← Old Python implementations
└── React/                     ← Old React implementation
```

---

## 📝 Example Version Update

**Scenario:** Fixed a bug in formula adjustment

**Before:** v1.0  
**After:** v1.1  

**Commands:**
```bash
# 1. Archive old version
mkdir -p "Archive/Old_Versions_2026-03-04"
mv "13WCF-Activity Rollforward App v1.0.py" "Archive/Old_Versions_2026-03-04/"

# 2. Save new version as v1.1

# 3. Update START_PYTHON_APP.command
# Change: python3 "13WCF-Activity Rollforward App v1.0.py"
# To:     python3 "13WCF-Activity Rollforward App v1.1.py"

# 4. Update VERSION file (add entry for v1.1)

# 5. Commit and tag
git add .
git commit -m "Version 1.1: Fix formula adjustment bug"
git tag -a v1.1 -m "Version 1.1: Bug fix"
git push origin main --tags
```

---

## 🔍 Why Versioning?

- **Track Changes:** Know exactly what changed between versions
- **Easy Rollback:** Can revert to any previous version from Archive
- **Professional:** Clear version history for collaboration
- **Documentation:** VERSION file serves as changelog

---

## ⚠️ Important Notes

- **Always** archive before creating new version
- **Never** skip version numbers (1.0 → 1.2 ❌, should be 1.0 → 1.1 ✅)
- **Always** update START_PYTHON_APP.command or it won't launch
- **Always** commit with version tag to git

