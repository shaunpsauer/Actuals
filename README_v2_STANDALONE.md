# SAP to HeavyBid Converter v2.0 - Standalone Edition

## 🎉 What's New in v2.0

**No external Excel file required!** All WBS operations data is now embedded directly in the Python scripts.

### Simple Two-File Setup
You only need:
1. ✅ `sap_to_heavybid.py` - Main transformation script
2. ✅ `wbs_operations_mapper.py` - WBS operations dictionary (embedded data)

**That's it!** No more `Gas_Transmission_WBS_and_Operations_Dictionary_REFERENCE.xlsx` needed.

---

## 🚀 Perfect for Project Templates

Copy these 2 files into your project template folder:
```
Project_Template/
├── sap_to_heavybid.py
└── wbs_operations_mapper.py
```

When you create a new project, the scripts are automatically there and ready to use!

---

## Quick Start

### 1. Copy Files
Put both Python files in your project folder:
- `sap_to_heavybid.py`
- `wbs_operations_mapper.py`

### 2. Install Python Libraries (one-time)
```bash
pip install pandas openpyxl
```

### 3. Run

**Just run the script:**
```bash
python sap_to_heavybid.py
```

You'll see a welcome message, then:
1. **File picker opens** - Select your SAP export file
2. **Folder picker opens** - Select where to save the output
3. **Output file is created** - Automatically named `<Order>_actuals.xlsx` in your chosen folder

**Done!** No typing, no configuration, no external files - just works.

---

## Features

✅ **Fully Standalone** - No Excel dependencies
✅ **Portable** - Copy 2 files, works anywhere
✅ **Template-Ready** - Perfect for project folder templates
✅ **Fast** - Embedded data loads instantly
✅ **Stable** - No missing file errors
✅ **Adaptive** - Works with any SAP export structure
✅ **Complete** - All 74 WBS operations included

---

## Usage Examples

### Standard Usage (File & Folder Pickers)
```bash
python sap_to_heavybid.py
```

**What happens:**
1. Welcome banner appears: "SAP Actuals to HeavyBid"
2. **File picker opens** - Select your SAP export file (`.xlsx` or `.xls`)
3. Order number is automatically extracted from the file
4. **Folder picker opens** - Select where to save the output file
5. Output file is automatically named `<Order>_actuals.xlsx` (e.g., `74051900_actuals.xlsx`)
6. If file already exists in that folder, adds timestamp: `74051900_actuals_20241216_143022.xlsx`
7. Transformation runs and creates 3 sheets in the output file

**You can cancel at any time** by closing either picker dialog.

### Multiple Projects
```bash
# Project 1
cd C:\Projects\Job_74051900
python sap_to_heavybid.py
# Select export file, select output folder, done!

# Project 2
cd C:\Projects\Job_74052100
python sap_to_heavybid.py
# Same simple process
```

### From Any Location
```bash
# Run from anywhere - pickers let you navigate to files/folders
python C:\MyProjects\scripts\sap_to_heavybid.py
```

---

## Output Format

Your output Excel file has **3 tabs**:

### 1. Actuals Report (73 rows in example)
- BidItem, Activity, Resource, Quantity, Unit Price, etc.
- Ready for HCSS HeavyBid import
- All calculations automatic (AFUDC, Labor OH, Unit Prices)

### 2. Actual BoE (17 rows in example)
- Notes for Labor activities only
- Format: `ENGSVC15249: 104 MH Actuals to date, Projected an additional 0 MH for the remainder of the Activity`
- No manual editing needed - uses `0` instead of `___`

### 3. Resource File (33 rows in example)
- All unique resources from Actuals Report
- Resource definitions for HeavyBid

---

## WBS Operations Included

The embedded WBS dictionary includes all 74 standard operations:

| Level 2 | Level 3 | Operations | Example Activities |
|---------|---------|------------|-------------------|
| 01 PMO | 01, 02 | 1010-1190 | 0101-1010A, 0102-1100A |
| 02 Project Controls | 01-03 | 2010-2210 | 0201-2010A, 0202-2110A |
| 03 Project Setup | 01-03 | 3010-3210 | 0301-3020A, 0302-3100A |
| 04 Design | 01-03 | 4010-4220 | 0401-4030A, 0402-4110A |
| 05 Initiate | 01-08 | 5010-8800 | 0503-5050A, 0506-8300A |
| 06 Close | 01-02 | 9010-9130 | 0601-9010A, 0602-9110A |

---

## What's Embedded

All this data is now built into `wbs_operations_mapper.py`:
- ✅ 74 Operation codes
- ✅ Activity code mappings (e.g., 1010 → 0101-1010A)
- ✅ Level 2 and Level 3 WBS hierarchy
- ✅ Helper functions for validation

No external lookups, no file dependencies!

---

## For Project Templates

### Setup Template Folder
```
C:\Project_Templates\New_Project\
├── Documents/
├── Drawings/
├── Estimates/
│   ├── sap_to_heavybid.py          ← Copy here
│   └── wbs_operations_mapper.py    ← Copy here
└── Reports/
```

### Create New Project
1. Copy entire template folder
2. Rename to your job number
3. Scripts are ready to use immediately
4. No configuration needed

### Example Workflow
```bash
# Copy template
cp -r "C:\Project_Templates\New_Project" "C:\Projects\Job_74052500"

# Navigate to new project
cd "C:\Projects\Job_74052500\Estimates"

# Run conversion (scripts already there!)
# Just run - file and folder pickers guide you through it!
python sap_to_heavybid.py
```

---

## Updating WBS Operations

If PG&E updates the WBS Dictionary:

### Option 1: Update the Python File
1. Edit `wbs_operations_mapper.py`
2. Update the `OPERATIONS_MAP` dictionary
3. Save and you're done

### Option 2: Regenerate from New Excel
If you get a new WBS Dictionary Excel file:
1. Use the extraction script (contact support)
2. Replace `wbs_operations_mapper.py`
3. Distribute to all projects

**Note:** Updates are rare - the WBS structure is fairly stable.

---

## Troubleshooting

### "No module named 'wbs_operations_mapper'"
- Make sure both .py files are in the same folder
- Check file names are exact: `wbs_operations_mapper.py`

### "No module named 'pandas'"
- Run: `pip install pandas openpyxl`

### Wrong Activity Codes
- Verify you're using the latest `wbs_operations_mapper.py`
- Check that operation codes in SAP export are valid

### File Not Found
- Use the file picker to navigate to your SAP export file
- Use the folder picker to select where to save the output
- No need to type paths - the pickers handle navigation for you

---

## Version History

### v2.0 (December 2024)
- ✨ Standalone operation - no Excel file needed
- ✨ All WBS data embedded in Python
- ✨ Perfect for project templates
- ✅ All v1.0 features retained

### v1.0 (December 2024)
- Initial release
- Required external Excel file for WBS dictionary

---

## Benefits of Standalone Version

### Stability
- ❌ No "Excel file not found" errors
- ❌ No "Wrong version" issues
- ❌ No path dependencies
- ✅ Just works everywhere

### Simplicity
- 2 files instead of 3
- No file location configuration
- Copy and go

### Performance
- Faster loading (no Excel parsing)
- Instant startup
- Less memory usage

### Portability
- Email 2 files to anyone
- Copy to USB drive
- Works on any computer with Python
- Perfect for network shares

---

## Requirements

- **Python**: 3.6 or newer
- **Libraries**: pandas, openpyxl
- **GUI Library**: tkinter (usually included with Python)
- **Operating System**: Windows, Mac, or Linux
- **Files**: Just the 2 Python scripts

That's all you need!

---

## Getting Help

The scripts are designed to be bulletproof and work with any project. If you encounter issues:

1. Verify both .py files are in the same folder
2. Check Python and libraries are installed
3. Ensure your SAP export has standard columns (Order, Operation, Cost Element, etc.)
4. Try running from the command line to see error messages

Most issues are solved by ensuring both Python files are together in the same directory.

---

## Summary

**Old Way (v1.0):**
- 3 files required
- Excel file dependency
- Path configuration needed
- Must type file names every time

**New Way (v2.0):**
- 2 files only
- Fully standalone
- Copy and run
- **Just run `python sap_to_heavybid.py` - guided file and folder pickers!**
- Auto-named output files
- Choose exactly where to save your output

Perfect for templates, portable, and reliable! 🎉
