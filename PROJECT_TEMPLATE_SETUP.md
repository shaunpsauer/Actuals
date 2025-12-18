# Project Template Setup Guide

## How to Add SAP Converter to Your Project Template

This guide shows you how to integrate the SAP to HeavyBid converter into your project template folder so it's automatically available in every new project.

---

## Step 1: Locate Your Project Template Folder

Find where you store your project template. Common locations:
- `C:\Project_Templates\New_Project\`
- `C:\Users\YourName\Documents\Project_Templates\`
- Company network share: `\\server\templates\Project_Template\`

---

## Step 2: Download the Two Required Files

From the outputs folder, download:
1. ✅ `sap_to_heavybid.py`
2. ✅ `wbs_operations_mapper.py`

**That's all you need!** No Excel files, no configuration files.

---

## Step 3: Choose Where to Put Them in Your Template

Recommended locations in your template:

### Option A: In an "Estimates" or "Cost" Subfolder
```
Project_Template/
├── Documents/
├── Drawings/
├── Estimates/                    ← Create this folder
│   ├── sap_to_heavybid.py       ← Put files here
│   └── wbs_operations_mapper.py
└── Reports/
```

### Option B: In a "Scripts" or "Tools" Folder
```
Project_Template/
├── Documents/
├── Drawings/
├── Scripts/                      ← Create this folder
│   ├── sap_to_heavybid.py       ← Put files here
│   └── wbs_operations_mapper.py
└── Reports/
```

### Option C: In the Root Template Folder
```
Project_Template/
├── Documents/
├── Drawings/
├── sap_to_heavybid.py           ← Put files here
├── wbs_operations_mapper.py
└── Reports/
```

**Choose what works best for your organization!**

---

## Step 4: Copy the Files

1. Navigate to your template folder
2. Create the subfolder if needed (e.g., "Estimates")
3. Copy both .py files into that location
4. Done!

---

## Step 5: Test Your Template

### Create a Test Project
1. Copy your template folder to create a new project
2. Name it something like "TEST_74000000"
3. Navigate to where you put the scripts

### Verify Files Are There
Check that both files copied:
```bash
dir Estimates\*.py     # Windows
ls Estimates/*.py      # Mac/Linux
```

You should see:
```
sap_to_heavybid.py
wbs_operations_mapper.py
```

### Run a Quick Test (Optional)
If you have Python installed:
```bash
cd Estimates
python sap_to_heavybid.py
```

You should see usage instructions (since you didn't provide an input file).

---

## Using It in New Projects

### When You Create a New Project
1. Copy template folder → Name it with Job Number
2. Scripts are automatically in the Estimates folder
3. Ready to use immediately!

### To Convert SAP Export
1. Put your SAP export Excel file in the project folder
2. Open Command Prompt
3. Navigate to the folder with the scripts:
   ```bash
   cd C:\Projects\Job_74052500\Estimates
   ```
4. Run the converter:
   ```bash
   python sap_to_heavybid.py ..\SAP_Export.xlsx output.xlsx
   ```

---

## Example: Complete Workflow

### Template Setup (Do Once)
```
1. Navigate: C:\Project_Templates\New_Project\
2. Create: Estimates\ folder
3. Copy: sap_to_heavybid.py → Estimates\
4. Copy: wbs_operations_mapper.py → Estimates\
```

### New Project (Every Time)
```
1. Copy: C:\Project_Templates\New_Project → C:\Projects\Job_74052500
2. Add: SAP export file to Job_74052500 folder
3. Run: cd C:\Projects\Job_74052500\Estimates
        python sap_to_heavybid.py ..\EXPORT.xlsx actuals.xlsx
4. Use: actuals.xlsx in HCSS HeavyBid
```

---

## Tips for Success

### Keep Both Files Together
The two .py files MUST be in the same folder. If you move one, move both!

### Don't Rename the Files
Keep the exact names:
- `sap_to_heavybid.py` (not "converter.py" or "SAP_script.py")
- `wbs_operations_mapper.py` (not "mapper.py" or "operations.py")

### Update All Templates
If you have multiple project templates (small jobs, large jobs, etc.), add the scripts to all of them.

### Share with Team
If multiple people use the templates, make sure they:
1. Have Python installed
2. Have pandas and openpyxl libraries installed
3. Know how to run the scripts

---

## Updating the Scripts

### When Scripts are Updated
1. Download new versions
2. Replace old files in your template
3. New projects will automatically get the updated version
4. Existing projects keep their old version (still work fine)

### Updating Existing Projects
If you want to update scripts in an existing project:
1. Navigate to the project's Estimates folder
2. Replace the two .py files
3. Done!

---

## Network Share Setup

### For Shared Templates
If your template is on a network share:

```
\\company-server\templates\Project_Template\
├── Estimates\
│   ├── sap_to_heavybid.py       ← Everyone has access
│   └── wbs_operations_mapper.py
```

### When Creating New Projects
Users copy from network share to their local drive:
```bash
xcopy "\\company-server\templates\Project_Template" "C:\Projects\Job_74052500" /E /I
```

Scripts copy automatically with everything else!

---

## Troubleshooting Template Setup

### "Files not found in new project"
- Check files are actually in the template folder
- Verify you copied the entire folder structure
- Make sure copy operation completed

### "Scripts in template but not working"
- This is a Python installation issue, not a template issue
- Install Python and libraries on the computer
- Scripts themselves are fine

### "Want to put scripts elsewhere"
That's fine! Put them anywhere that makes sense for your organization:
- Root project folder
- Estimates subfolder
- Tools subfolder
- Cost folder

Just keep both .py files together!

---

## Benefits

### Template Integration Benefits
✅ Scripts ready in every new project
✅ No manual copying each time
✅ Consistent location across all projects
✅ Easy to find and use
✅ Automatic version control (update template, all new projects get it)

### No Configuration Needed
✅ No paths to set up
✅ No file locations to configure
✅ No company-specific customization
✅ Works the same for everyone

---

## Summary

1. **Setup** (once): Copy 2 files into your project template
2. **Use** (every project): Scripts are automatically there
3. **Run** (when needed): Convert SAP exports to HeavyBid format

Simple, portable, reliable! 🎉

---

## Quick Reference Card

```
┌─────────────────────────────────────────────┐
│  SAP Converter Template Setup               │
├─────────────────────────────────────────────┤
│  Files Needed: 2                            │
│  - sap_to_heavybid.py                      │
│  - wbs_operations_mapper.py                │
│                                             │
│  Template Location:                         │
│  ProjectTemplate\Estimates\*.py             │
│                                             │
│  New Project:                               │
│  1. Copy template                           │
│  2. Rename to job number                    │
│  3. Scripts ready in Estimates folder       │
│                                             │
│  Usage:                                     │
│  python sap_to_heavybid.py input.xlsx out.xlsx │
└─────────────────────────────────────────────┘
```

Keep this guide handy for training new team members!
