# Quick Start Guide

Get up and running with the ADA Audit Tool in 5 minutes!

## Prerequisites

- Python 3.8 or higher installed
- Your PADC Excel file ready
- Basic familiarity with running Python applications

## Installation Steps

### 1. Download or Clone the Project

```bash
# If using Git
git clone https://github.com/yourusername/ada-audit-tool.git
cd ada-audit-tool

# Or download and extract the ZIP file, then navigate to the folder
```

### 2. Install Dependencies

```bash
# Windows
python -m pip install -r requirements.txt

# macOS/Linux
python3 -m pip install -r requirements.txt
```

**Note:** If you get a "tkinter not found" error:
- **Windows**: Reinstall Python with "tcl/tk and IDLE" option checked
- **Ubuntu/Debian**: `sudo apt-get install python3-tk`
- **macOS**: `brew install python-tk`

### 3. Run the Application

```bash
# Windows
python ADA_Audit_GUI.py

# macOS/Linux
python3 ADA_Audit_GUI.py
```

## First Use Workflow

### Step 1: Select Your Input File
1. Click **"Browse Input File"** (or press `Ctrl+O`)
2. Navigate to your PADC Excel file
3. Select the file and click Open

### Step 2: Set Output Location
1. Click **"Browse Output File"** (or press `Ctrl+S`)
2. Choose where you want to save the processed data
3. Name your output file and click Save

### Step 3: Load and Analyze Data
1. Click **"Load & Analyze Data"** (or press `Ctrl+L`)
2. Wait for the tool to analyze your file (progress bar will show status)
3. The program boundaries will be automatically detected and displayed in the table

### Step 4: Review Boundaries (Optional)
- The table shows detected program boundaries
- Verify the Start Row and End Row for each program
- Double-click any cell to edit if needed
- Programs not found in your file will show "Not Found"

### Step 5: Run the Audit
1. Click **"Run Audit Process"** (or press `Ctrl+R`)
2. Wait for processing to complete
3. Check the log area for progress and results
4. Your output file will be created at the specified location

### Step 6: Generate Dashboard (Optional)
1. Click **"Run ADA Dashboard"** (or press `Ctrl+D`)
2. Enter school year, location, and school name when prompted
3. A CSV dashboard file will be created in the same directory

## Saving Time with Configurations

If you process files for the same school regularly:

1. After loading and analyzing data, click **"Save Configuration"**
2. Give your configuration a meaningful name (e.g., "Lincoln Elementary")
3. Next time, click **"Load Configuration"** to instantly load these boundaries
4. No need to analyze the file again if the structure is the same!

## Keyboard Shortcuts

Master these shortcuts to work faster:

| Shortcut | Action |
|----------|--------|
| `Ctrl+O` | Open input file |
| `Ctrl+S` | Select output file |
| `Ctrl+L` | Load & analyze data |
| `Ctrl+R` | Run audit |
| `Ctrl+D` | Run dashboard |
| `F1` | Show help |

## Alternative: Run with Pre-configured Settings

If you have a saved configuration, you can skip the GUI entirely:

```bash
python run_with_config.py
```

This script uses pre-configured boundary settings from the `configs/` or `boundary_settings/` folders.

## Common Issues & Solutions

### "No data extracted"
- **Solution**: Check that your Excel file has the expected structure
- Verify program names match exactly (including spaces and capitalization)
- Review the boundary settings in the table

### Processing is slow
- **Normal**: Large files with 12 months of data can take 1-2 minutes
- Progress bar shows current status
- Don't close the window during processing

### Boundaries look wrong
- **Solution**: Manually edit the Start Row and End Row values
- Double-click the cell to edit
- Use Excel row numbers (the numbers on the left side of Excel)

### Data appearing in wrong row
- **Solution**: Verify the correct school type (TK-12 vs K-12) is selected
- Check that TK programs are correctly separated from main programs
- Review the log for consolidation details

### Git errors with Excel files
- **Solution**: Close Excel before running git commands
- Temp files (`~$*.xlsx`) are automatically ignored

### Application won't start
- Check Python version: `python --version` (need 3.8+)
- Verify tkinter is installed: `python -c "import tkinter"`
- Reinstall dependencies: `pip install -r requirements.txt --force-reinstall`

## Next Steps

- **Save configurations** for schools you process regularly
- **Explore the dashboard** feature for CSV summaries
- **Try run_with_config.py** for automated processing
- **Read the full README.md** for advanced features
- **Check CHANGELOG.md** for recent updates and fixes
- **Check CONTRIBUTING.md** if you want to improve the tool

## Getting Help

- Press `F1` in the application for accessibility help
- Check the log area for error messages
- Review README.md for detailed documentation
- Review CHANGELOG.md for known issues and fixes
- Create an issue on GitHub if you find a bug

## Tips for Success

✅ **Always verify** the boundary detection before running the audit  
✅ **Save configurations** for schools you process regularly  
✅ **Use descriptive names** when saving output files  
✅ **Keep backups** of your original Excel files  
✅ **Check the logs** if something seems wrong  
✅ **Close Excel files** before running git commands  

---

**That's it!** You're ready to process attendance data efficiently. Happy auditing! 🎉

