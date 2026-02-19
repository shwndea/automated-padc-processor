# ADA Audit Tool - PADC Processor

A comprehensive, ADA-compliant GUI application for processing school attendance data and generating ADA (Average Daily Attendance) audit reports and dashboards.

## Overview

This tool automates the processing of Program Apportionment Data Collection (PADC) files for charter schools, specifically designed to handle multiple program types (C, N, J, K) with various location-specific variants (TK, CM, SYC). The application provides an accessible, keyboard-navigable interface that meets WCAG 2.1 AA accessibility standards.

## Features

### Core Functionality
- **Automated Data Processing**: Extracts student attendance data from Excel files based on configurable program boundaries
- **Multiple Program Support**: Handles 12+ program variations across resident/non-resident and independent study programs
- **Boundary Management**: Save and load program boundary configurations for different schools/locations
- **Dashboard Generation**: Creates CSV dashboards with ADA summaries by program, month, and grade level
- **Real-time Progress Tracking**: Visual progress indicators and detailed logging

### Accessibility Features
- **WCAG 2.1 AA Compliant**: High contrast color scheme with proper text-to-background ratios
- **Keyboard Navigation**: Full keyboard access with comprehensive shortcuts (Ctrl+O, Ctrl+L, Ctrl+R, etc.)
- **Screen Reader Compatible**: Proper ARIA labels and accessible widget configurations
- **Focus Indicators**: Clear visual focus indicators for all interactive elements
- **High Contrast Mode**: Automatic detection and support for Windows high contrast settings
- **Scrolling Support**: Multiple scrolling methods (keyboard, mouse, page navigation)

## System Requirements

- **Operating System**: Windows 10/11 (primary), macOS, or Linux
- **Python**: 3.8 or higher
- **Excel**: Input files must be in .xlsx format
- **Memory**: Minimum 4GB RAM recommended
- **Display**: Minimum 1200x800 resolution for optimal GUI experience

## Installation

### 1. Clone the Repository
```bash
git clone https://github.com/yourusername/ada-audit-tool.git
cd ada-audit-tool
```

### 2. Create Virtual Environment (Recommended)
```bash
python -m venv venv
# Windows
venv\Scripts\activate
# macOS/Linux
source venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

## Quick Start

### Running the GUI Application
```bash
python ADA_Audit_GUI.py
```

### Basic Workflow
1. **Select Input File** (Ctrl+O): Choose your PADC Excel file
2. **Select Output File** (Ctrl+S): Choose where to save processed data
3. **Load & Analyze** (Ctrl+L): Load data and auto-detect program boundaries
4. **Review Boundaries**: Edit boundaries in the table if needed
5. **Run Audit** (Ctrl+R): Execute the audit process
6. **Run Dashboard** (Ctrl+D): Generate CSV dashboard (optional)
7. **Export Results**: Save results to Excel

## Program Structure

### Program Types

The tool supports the following program configurations:

**Program C - Charter Resident**
- `Prog_C`: Main program (includes CM and SYC locations)
- `Prog_C_TK`: Transitional Kindergarten (separate)
- `Prog_C_CM`: McClellan location (consolidated with main)
- `Prog_C_SYC`: Sacramento Youth Center (consolidated with main)

**Program N - Non-Resident Charter**
- `Prog_N`: Main program (includes CM and SYC locations)
- `Prog_N_TK`: Transitional Kindergarten (separate)
- `Prog_N_CM`: McClellan location (consolidated with main)
- `Prog_N_SYC`: Sacramento Youth Center (consolidated with main)

**Program J - Independent Study Charter Resident**
- `Prog_J`: Main program
- `Prog_J_TK`: Transitional Kindergarten

**Program K - Independent Study Charter Non-Resident**
- `Prog_K`: Main program
- `Prog_K_TK`: Transitional Kindergarten

### File Structure

```
ada-audit-tool/
├── ADA_Audit_GUI.py              # Main GUI application
├── ADA_Audit_25_26_IMPROVED.py   # Core audit processing functions
├── ADA_Dashboard_Module.py       # Dashboard generation module
├── run_with_config.py            # Run audits with pre-configured settings
├── check_available_months.py     # Utility to check available months
├── boundary_settings/            # Saved boundary configurations
│   ├── COA Elem.json
│   ├── COA Mid.json
│   ├── HLA.json
│   └── example_configuration.json
├── configs/                      # Additional configuration files
│   └── intake_locations.json
├── sample_data/                  # Sample data files for testing
│   └── README.md
├── requirements.txt              # Python dependencies
├── README.md                     # This file
├── CHANGELOG.md                  # Version history and changes
├── QUICKSTART.md                 # Quick start guide
├── CONTRIBUTING.md               # Contribution guidelines
└── LICENSE                       # License information
```

## Configuration

### Boundary Settings

Boundary settings define where each program's data begins and ends in the Excel file. The tool can:
- Auto-detect boundaries by analyzing the input file
- Save configurations for reuse across similar files
- Load pre-configured boundary settings for specific schools

Example boundary configuration (JSON):
```json
{
  "name": "COA Elementary",
  "program_boundaries": {
    "Prog_C": {"start": 5, "stop": 28},
    "Prog_C_TK": {"start": 29, "stop": 48}
  }
}
```

### Running with Pre-configured Settings

Use the `run_with_config.py` script to run audits with saved boundary configurations:

```bash
python run_with_config.py
```

### Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+O | Open input file |
| Ctrl+S | Select output file |
| Ctrl+L | Load and analyze data |
| Ctrl+R | Run audit process |
| Ctrl+E | Export results |
| Ctrl+D | Run ADA Dashboard |
| F1 | Show help |
| F2-F5 | Sort table by different columns |
| Esc | Return focus to main window |
| Ctrl+Up/Down | Scroll line by line |
| Page Up/Down | Scroll page by page |

## Data Mapping (TK-12 Mode)

The tool correctly maps data to Excel rows based on program type:

| Row | Program | Age Group | Description |
|-----|---------|-----------|-------------|
| 57 | Prog_C_TK | TK-3 | Charter Resident TK Program |
| 58-61 | Prog_C | K-3, 4-6, 7-8, 9-12 | Charter Resident Main Program |
| 65 | Prog_J_TK | TK-3 | Independent Study Resident TK Program |
| 66-69 | Prog_J | K-3, 4-6, 7-8, 9-12 | Independent Study Resident Main Program |
| 74 | Prog_N_TK | TK-3 | Non-Resident Charter TK Program |
| 75-78 | Prog_N | K-3, 4-6, 7-8, 9-12 | Non-Resident Charter Main Program |
| 82 | Prog_K_TK | TK-3 | Independent Study Non-Resident TK Program |
| 83-86 | Prog_K | K-3, 4-6, 7-8, 9-12 | Independent Study Non-Resident Main Program |

## Output Files

### Audit Output
- Excel file with extracted attendance data
- Organized by program, month, and grade level
- Column AJ values (attendance figures)

### Dashboard Output
- CSV file with ADA summary data
- Columns: Year, School, Location, Month, Program, TK, Grade Level, ADA %, Total ADA
- Suitable for import into BI tools or further analysis

## Troubleshooting

### Common Issues

**"No module named 'tkinter'"**
```bash
# Ubuntu/Debian
sudo apt-get install python3-tk
# macOS (via Homebrew)
brew install python-tk
```

**"File not found" errors**
- Ensure input file path is correct
- Check that the file is not open in Excel
- Verify file permissions

**Excel temp file errors (`~$` files)**
- Close the Excel file before running git commands
- These files are automatically ignored by git

**Boundary detection fails**
- Manually review the Excel file structure
- Check that program names match expected values
- Adjust boundaries manually in the GUI table

**GUI scaling issues**
- Adjust Windows display scaling (Settings > Display)
- Minimum recommended resolution: 1200x800

**Data appearing in wrong rows**
- Verify the correct school type (TK-12 vs K-12) is selected
- Check that TK programs (`_TK` suffix) have correct boundaries
- Review the consolidation log for data mapping details

## Development

### Running from Source
```bash
python ADA_Audit_GUI.py
```

### Module Usage
You can also import and use the modules programmatically:

```python
from ADA_Audit_25_26_IMPROVED import (
    find_rows_containing_program_name,
    find_program_boundary_rows,
    extract_student_attendance_data
)
from ADA_Dashboard_Module import run_ada_dashboard_with_boundaries
```

### Utility Scripts

- `check_available_months.py` - Check which months have data in an Excel file
- `print_ada_consolidation.py` - Debug consolidation results
- `test_display_values.py` - Verify display values are correct

## Contributing

We welcome contributions! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

### Code Style
- Follow PEP 8 guidelines
- Use descriptive variable names
- Add docstrings to all functions
- Maintain accessibility standards

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

For issues, questions, or feature requests:
- Create an issue on GitHub
- Review the [CHANGELOG.md](CHANGELOG.md) for recent changes

## Acknowledgments

- Built for charter school attendance reporting compliance
- Designed with accessibility as a primary consideration
- Integrates with FastAPI, SQLModel, Celery, Redis backend systems

## Version History

### v2.1.0 (Current - February 2026)
- Fixed critical TK vs K-3 data mapping issues
- Added utility scripts for debugging and configuration
- Improved consolidation logic for TK/main program separation
- Added pre-configured boundary settings for multiple schools

### v2.0.0 (November 2025)
- ADA-compliant GUI interface
- Saved boundary configurations
- Dashboard generation module
- Comprehensive keyboard navigation
- High contrast mode support

### v1.0.0 (October 2025)
- Initial command-line version
- Basic audit processing functionality

---

**Note**: This tool is specifically designed for PADC (Program Apportionment Data Collection) processing for California charter schools. Ensure your input files match the expected format for proper operation.

