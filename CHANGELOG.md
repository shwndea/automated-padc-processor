# Changelog

All notable changes to the ADA Audit Tool will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.1.0] - 2026-02-18

### Fixed
- **Critical: TK vs K-3 Data Mapping Issue**
  - Fixed Program C row 58 to correctly fetch K-3 data instead of TK-3
  - Fixed Program N row 75 to correctly fetch K-3 data instead of TK-3
  - Fixed Program J row 65 to correctly fetch `Prog_J_TK` (TK program) data instead of main program TK-3 age group
  - Fixed Program J row 66 to correctly fetch main program K-3 data
  - Fixed Program K row 83 to correctly fetch K-3 data instead of TK-3
  
- **Consolidation Logic Overhaul**
  - TK programs (`_TK` suffix) now only create TK-3 age group data
  - Main programs now correctly create K-3, 4-6, 7-8, 9-12 age group data (not TK-3)
  - This ensures proper separation between TK programs and main programs with TK-3 age groups

### Added
- **New Utility Scripts**
  - `run_with_config.py`: Run audits with pre-configured boundary settings
  - `check_available_months.py`: Utility to check available months in Excel files
  - `print_ada_consolidation.py`: Debug utility for viewing consolidation results
  - `print_ada_consolidation_FIXED.py`: Fixed version of consolidation printer
  - `test_display_values.py`: Test utility for verifying display values

- **New Configuration Files**
  - `configs/intake_locations.json`: Pre-configured intake location boundaries
  - `boundary_settings/COA Elem.json`: COA Elementary boundary configuration
  - `boundary_settings/COA Mid.json`: COA Middle School boundary configuration
  - `boundary_settings/HLA.json`: HLA boundary configuration

- **Helper Function for Attendance Values**
  - Added `get_attendance_value()` helper function to handle TK-3/K-3 notation fallback
  - Improves reliability when source Excel files use inconsistent notation

### Changed
- **Cell Mapping Structure**
  - Reorganized cell mappings to group TK rows before main program rows
  - Added clear comments documenting row assignments for each program
  - Improved code readability with consistent ordering

- **Row Assignments (TK-12 Mode)**
  - Row 57: Prog_C_TK (TK-3 only)
  - Rows 58-61: Prog_C Main (K-3, 4-6, 7-8, 9-12)
  - Row 65: Prog_J_TK (TK-3 only)
  - Rows 66-69: Prog_J Main (K-3, 4-6, 7-8, 9-12)
  - Row 74: Prog_N_TK (TK-3 only)
  - Rows 75-78: Prog_N Main (K-3, 4-6, 7-8, 9-12)
  - Row 82: Prog_K_TK (TK-3 only)
  - Rows 83-86: Prog_K Main (K-3, 4-6, 7-8, 9-12)

### Documentation
- Updated .gitignore to exclude Excel temp files (`~$*`)

## [2.0.0] - 2025-11-06

### Added
- **ADA-Compliant GUI Interface**
  - WCAG 2.1 AA compliant design with high contrast colors
  - Comprehensive keyboard navigation with shortcuts
  - Screen reader compatible with proper ARIA labels
  - Visible focus indicators for all interactive elements
  - Automatic high contrast mode detection for Windows
  
- **Configuration Management**
  - Save program boundary configurations for reuse
  - Load pre-configured boundary settings
  - Support for multiple school configurations
  - JSON-based configuration storage
  
- **Dashboard Module (`ADA_Dashboard_Module.py`)**
  - Generate CSV dashboards with ADA summaries
  - Configurable school year, location, and school name
  - Summary by program, month, and grade level
  - Integration with main GUI application
  
- **Enhanced User Experience**
  - Real-time progress tracking with progress bar
  - Detailed logging with color-coded messages
  - Scrollable interface for all screen sizes
  - Table sorting by clicking column headers
  - Inline editing of boundary values
  
- **Accessibility Features**
  - Full keyboard navigation (Ctrl+O, Ctrl+L, Ctrl+R, etc.)
  - Scrolling shortcuts (Ctrl+Up/Down, Page Up/Down, Home/End)
  - Table sorting shortcuts (F2-F5)
  - Help dialog (F1) with comprehensive instructions
  - Status announcements for screen readers
  
- **Program Consolidation**
  - McClellan (CM) locations consolidated with main programs
  - Sacramento Youth Center (SYC) locations consolidated with main programs
  - TK programs kept separate as specified
  - Configurable consolidation rules

### Changed
- **Architecture**: Modularized codebase with separate GUI, processing, and dashboard modules
- **Data Processing**: More efficient Excel writing with batch operations
- **User Interface**: Complete redesign focusing on accessibility and usability
- **Configuration**: Moved from hardcoded values to JSON-based configurations

### Fixed
- Improved boundary detection accuracy
- Better error handling and user feedback
- More reliable Excel file processing
- Enhanced data validation

### Documentation
- Comprehensive README.md
- CONTRIBUTING.md with development guidelines
- QUICKSTART.md for new users
- Example configuration files
- Inline code documentation with detailed docstrings

## [1.0.0] - 2025-10-09

### Added
- Initial release with command-line interface
- Basic ADA audit processing functionality
- Support for Programs C, N, J, K with variants
- Excel file input/output
- Month-by-month attendance extraction
- Program boundary detection

### Features
- Automated program name detection
- Month number identification (1-12)
- Attendance data extraction (Column AJ)
- Excel output generation
- Progress tracking with tqdm

## Version Comparison

### v2.0.0 vs v1.0.0

**Major Improvements:**
- ✅ Full GUI interface (was command-line only)
- ✅ ADA/WCAG 2.1 AA accessibility compliance
- ✅ Configuration save/load functionality
- ✅ Dashboard CSV generation
- ✅ Real-time progress and logging
- ✅ Keyboard navigation and shortcuts
- ✅ Comprehensive documentation

**Maintained Features:**
- ✅ All program types support (C, N, J, K)
- ✅ TK, CM, SYC location variants
- ✅ Month-by-month processing (1-12)
- ✅ Automatic boundary detection
- ✅ Excel file processing

## Upcoming Features (Roadmap)

### Planned for v2.1.0
- [ ] Batch processing of multiple files
- [ ] Export to additional formats (PDF reports)
- [ ] Data validation and error checking
- [ ] Template management system
- [ ] Command-line interface option alongside GUI

### Planned for v2.2.0
- [ ] Integration with web-based dashboard
- [ ] Database backend option (PostgreSQL)
- [ ] User authentication and multi-user support
- [ ] Audit trail and logging
- [ ] Scheduled automated processing

### Under Consideration
- [ ] Cloud storage integration (Google Drive, OneDrive)
- [ ] Email notifications on completion
- [ ] Advanced reporting and analytics
- [ ] REST API for integration with other systems
- [ ] Mobile app for reviewing results

## Breaking Changes

### v2.0.0
- **Configuration format**: Changed to JSON (old configs incompatible)
- **Python version**: Minimum version increased to 3.8
- **Dependencies**: Added tkinter, prettytable requirements
- **File structure**: New modular architecture

---

## Release Notes Format

Each release includes:
- **Version number** following semantic versioning (MAJOR.MINOR.PATCH)
- **Release date** in ISO format (YYYY-MM-DD)
- **Added** for new features
- **Changed** for changes in existing functionality
- **Deprecated** for soon-to-be removed features
- **Removed** for now removed features
- **Fixed** for bug fixes
- **Security** for vulnerability fixes

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for information on how to contribute to this project.

## Support

For questions or issues, please:
- Check the [README.md](README.md) for documentation
- Review the [QUICKSTART.md](QUICKSTART.md) for common solutions
- Create an issue on GitHub with details

---

**Legend:**
- ✅ Completed
- 🚧 In Progress
- 📋 Planned
- ⚠️ Breaking Change

