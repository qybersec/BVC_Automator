# BVC Automator - Deployment Guide

## Quick Setup for Boss/Other Users

This guide explains how to set up and run the BVC Automator on any Windows computer.

### Prerequisites

1. **Python 3.7 or higher** - Download from [python.org](https://python.org)
   - During installation, **check "Add Python to PATH"**

### Installation Steps

#### Option 1: Easy Setup (Recommended)

1. **Download the project** - Extract the ZIP file to any folder (e.g., `C:\BVC_Automator\`)
2. **Double-click** `setup_and_run.bat` - This will automatically install dependencies and run the app

#### Option 2: Manual Setup

1. **Download the project** - Extract the ZIP file to any folder (e.g., `C:\BVC_Automator\`)
2. **Install dependencies** - Open Command Prompt in the project folder and run:

   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**:

   ```bash
   python main.py
   ```

   Or double-click `run_bvc_automator.bat`

### Running After Setup

After the first setup, you can simply double-click `run_bvc_automator.bat` to start the application.

### What You Get

- **Basic Reports**: Automated TMS data processing for standard reports
- **Detailed Reports**: Advanced processing for detailed 27-column reports
- **BVC Template**: Generate date-range templates for reporting
- **Automatic Business Logic**:
  - Same carrier optimization
  - Negative savings correction
  - TL carrier special handling (LANDSTAR/SMARTWAY)
  - DDI carrier matching

### Core Files for Adaptation

If adapting this tool for other systems, focus on these key files:

#### Processing Logic

- `tms_processor.py` - Main processor for basic reports (lines 464-700 contain business rules)
- `tms_detailed_processor.py` - Processor for detailed reports (lines 493-700 contain business rules)

#### Business Rules Location

- **Basic Reports**: `_apply_business_logic_enhanced()` method in `tms_processor.py`
- **Detailed Reports**: `_apply_business_logic_detailed()` method in `tms_detailed_processor.py`
- **TL Carrier Logic**: See `TL_CARRIER_LOGIC.md` for documentation

#### Key Methods

- Data loading and structure detection
- Column mapping and cleaning
- Business rule application
- Excel output formatting
- Statistics calculation

### Configuration

- **TL Carriers**: Modify `TL_CARRIERS` set in both processor files to add/remove carriers
- **Column Mapping**: Adjust column names in the respective processor files
- **Business Rules**: Add new rules in the `_apply_business_logic_*()` methods

### File Structure

```txt
BVC_Automator/
├── main.py                    # Application entry point
├── tms_processor.py           # Basic report processor + GUI
├── tms_detailed_processor.py  # Detailed report processor
├── TL_CARRIER_LOGIC.md        # Business logic documentation
└── DEPLOYMENT_GUIDE.md        # This file
```

### Troubleshooting

#### Error: "No module named 'pandas'"

- Solution: Run `pip install pandas openpyxl numpy tkcalendar`

#### Error: "Python not found"

- Solution: Reinstall Python with "Add to PATH" checked

#### Calendar not showing

- Solution: Run `pip install tkcalendar`

### Support

For technical issues or modifications, refer to the codebase comments and documentation within the Python files.
