# 🚛 TMS Data Processor Pro

A professional, enterprise-grade Transportation Management System (TMS) data processor that transforms raw shipping data into beautifully formatted Excel reports with automatic cost analysis and savings calculations. Built for logistics professionals who need to analyze carrier performance and identify cost optimization opportunities.

## ✨ Features

### 🎯 **Core Functionality**

- **5 Processor Types**: Basic, UTC Main, UTC FS, Transco, Cast Nylon (Detailed)
- **Smart Business Logic**: 5+ automated TMS business rules with city-specific exclusions
- **Cost Analysis**: Calculates potential savings between selected and least-cost carriers
- **Professional Formatting**: Consistent color-coded Excel reports across all processors

### 🤖 **Automation Ready**

- **CLI Tool**: `automation_cli.py` - Full command-line interface for batch processing
- **Python API**: `automation_api.py` - Programmatic integration for enterprise systems
- **Auto-Detection**: Intelligent processor type selection based on file names
- **Batch Processing**: Handle multiple files with unified reporting

### 🎨 **Visual Enhancements**

- **Color-Coded Sections**: Selected Carrier (Blue), Least Cost Carrier (Orange), Potential Savings (Green)
- **Performance Insights**: Compact summary block with key metrics
- **Auto-sizing Columns**: Intelligent width adjustment based on content
- **Dynamic Row Heights**: Automatically adjusts row heights for long content (like carrier names)
- **Centered Content**: All cells are professionally centered for consistency

### 📊 **Output Features**

- **Summary Statistics**: Total loads, costs, savings, and optimization opportunities
- **Detailed Data Sheet**: Processed data with enhanced formatting and validation
- **Professional Headers**: Full descriptive column names (no abbreviations)
- **Auto-filtering**: Built-in Excel filters for easy data exploration

## 🚀 Quick Start

### For End Users (GUI)

1. Run `python main.py` or use `run_bvc_automator.bat`
2. Select processor type (Basic, UTC Main, UTC FS, Transco, Cast Nylon)
3. Upload Excel files and view real-time processing results
4. Download professionally formatted reports

### For Automation (CLI)

```bash
# Process single file with auto-detection
python automation_cli.py data.xlsx

# Process multiple files with specific type
python automation_cli.py *.xlsx --type utc_main --output-dir ./processed

# Export results to JSON
python automation_cli.py files/*.xlsx --export results.json
```

### For Integration (Python API)

```python
from automation_api import TMSAutomator

automator = TMSAutomator()
result = automator.process_file('data.xlsx', 'utc_main')
print(f"Savings: ${result['stats']['total_potential_savings']}")
```

## 📋 Requirements

- **Python 3.8+**
- **Windows 10/11** (primary target)
- **Excel files** (.xlsx format)

### Dependencies

```txt
pandas
openpyxl
numpy
tkcalendar
```

## 🔧 Installation

### Automatic Installation (Recommended)

1. Download the complete package
2. Run `setup_and_run.bat` - installs dependencies and launches application

### Manual Installation

```bash
pip install -r requirements.txt
python main.py
```

## 📁 Project Structure

```txt
BVC_Automator/
├── main.py                       # Main entry point
├── tms_processor.py              # Main GUI processor
├── automation_cli.py             # CLI interface for automation
├── automation_api.py             # Python API for integration
├── processor_interface.py        # Processor factory and interface
├── basic_processor.py            # Basic TMS processor
├── city_processors.py            # City-specific processors
├── tms_detailed_processor.py    # Cast Nylon (Detailed) processor
├── tms_utils.py                  # Utility functions
├── validators.py                 # Data validation utilities
├── config.py                     # Configuration management
├── logger_config.py              # Logging configuration
├── requirements.txt              # Python dependencies
├── run_bvc_automator.bat         # Windows launcher
├── setup_and_run.bat             # Setup and launch script
├── tms_config.json               # TMS configuration
├── README.md                     # This file
├── DEPLOYMENT_GUIDE.md           # Deployment instructions
└── TL_CARRIER_LOGIC.md           # TL carrier business logic
```

## 🎯 How It Works

### 1. **Data Detection**

- Automatically detects TMS file structure
- Identifies headers, data rows, and company information
- Handles various file formats intelligently

### 2. **Business Logic Application**

- **Same Carrier Rule**: Sets savings to 0 when selected = least cost
- **Empty Data Handling**: Copies selected data to least cost when missing
- **Negative Savings**: Corrects negative values to 0
- **DDI/Carrier Matching**: When Selected Carrier contains "Company/Carrier Name" and the carrier after "/" matches Least Cost Carrier, copies selected data and zeros out savings

### 3. **Formatting & Output**

- Creates professional Excel workbook with multiple sheets
- Applies color coding and styling automatically
- Generates summary statistics and insights

## 📊 Input Format

The processor expects Excel files with columns like:

- Load information (Load No., Ship Date, Origin/Destination)
- Selected carrier details (Carrier, Service Type, Costs)
- Least cost alternatives (Carrier, Service Type, Costs)

## 🎨 Output Format

### Main Data Sheet

- **Columns A-H**: Load and location details
- **Columns I-N**: Selected carrier information (Blue theme)
- **Columns O-T**: Least cost carrier information (Orange theme)
- **Column U**: Potential savings (Green theme)

### Summary Sheet

- Comprehensive statistics and insights
- Company and date information
- Key performance metrics

## 🧪 Testing

The project includes comprehensive unit tests and validation. All core processors have been tested with real-world data to ensure accuracy and reliability.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

- **For Users**: Check `DEPLOYMENT_GUIDE.md` for setup instructions
- **For Developers**: Review the code and processor interface documentation
- **Issues**: Use GitHub Issues for bug reports and feature requests

## 🙏 Acknowledgments

- Built for transportation and logistics optimization
- Designed with user experience in mind
- Optimized for professional business use

---
