# 🚛 TMS Data Processor Pro

A professional, enterprise-grade Transportation Management System (TMS) data processor that transforms raw shipping data into beautifully formatted Excel reports with automatic cost analysis and savings calculations. Built for logistics professionals who need to analyze carrier performance and identify cost optimization opportunities.

## ✨ Features

### 🎯 **Core Functionality**

- **Automated Data Processing**: Handles various TMS file formats automatically
- **Smart Business Logic**: Applies industry-standard TMS rules and validations
- **Cost Analysis**: Calculates potential savings between selected and least-cost carriers
- **Professional Formatting**: Creates Excel reports with color-coded sections and insights

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

### For End Users (Non-Technical)

1. Download `TMS_Processor_Complete.zip`
2. Extract to a folder
3. Double-click `Run TMS Processor.bat`
4. Select your Excel file and click "Process File"

### For Developers

```bash
git clone https://github.com/yourusername/BVC_Automator.git
cd BVC_Automator
pip install -r requirements.txt
python tms_processor.py
```

## 📋 Requirements

- **Python 3.8+**
- **Windows 10/11** (primary target)
- **Excel files** (.xlsx format)

### Dependencies

```txt
pandas>=1.3.0
openpyxl>=3.0.0
tkinter (usually included with Python)
```

## 🔧 Installation

### Automatic Installation (Recommended)

1. Download the complete package
2. Run `Install_Requirements.bat` if needed
3. Use `Run TMS Processor.bat` to start

### Manual Installation

```bash
pip install pandas openpyxl
```

## 📁 Project Structure

```txt
BVC_Automator/
├── tms_processor.py          # Main processor logic
├── run_tms_processor.py      # GUI launcher
├── Run TMS Processor.bat     # Windows batch launcher
├── Install_Requirements.bat  # Requirements installer
├── requirements.txt          # Python dependencies
├── INSTALLATION_GUIDE.md    # User-friendly setup guide
├── README.md                # This file
├── test_processor.py        # Test suite
└── test_improvements.py     # Additional tests
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

Run the test suite to verify functionality:

```bash
python test_processor.py
python test_improvements.py
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

- **For Users**: Check `INSTALLATION_GUIDE.md` first
- **For Developers**: Review the code and test files
- **Issues**: Use GitHub Issues for bug reports and feature requests

## 🙏 Acknowledgments

- Built for transportation and logistics optimization
- Designed with user experience in mind
- Optimized for professional business use

---
