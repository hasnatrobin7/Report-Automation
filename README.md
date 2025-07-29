# Daily TLA Report Automation System

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A comprehensive automated reporting system that processes test data from Excel files, generates detailed TLA (Test Line Analyzer) reports with trend analysis, and automatically distributes results via email with embedded charts and images.

## 🚀 Features

### 📊 Report Generation
- **Automated Excel Reports**: Generates detailed reports with test results, failure rates, and serial number tracking
- **Trend Analysis**: Creates daily and weekly fail rate trends for top 3 syndroms
- **Visual Integration**: Embeds golden/defect images from syndrome database
- **Smart Formatting**: Automatically merges consecutive cells for cleaner presentation
- **Multi-shift Analysis**: Separates 1st Shift (00:00-15:30) and 2nd Shift (15:30-23:59) data

### 📧 Email Automation
- **HTML Email Reports**: Sends professionally formatted emails with embedded content
- **Summary Tables**: Pivot tables showing shift-specific failure rates
- **Image Embedding**: Inline golden and defect images (400x400 pixels)
- **Chart Integration**: Daily and weekly trend charts embedded as images
- **Recipient Management**: Configurable email distribution lists

### 🗄️ Data Management
- **Multi-file Processing**: Handles multiple Excel files across date ranges
- **Performance Optimization**: Uses DuckDB and Parquet caching for fast data processing
- **Syndrome Database**: Visual defect database with images and descriptions
- **Exclusion Lists**: Configurable syndrome exclusion for focused reporting
- **Date Range Selection**: Interactive date selection with multiple options

### 🔧 Management Tools
- **Syndrome Database UI**: Tkinter-based GUI for managing defect images and descriptions
- **File Analysis**: Excel file structure analyzer for debugging
- **Batch Execution**: Windows batch files for easy operation

## 🏗️ System Architecture

```
Daily TLA Report System
├── Data Processing Engine (generate_daily_report.py)
│   ├── Excel File Reader with Parquet Caching
│   ├── DuckDB Query Engine
│   ├── Shift Analysis Calculator
│   └── Trend Analysis Generator
├── Syndrome Database (SyndromDB/)
│   ├── Visual Defect Library
│   ├── Image Storage (golden.jpg, defect.jpg)
│   └── Description Repository
├── Database Management UI (syndrom_db_ui.py)
│   ├── Tkinter Interface
│   ├── Image Browser
│   └── Data Entry Forms
├── Email Engine
│   ├── HTML Template Generator
│   ├── Chart Image Creator
│   └── Outlook Integration
└── Configuration Management
    ├── Recipient Lists
    ├── Exclusion Rules
    └── Caching System
```

## 📋 Prerequisites

- **Python 3.7+**
- **Microsoft Outlook** (for email functionality)
- **Windows OS** (for COM automation)

## 🔧 Installation

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd Daily
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

   Or install manually:
   ```bash
   pip install pandas openpyxl matplotlib pywin32 pillow numpy pyarrow duckdb
   ```

3. **Setup configuration files**
   ```bash
   # Create email recipients list
   echo "# Email recipients" > recipients.txt
   echo "user@company.com" >> recipients.txt
   
   # Create syndrome exclusion list
   echo "# Syndromes to exclude" > exclude_syndroms.txt
   ```

## 📁 Project Structure

```
Daily/
├── 📄 generate_daily_report.py    # Main report generation engine
├── 📄 analyze_excel.py            # Excel file analyzer
├── 📄 syndrom_db_ui.py           # Syndrome database management UI
├── 📄 runme.bat                  # Main execution script
├── 📄 DB Updater.bat             # Database management launcher
├── 📄 requirements.txt           # Python dependencies
├── 📄 recipients.txt             # Email recipient list
├── 📄 exclude_syndroms.txt       # Syndrome exclusion list
├── 📄 Daily_TLA_Report.xlsx      # Generated report output
├── 📂 SyndromDB/                 # Syndrome database
│   ├── 📂 Syndrome Name 1/
│   │   ├── 🖼️ golden.jpg
│   │   ├── 🖼️ defect.jpg
│   │   └── 📄 description.txt
│   └── 📂 Syndrome Name 2/
│       ├── 🖼️ golden.jpg
│       ├── 🖼️ defect.jpg
│       └── 📄 description.txt
├── 📂 _parquet_cache/            # Performance cache (auto-generated)
└── 📂 Excel Data Files/          # Test data files (*.xlsx)
```

## 🚀 Quick Start

### Method 1: Using Batch Files (Recommended)
```bash
# Generate daily report
runme.bat

# Manage syndrome database
"DB Updater.bat"
```

### Method 2: Command Line
```bash
# Generate report
python generate_daily_report.py

# Manage syndrome database
python syndrom_db_ui.py

# Analyze Excel files
python analyze_excel.py
```

## ⚙️ Configuration

### 1. Email Recipients (`recipients.txt`)
```
# Daily TLA Report Recipients
# Lines starting with # are comments
john.doe@company.com
jane.smith@company.com
manager@company.com
```

### 2. Syndrome Exclusions (`exclude_syndroms.txt`)
```
# Syndromes to exclude from reports
# Case-sensitive, must match exactly
Alignment Verification - Horizontaly not aligned
Count Verification - DQM-Obstruction
```

### 3. Syndrome Database Setup
Create folders in `SyndromDB/` for each syndrome:
```
SyndromDB/
├── Count Verification - Missing Parts/
│   ├── golden.jpg      # Reference image
│   ├── defect.jpg      # Defect image
│   └── description.txt # Text description
└── Connector Not Flush/
    ├── golden.jpg
    ├── defect.jpg
    └── description.txt
```

## 🔄 Workflow

1. **Data Collection**: Place Excel test data files in the project directory
2. **Report Generation**: Run `runme.bat` or `python generate_daily_report.py`
3. **Date Selection**: Choose from available date ranges interactively
4. **Processing**: System analyzes data, calculates failure rates, generates trends
5. **Excel Output**: Creates `Daily_TLA_Report.xlsx` with embedded images and charts
6. **Email Distribution**: Automatically sends HTML email with summary to recipients
7. **Cleanup**: System manages cache files and temporary images

## 📊 Data Requirements

### Excel File Format
Your test data files must contain these columns:
- `StartDateTime` - Test timestamp (for shift calculation)
- `SerialNumber` - Unique identifier for each test
- `Syndrom` - Failure type/syndrome name
- `SyndromStatus` - Pass/Fail status
- `UUT` - Unit Under Test identifier

### Shift Definitions
- **1st Shift**: 00:00:00 to 15:29:59
- **2nd Shift**: 15:30:00 to 23:59:59

## 🎛️ Advanced Configuration

### Custom Shift Times
Edit constants in `generate_daily_report.py`:
```python
SHIFT_1_START = time(0, 0)
SHIFT_1_END = time(15, 30)
SHIFT_2_START = time(15, 30)
SHIFT_2_END = time(23, 59, 59)
```

### Image Dimensions
Modify image size for emails:
```python
IMG_WIDTH = 400   # Email image width
IMG_HEIGHT = 400  # Email image height
```

### Performance Tuning
- **Parquet Caching**: Automatically caches Excel data as Parquet files
- **DuckDB Integration**: Uses columnar database for fast queries
- **Parallel Processing**: Multi-threaded Excel file processing

## 🔍 Troubleshooting

### Common Issues

**📧 Email Not Sending**
- Verify Outlook is installed and configured
- Check `recipients.txt` format and email addresses
- Ensure Windows firewall allows Outlook automation

**🖼️ Images Not Loading**
- Verify SyndromDB folder structure
- Check image file formats (JPG recommended)
- Ensure syndrome names match folder names

**📊 No Data Found**
- Verify Excel files contain required columns
- Check date ranges in data files
- Ensure StartDateTime is properly formatted

**⚡ Performance Issues**
- Clear `_parquet_cache/` folder to rebuild cache
- Reduce date ranges for large datasets
- Check available disk space

### Debug Mode
Use the analyzer tool for troubleshooting:
```bash
python analyze_excel.py
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with Python and modern data processing libraries
- Uses DuckDB for high-performance analytics
- Integrates with Microsoft Outlook for seamless email automation
- Designed for manufacturing test line environments

---

**Need Help?** Create an issue or check the troubleshooting section above.