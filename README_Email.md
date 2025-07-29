# Daily TLA Report Generator - Complete Usage Guide

## Overview
The Daily TLA Report Generator is a comprehensive Python tool that automatically generates Excel reports and sends email summaries with embedded charts and images. It processes test data from Excel files, creates detailed reports with trend analysis, and automatically emails results to specified recipients.

## Features

### Excel Report Generation
- **Detailed Test Reports**: Creates Excel files with test results, failure rates, and SN tracking
- **Trend Charts**: Daily and weekly fail rate trends for top 3 syndroms
- **Image Integration**: Embeds golden and defect images from SyndromDB
- **Cell Merging**: Merges consecutive cells for cleaner presentation
- **Special Character Handling**: Automatically matches syndrom names to sanitized folder names

### Email Functionality
- **Automatic Email Sending**: Sends HTML emails with embedded content
- **Summary Tables**: Pivoted tables showing 1st and 2nd shift rates separately
- **Embedded Images**: Golden and defect images embedded inline (400x400 pixels)
- **Trend Charts**: Daily and weekly charts embedded as images
- **Professional Formatting**: Clean HTML layout with proper styling

### Data Processing
- **Multi-file Support**: Processes multiple Excel files in date ranges
- **Shift Analysis**: Separates 1st Shift (00:00-15:30) and 2nd Shift (15:30-23:59)
- **Exclusion Lists**: Configurable syndrom exclusion via exclude_syndroms.txt
- **Rate Calculations**: Accurate fail rates per shift based on unique SNs

## Setup Instructions

### 1. Install Dependencies
```bash
pip install pandas openpyxl matplotlib pywin32 pillow
```

Or install from requirements.txt:
```bash
pip install -r requirements.txt
```

### 2. File Structure Setup
```
Daily/
├── generate_daily_report.py    # Main script
├── recipients.txt              # Email recipients (create this)
├── exclude_syndroms.txt       # Syndroms to exclude (create this)
├── SyndromDB/                 # Images and descriptions folder
│   ├── README.txt             # Folder structure instructions
│   ├── SyndromName1/          # One folder per syndrom
│   │   ├── golden.jpg         # Golden/reference image
│   │   ├── defect.jpg         # Defect/failure image
│   │   └── description.txt    # Syndrom description
│   └── SyndromName2/
│       ├── golden.jpg
│       ├── defect.jpg
│       └── description.txt
└── Excel Files/               # Your test data files
    ├── data_2024-01-01.xlsx
    ├── data_2024-01-02.xlsx
    └── ...
```

### 3. Configure Email Recipients
Create `recipients.txt` file:
```
# Email recipients for daily TLA report
# Add one email address per line
# Lines starting with # are comments and will be ignored

john.doe@company.com
jane.smith@company.com
manager@company.com
```

### 4. Configure Syndrom Exclusions
Create `exclude_syndroms.txt` file:
```
# List of syndroms to exclude from the daily report
# Add one syndrom per line, exactly as it appears in your data (case-sensitive)
# Lines starting with # are comments and will be ignored

Alignment Verification - Horizontaly not aligned
Alignment Verification - Vertical not aligned
Count Verification - DQM-Obstruction
```

### 5. Setup SyndromDB
For each unique syndrom (failure type), create a subfolder in `SyndromDB/`:

**Folder Naming Rules:**
- Use hyphens (-) instead of slashes (/)
- Avoid special characters: `\ : * ? " < > |`
- Examples:
  - `Count Verification - DQM-Obstruction/`
  - `Connector Not Flush-4-4-9-1-2-14/`

**Required Files in Each Folder:**
- `golden.jpg` - Golden/reference image for this syndrom
- `defect.jpg` - Defect/failure image for this syndrom  
- `description.txt` - Text description of the syndrom

## Usage Instructions

### Running the Script
```bash
python generate_daily_report.py
```

### Interactive Prompts

1. **Date Selection**: Choose from available Excel files and date ranges
   - Single date
   - Date range
   - Latest date
   - All available data

2. **Trend Charts**: Choose whether to include trend charts in email
   - `y` - Include daily and weekly trend charts
   - `n` - Skip trend charts

3. **Email Confirmation**: Confirm email sending
   - `y` - Send email to recipients
   - `n` - Skip email sending

### Output Files

**Excel Report**: `Daily_TLA_Report.xlsx`
- Main report sheet with test results
- Daily trend chart sheet
- Weekly trend chart sheet
- Merged cells for cleaner presentation
- Embedded images from SyndromDB

**Email Content**:
- HTML email with embedded summary table
- Golden and defect images (400x400 pixels)
- Daily and weekly trend charts (if selected)
- Professional formatting with proper styling

## Data Requirements

### Excel File Format
Your Excel files must contain these columns:
- `StartDateTime` - Test start time (used for shift calculation)
- `SerialNumber` - Unique serial number for each test
- `Syndrom` - Failure type/syndrom name
- `SyndromStatus` - Pass/Fail status
- `UUT` - Unit Under Test identifier

### Shift Time Boundaries
- **1st Shift**: 00:00:00 to 15:29:59
- **2nd Shift**: 15:30:00 to 23:59:59

## Rate Calculation

The script calculates fail rates correctly:
- **Per Shift**: `(fail count for shift) / (unique SNs in shift) × 100`
- **Per Syndrom**: Based on top 3 syndroms by total fail count
- **Per UUT**: Separate rates for each Unit Under Test

## Troubleshooting

### Common Issues

**Email Not Sending**
- Check `recipients.txt` exists and contains valid email addresses
- Ensure Outlook is installed and configured
- Verify you have permission to send emails
- Check Windows firewall settings

**Images Not Found**
- Verify SyndromDB folder structure is correct
- Check that syndrom names match folder names (special characters are handled automatically)
- Ensure golden.jpg and defect.jpg exist in each syndrom folder

**Charts Not Generating**
- Make sure you selected "y" when prompted for trend charts
- Check that matplotlib is properly installed
- Verify you have sufficient data for trend analysis

**Rate Calculations Wrong**
- Ensure SerialNumber column contains unique values
- Check that StartDateTime is in correct format
- Verify shift time boundaries match your data

**Missing Dependencies**
```bash
pip install pandas openpyxl matplotlib pywin32 pillow
```

### Error Messages

**"No Excel files found"**
- Ensure Excel files are in the same directory as the script
- Check file extensions are `.xlsx`

**"No data found for selected date range"**
- Verify Excel files contain data for the selected dates
- Check that StartDateTime column exists and is properly formatted

**"Error processing image"**
- Check image file formats (JPEG recommended)
- Verify image files are not corrupted
- Ensure sufficient disk space for image processing

## Advanced Configuration

### Customizing Shift Times
Edit the constants in `generate_daily_report.py`:
```python
SHIFT_1_START = time(0, 0)
SHIFT_1_END = time(15, 30)
SHIFT_2_START = time(15, 30)
SHIFT_2_END = time(23, 59, 59)
```

### Image Size Configuration
Modify image dimensions in the script:
```python
IMG_WIDTH = 400   # Email image width
IMG_HEIGHT = 400  # Email image height
```

### Email Template Customization
Edit the HTML template in `send_email_with_charts()` function to customize email appearance.

## File Maintenance

### Regular Tasks
1. **Update SyndromDB**: Add new syndroms as they are discovered
2. **Review Exclusions**: Update `exclude_syndroms.txt` as needed
3. **Clean Old Files**: Remove old Excel files to improve performance
4. **Update Recipients**: Modify `recipients.txt` for personnel changes

### Backup Recommendations
- Backup SyndromDB folder regularly
- Keep copies of important Excel data files
- Archive old reports for historical analysis

## Support

For issues or questions:
1. Check this README for troubleshooting steps
2. Verify all dependencies are installed correctly
3. Ensure file structure matches requirements
4. Test with sample data before using production data

## Version History

- **v1.0**: Initial release with Excel report generation
- **v1.1**: Added email functionality with embedded charts
- **v1.2**: Added image embedding and summary tables
- **v1.3**: Fixed shift rate calculations and special character handling
- **v1.4**: Enhanced image processing and email formatting 