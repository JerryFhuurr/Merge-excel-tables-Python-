# Excel Tables Merger

A Python script that automatically merges multiple Excel files into a single consolidated table while preserving formatting and handling password-protected files.

## Features

- 🔄 **Automatic Detection**: Auto-detects all Excel files in the current directory
- 🔐 **Password Protection**: Handles password-protected Excel files automatically
- 🎨 **Formatting Preservation**: Maintains original cell formatting (fonts, colors, borders, alignment)
- 📊 **Smart Header Detection**: Automatically identifies and merges header rows
- 📝 **Detailed Logging**: Generates comprehensive logs of the merge process
- 🔧 **Auto Column Sizing**: Automatically adjusts column widths for better readability
- 📋 **Summary Reports**: Provides detailed success/failure reports

## Requirements

Before running the script, you need to install the following Python packages:

```bash
pip install pandas openpyxl msoffcrypto-tool
```

### Package Details:
- **pandas**: For data manipulation and Excel file handling
- **openpyxl**: For reading/writing Excel files with formatting support
- **msoffcrypto-tool**: For handling password-protected Microsoft Office files

### Built-in Python Libraries (No installation needed):
- `os`, `glob`, `pathlib`: File system operations
- `io`: Input/output operations
- `datetime`: Date and time handling
- `logging`: Logging functionality

## Installation

1. **Clone or download** this repository to your local machine
2. **Install required packages** using pip:
   ```bash
   pip install pandas openpyxl msoffcrypto-tool
   ```
3. **Create a batch file** (for Windows users) to run the script easily

## How to Use

### Method 1: Double-click BAT file (Windows - Recommended)

1. **Create a batch file** named `run_merger.bat` in the same directory as your Python script:
   ```batch
   @echo off
   echo Starting Excel Merger...
   python mergeTable.py
   pause
   ```

2. **Place your Excel files** in the same directory as the script and batch file

3. **Double-click the `run_merger.bat` file** to run the merger

4. **Check the results**:
   - Merged file will be saved as `1.xlsx`
   - Detailed logs will be saved in the `logs/` folder
   - Console will show real-time progress

### Method 2: Command Line

1. **Open Command Prompt** or Terminal
2. **Navigate** to the directory containing the script:
   ```bash
   cd path/to/your/script/directory
   ```
3. **Run the script**:
   ```bash
   python mergeTable.py
   ```

### Method 3: Python IDE

1. **Open** `mergeTable.py` in your preferred Python IDE
2. **Ensure** all Excel files are in the same directory
3. **Run** the script directly from the IDE

## Configuration

### Password Settings
By default, the script uses `"8888"` as the password for protected files. To change this:

1. **Edit the script**: Open `mergeTable.py`
2. **Find this line** (near the bottom):
   ```python
   DEFAULT_PASSWORD = "8888"  # Change this to your actual password
   ```
3. **Replace `"8888"`** with your actual password
4. **Save** the file

### Output File Name
To change the output filename:

1. **Edit the script**: Open `mergeTable.py`
2. **Find this line**:
   ```python
   OUTPUT_FILE = "1.xlsx"
   ```
3. **Change** `"1.xlsx"` to your desired filename
4. **Save** the file

## File Structure

Your working directory should look like this:
```
📁 Your Project Folder/
├── 📄 mergeTable.py          # Main Python script
├── 📄 run_merger.bat         # Batch file for easy execution (optional)
├── 📄 file1.xlsx            # Excel files to merge
├── 📄 file2.xlsx            # 
├── 📄 file3.xlsm            # Supports .xlsx, .xls, .xlsm formats
├── 📁 logs/                 # Created automatically
│   └── 📄 excel_merger_YYYYMMDD_HHMMSS.log
└── 📄 1.xlsx               # Output file (created after running)
```

## Supported File Formats

- ✅ `.xlsx` (Excel 2007+)
- ✅ `.xls` (Excel 97-2003)
- ✅ `.xlsm` (Excel Macro-Enabled)
- ✅ Password-protected files (all formats above)

## How It Works

1. **Scans Directory**: Automatically finds all Excel files in the current directory
2. **Password Detection**: Checks if files are password-protected
3. **Header Recognition**: Identifies header rows containing keywords like:
   - 跟团号 (Group Number)
   - 下单人 (Order Person)
   - 团员备注 (Member Notes)
   - 支付时间 (Payment Time)
   - And more...
4. **Data Extraction**: Extracts data rows while preserving formatting
5. **Merge Process**: Combines all data into a single worksheet
6. **Formatting**: Maintains original cell formatting and auto-adjusts column widths
7. **Output**: Saves merged file and generates detailed logs

## Logging

The script creates detailed logs in the `logs/` folder:
- **Filename format**: `excel_merger_YYYYMMDD_HHMMSS.log`
- **Contains**:
  - Processing status for each file
  - Number of rows processed
  - Success/failure reports
  - Error messages and troubleshooting info

## Troubleshooting

### Common Issues:

1. **"Missing required packages" error**:
   ```bash
   pip install pandas openpyxl msoffcrypto-tool
   ```

2. **Password-protected files not opening**:
   - Check if the password in the script matches your file password
   - Edit `DEFAULT_PASSWORD = "8888"` in the script

3. **No Excel files found**:
   - Ensure Excel files are in the same directory as the script
   - Check file extensions (.xlsx, .xls, .xlsm)

4. **Script not running on double-click**:
   - Ensure Python is installed and added to PATH
   - Try running from command line first
   - Check if the batch file is in the same directory

### Getting Help:

- Check the log files in the `logs/` folder for detailed error information
- Ensure all Excel files are closed before running the merger
- Verify that you have read/write permissions in the directory

## Example Output

After successful execution, you'll see:
```
🚀 Automated Excel Files Merger
==================================================
📂 Starting formatted merge process...
📋 Found 5 Excel files to process
📖 Processing file 1/5: data1.xlsx
✅ data1.xlsx - Added 150 data rows with formatting
📖 Processing file 2/5: data2.xlsx
✅ data2.xlsx - Added 200 data rows with formatting
...
🎉 Successfully merged 5 files with formatting preserved
📊 Total data rows in merged file: 1000
💾 Output saved as: 1.xlsx
```

## License

This project is open source and available under the MIT License.
