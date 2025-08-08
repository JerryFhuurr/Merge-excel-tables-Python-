import pandas as pd
import os
import glob
from pathlib import Path
import msoffcrypto
import io
from datetime import datetime
import logging
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

class ExcelMerger:
    def __init__(self, default_password="8888"):
        """
        Initialize Excel Merger
        
        Args:
            default_password (str): Default password to try for protected files
        """
        self.default_password = default_password
        self.setup_logging()
        
    def setup_logging(self):
        """Set up logging configuration"""
        # Create logs directory if it doesn't exist
        if not os.path.exists('logs'):
            os.makedirs('logs')
        
        # Set up logging
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        log_filename = f'logs/excel_merger_{timestamp}.log'
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_filename, encoding='utf-8'),
                logging.StreamHandler()  # Also print to console
            ]
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info(f"Excel Merger started - Log file: {log_filename}")
    
    def is_password_protected(self, file_path):
        """
        Check if Excel file is password protected
        
        Args:
            file_path (str): Path to Excel file
            
        Returns:
            bool: True if password protected, False otherwise
        """
        try:
            # Try to read without password first
            pd.read_excel(file_path, nrows=0)  # Just read header
            return False
        except Exception as e:
            error_msg = str(e).lower()
            if 'password' in error_msg or 'encrypted' in error_msg or 'protected' in error_msg:
                return True
            # Try with msoffcrypto to detect encryption
            try:
                with open(file_path, 'rb') as file:
                    office_file = msoffcrypto.OfficeFile(file)
                    if office_file.is_encrypted():
                        return True
            except:
                pass
            return False
    
    def read_excel_with_formatting(self, file_path, password=None):
        """
        Read Excel file with formatting preserved using openpyxl
        
        Args:
            file_path (str): Path to the Excel file
            password (str): Password for protected files (optional)
            
        Returns:
            tuple: (openpyxl.Worksheet or None, success_status, error_message)
        """
        filename = os.path.basename(file_path)
        
        try:
            if password:
                # Handle password-protected files
                with open(file_path, 'rb') as file:
                    office_file = msoffcrypto.OfficeFile(file)
                    office_file.load_key(password=password)
                    
                    decrypted = io.BytesIO()
                    
                    # Try both methods for compatibility
                    try:
                        office_file.save(decrypted)  # Older version
                    except AttributeError:
                        office_file.decrypt(decrypted)  # Newer version
                    
                    decrypted.seek(0)
                    workbook = load_workbook(decrypted)
                    worksheet = workbook.active
                    
                    return worksheet, True, None
            else:
                # Handle regular files
                workbook = load_workbook(file_path)
                worksheet = workbook.active
                return worksheet, True, None
                
        except Exception as e:
            error_msg = str(e)
            self.logger.error(f"❌ {filename} - Failed to read with formatting: {error_msg}")
            return None, False, error_msg

    def copy_cell_formatting(self, source_cell, target_cell):
        """
        Copy formatting from source cell to target cell
        """
        if source_cell.font:
            target_cell.font = Font(
                name=source_cell.font.name,
                size=source_cell.font.size,
                bold=source_cell.font.bold,
                italic=source_cell.font.italic,
                underline=source_cell.font.underline,
                strike=source_cell.font.strike,
                color=source_cell.font.color
            )
        
        if source_cell.fill:
            target_cell.fill = PatternFill(
                fill_type=source_cell.fill.fill_type,
                start_color=source_cell.fill.start_color,
                end_color=source_cell.fill.end_color
            )
        
        if source_cell.border:
            target_cell.border = Border(
                left=source_cell.border.left,
                right=source_cell.border.right,
                top=source_cell.border.top,
                bottom=source_cell.border.bottom
            )
        
        if source_cell.alignment:
            target_cell.alignment = Alignment(
                horizontal=source_cell.alignment.horizontal,
                vertical=source_cell.alignment.vertical,
                wrap_text=source_cell.alignment.wrap_text
            )
    
    def is_header_row(self, row):
        """
        Check if a row looks like a header row
        Header rows typically contain text like: 跟团号, 下单人, 团员备注, etc.
        """
        header_keywords = ['跟团号', '下单人', '团员备注', '支付时间', '团长备注', '商品', 
                          '订单金额', '退款金额', '订单状态', '自提点', '收货人', '联系电话', '详细地址']
        
        row_values = [str(cell.value).strip() if cell.value is not None else '' for cell in row]
        
        # Check if at least 3 header keywords are found in this row
        matching_keywords = sum(1 for keyword in header_keywords if any(keyword in value for value in row_values))
        
        return matching_keywords >= 3
    
    def extract_header_from_row(self, row):
        """
        Extract header values from a detected header row
        
        Args:
            row: openpyxl row object
            
        Returns:
            list: List of header values
        """
        header_values = []
        for cell in row:
            if cell.value is not None:
                header_values.append(str(cell.value).strip())
            else:
                header_values.append('')
        return header_values

    def merge_excel_files_with_formatting(self, folder_path=".", output_file='1.xlsx'):
        """
        Merge multiple Excel files with formatting preserved
        
        Args:
            folder_path (str): Path to folder containing Excel files
            output_file (str): Name of output file
        """
        
        self.logger.info(f"📂 Starting formatted merge process in folder: {os.path.abspath(folder_path)}")
        
        # Find all Excel files in the current directory
        excel_extensions = ['*.xlsx', '*.xls', '*.xlsm']
        excel_files = []
        
        for extension in excel_extensions:
            excel_files.extend(glob.glob(os.path.join(folder_path, extension)))
        
        # Remove the output file from the list if it exists
        excel_files = [f for f in excel_files if not f.endswith(output_file)]
        
        if not excel_files:
            self.logger.warning(f"⚠️ No Excel files found in {folder_path}")
            return
        
        self.logger.info(f"📋 Found {len(excel_files)} Excel files to process")
        
        # Create new workbook for output
        output_workbook = Workbook()
        output_worksheet = output_workbook.active
        output_worksheet.title = "Merged Data"
        
        successful_files = []
        failed_files = []
        current_row = 1
        header_added = False
        detected_headers = None  # Store the actual headers from source files
        
        # Process each file
        for i, file_path in enumerate(excel_files):
            filename = os.path.basename(file_path)
            self.logger.info(f"📖 Processing file {i+1}/{len(excel_files)}: {filename}")
            
            # Determine if file is password protected
            is_protected = self.is_password_protected(file_path)
            password = self.default_password if is_protected else None
            
            # Read with formatting preserved
            worksheet, success, error = self.read_excel_with_formatting(file_path, password)
            
            if success and worksheet is not None:
                rows_data = list(worksheet.iter_rows())
                
                # Debug logging
                self.logger.info(f"🔍 {filename} - Total rows found: {len(rows_data)}")
                
                # Check if worksheet has any rows
                if len(rows_data) == 0:
                    self.logger.warning(f"⚠️ {filename} - File is completely empty, skipping")
                    failed_files.append((filename, "File is completely empty"))
                    continue
                
                # Find header row and data rows
                header_row_index = -1
                data_rows = []
                
                # Look for header row
                for idx, row in enumerate(rows_data):
                    if self.is_header_row(row):
                        header_row_index = idx
                        self.logger.info(f"🔍 {filename} - Header found at row {idx + 1}")
                        break
                
                if header_row_index >= 0:
                    # Found header, get data rows after header
                    header_row = rows_data[header_row_index]
                    
                    # Extract and store header values if not done yet
                    if detected_headers is None:
                        detected_headers = self.extract_header_from_row(header_row)
                        self.logger.info(f"🔍 {filename} - Detected headers: {detected_headers[:5]}...")  # Show first 5
                    
                    potential_data_rows = rows_data[header_row_index + 1:]
                    
                    # Filter non-empty data rows
                    for row in potential_data_rows:
                        row_values = [cell.value for cell in row if cell.value is not None and str(cell.value).strip() != '']
                        if row_values:  # Row has actual data
                            data_rows.append(row)
                    
                    self.logger.info(f"🔍 {filename} - Found {len(data_rows)} data rows after header")
                    
                else:
                    # No header found, treat all non-empty rows as data
                    self.logger.info(f"🔍 {filename} - No header found, treating all rows as data")
                    for row in rows_data:
                        row_values = [cell.value for cell in row if cell.value is not None and str(cell.value).strip() != '']
                        if row_values:  # Row has actual data
                            data_rows.append(row)
                
                # Add header if not added yet and we have detected headers
                if not header_added and detected_headers is not None:
                    # Write the actual detected headers
                    for col_idx, header_value in enumerate(detected_headers, 1):
                        target_cell = output_worksheet.cell(row=current_row, column=col_idx)
                        target_cell.value = header_value
                        
                        # Apply header formatting if we have the original header row
                        if header_row_index >= 0:
                            source_cell = rows_data[header_row_index][col_idx - 1] if col_idx - 1 < len(rows_data[header_row_index]) else None
                            if source_cell:
                                self.copy_cell_formatting(source_cell, target_cell)
                            else:
                                # Apply basic header formatting
                                target_cell.font = Font(bold=True)
                        else:
                            # Apply basic header formatting
                            target_cell.font = Font(bold=True)
                    
                    current_row += 1
                    header_added = True
                    self.logger.info(f"📝 {filename} - Added detected header row with formatting")
                
                # Add data rows
                if data_rows:
                    for row in data_rows:
                        for col_idx, cell in enumerate(row, 1):
                            target_cell = output_worksheet.cell(row=current_row, column=col_idx)
                            target_cell.value = cell.value
                            self.copy_cell_formatting(cell, target_cell)
                        current_row += 1
                    
                    successful_files.append(filename)
                    self.logger.info(f"✅ {filename} - Added {len(data_rows)} data rows with formatting")
                    
                else:
                    self.logger.warning(f"⚠️ {filename} - No data rows found, skipping")
                    failed_files.append((filename, "No data rows found"))
                    
            else:
                failed_files.append((filename, error))
        
        # Auto-adjust column widths
        self.logger.info("📐 Auto-adjusting column widths...")
        for column in output_worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            output_worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Generate summary report
        self.generate_summary_report(successful_files, failed_files)
        
        if successful_files:
            # Save the formatted workbook
            output_workbook.save(output_file)
            
            total_data_rows = current_row - 2 if header_added else 0  # Subtract header
            self.logger.info(f"🎉 Successfully merged {len(successful_files)} files with formatting preserved")
            self.logger.info(f"📊 Total data rows in merged file: {total_data_rows}")
            self.logger.info(f"💾 Output saved as: {output_file}")
            
        else:
            self.logger.error("❌ No data was successfully read from any files")
        
        output_workbook.close()
    
    def generate_summary_report(self, successful_files, failed_files):
        """Generate a summary report of the merge process"""
        
        self.logger.info("=" * 60)
        self.logger.info("📊 MERGE SUMMARY REPORT")
        self.logger.info("=" * 60)
        
        self.logger.info(f"✅ Successfully processed files ({len(successful_files)}):")
        if successful_files:
            for i, filename in enumerate(successful_files, 1):
                self.logger.info(f"   {i}. {filename}")
        else:
            self.logger.info("   None")
        
        self.logger.info(f"\n❌ Failed to process files ({len(failed_files)}):")
        if failed_files:
            for i, (filename, error) in enumerate(failed_files, 1):
                self.logger.info(f"   {i}. {filename} - Reason: {error}")
        else:
            self.logger.info("   None")
        
        success_rate = len(successful_files) / (len(successful_files) + len(failed_files)) * 100 if (successful_files or failed_files) else 0
        self.logger.info(f"\n📈 Success Rate: {success_rate:.1f}%")
        self.logger.info("=" * 60)

def main():
    """
    Main function to run the Excel merger automatically
    """
    # Configuration - SET YOUR PASSWORD HERE
    DEFAULT_PASSWORD = "8888"  # Change this to your actual password
    OUTPUT_FILE = "1.xlsx"
    
    print("🚀 Automated Excel Files Merger")
    print("=" * 50)
    print("This script will:")
    print("- Auto-detect Excel files in current directory")
    print("- Auto-detect password protection")
    print("- Merge all files (excluding headers)")
    print("- Generate detailed logs")
    print("=" * 50)
    
    # Create merger instance
    merger = ExcelMerger(default_password=DEFAULT_PASSWORD)
    
    # Run the merger with formatting preserved
    merger.merge_excel_files_with_formatting(folder_path=".", output_file=OUTPUT_FILE)
    
    print("\n✨ Process completed! Check the log file for detailed information.")

if __name__ == "__main__":
    # Check required packages
    try:
        import pandas as pd
        import msoffcrypto
        from openpyxl import Workbook, load_workbook
        from openpyxl.styles import Font, PatternFill, Border, Alignment
    except ImportError as e:
        print("Missing required packages. Please install them using:")
        print("pip install pandas openpyxl msoffcrypto-tool")
        print(f"Error: {e}")
        exit(1)
    
    main()