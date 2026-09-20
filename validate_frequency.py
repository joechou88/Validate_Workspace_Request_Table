import os
import logging
from datetime import datetime
import openpyxl
import config

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_filename = f"frequency_log_{timestamp}.txt"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(log_filename, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

def validate_excel_frequency(filepath, expected_frequency):
    filename = os.path.basename(filepath)

    try:
        workbook = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        if 'REQUEST_TABLE' not in workbook.sheetnames:
            return f"[ERROR] Sheet 'REQUEST_TABLE' missing in {filename}"

        worksheet = workbook['REQUEST_TABLE']
        current_row = 7
        
        # Track if we actually checked any data
        has_data = False

        # Loop using Column G (Start Date) as an "anchor" column because Column I (Frequency) can legitimately be blank for 'annual' requests
        while worksheet[f"G{current_row}"].value is not None:
            has_data = True
            
            # Read the frequency from column I
            cell_value = worksheet[f"I{current_row}"].value
            
            # If the cell is blank (None or empty string), default to 'annual'
            if cell_value is None or str(cell_value).strip() == "":
                actual_frequency = "annual"
            else:
                actual_frequency = str(cell_value).strip().lower()
                if actual_frequency == "yearly":
                    actual_frequency = "annual"

            if actual_frequency != expected_frequency:
                return f"[MISMATCH] {filename} row {current_row}: Expected '{expected_frequency}' from folder name, found '{actual_frequency}' in request table."
            
            current_row += 1

        if not has_data:
            return f"[ERROR] Request table starting at row 7 is empty for {filename}"

        return f"[OK] {filename}"

    except Exception as exception:
        return f"[ERROR] Failed to read {filename}: {str(exception)}"
    finally:
        if 'workbook' in locals():
            workbook.close()

base_directory = '.' 
error_list = []
logging.info("Starting frequency validation process...")

for root_directory, directories, files in os.walk(base_directory):
    # Mutate directories in-place to skip excluded folders
    directories[:] = [directory for directory in directories if directory not in config.EXCLUDED_FOLDERS]
    
    folder_name = os.path.basename(root_directory).lower()
    
    # Determine the expected frequency from the folder prefix
    expected_frequency = None
    if folder_name.startswith('annual'):
        expected_frequency = 'annual'
    elif folder_name.startswith('quarterly'):
        expected_frequency = 'quarterly'
    elif folder_name.startswith('daily'):
        expected_frequency = 'daily'

    # Only process files if the folder corresponds to a known frequency
    if expected_frequency:
        logging.info(f"===== Scanning directory: {root_directory} (Expected: {expected_frequency}) =====")
        
        for file_name in files:
            if file_name.endswith('.xlsm') and not file_name.startswith('~'):
                filepath = os.path.join(root_directory, file_name)
                result = validate_excel_frequency(filepath, expected_frequency)
                
                if result.startswith('[MISMATCH]') or result.startswith('[ERROR]'):
                    logging.error(result)
                    error_list.append(result)
                elif result.startswith('[SKIPPED]'):
                    logging.warning(result)
                else:
                    logging.info(result)

if error_list:
    logging.warning("--- Errors Found ---")
    with open('frequency_manual_error_list.txt', 'w', encoding='utf-8') as error_file:
        for error_message in error_list:
            error_file.write(error_message + '\n')
    logging.info("Errors saved to frequency_manual_error_list.txt")
else:
    logging.info("\n--- All checks passed ---")
