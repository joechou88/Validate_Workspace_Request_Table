import os
import logging
from datetime import datetime
import openpyxl
import config
from utility import Utility

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_filename = f"series_log_{timestamp}.txt"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(log_filename, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

util = Utility()

def validate_excel_series(filepath):
    filename = os.path.basename(filepath)
    file_info = util.parse_filename(filename)

    if not file_info:
        return f"[SKIPPED] Unmatched pattern in filename: {filename}"

    file_country = file_info.get("country")
    
    # If the file is a category (e.g., 'financial') instead of a country, skip it safely
    if not file_country:
        return f"[SKIPPED] No country found in filename, skipping series check: {filename}"

    # Reconstruct the full country name (e.g., "Germany" + "1" -> "Germany1") to match the dictionary keys
    file_company = str(file_info["company"]) if file_info["company"] is not None else ""
    file_full_country = f"{file_country}{file_company}"

    file_series = config.COUNTRY_MNEMONIC.get(file_full_country)

    # If the file category (like 'life-insurance') is not in the mnemonic dictionary, we safely skip it
    if not file_series:
        return f"[SKIPPED] No series mapping found in config for: {file_full_country} ({filename})"

    try:
        workbook = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        if 'REQUEST_TABLE' not in workbook.sheetnames:
            return f"[ERROR] Sheet 'REQUEST_TABLE' missing in {filename}"

        worksheet = workbook['REQUEST_TABLE']
        current_row = 7

        # Track if we actually checked any data
        has_data = False

        # Loop downwards in column E until an empty cell is encountered
        while worksheet[f"E{current_row}"].value is not None:
            has_data = True
            series = str(worksheet[f"E{current_row}"].value).strip()

            if series != file_series:
                return f"[MISMATCH] {filename} row {current_row}: Expected '{file_series}', found '{series}' in request table."
            
            current_row += 1

        if not has_data:
            return f"[ERROR] Column E starting at row 7 is empty for {filename}"

        return f"[OK] {filename}"

    except Exception as exception:
        return f"[ERROR] Failed to read {filename}: {str(exception)}"
    finally:
        if 'workbook' in locals():
            workbook.close()

base_directory = '.' 
error_list = []
logging.info("Starting series validation process...")

for root_directory, directories, files in os.walk(base_directory):
    # Mutate directories in-place to skip excluded folders
    directories[:] = [directory for directory in directories if directory not in config.EXCLUDED_FOLDERS]
    
    logging.info(f"===== Scanning directory: {root_directory} =====")

    for file_name in files:
        if file_name.endswith('.xlsm') and not file_name.startswith('~'):
            filepath = os.path.join(root_directory, file_name)
            result = validate_excel_series(filepath)
            
            if result.startswith('[MISMATCH]') or result.startswith('[ERROR]'):
                logging.error(result)
                error_list.append(result)
            elif result.startswith('[SKIPPED]'):
                logging.warning(result)
            else:
                logging.info(result)

if error_list:
    logging.warning("--- Errors Found ---")
    with open('series_manual_error_list.txt', 'w', encoding='utf-8') as error_file:
        for error_message in error_list:
            error_file.write(error_message + '\n')
    logging.info("Errors saved to series_manual_error_list.txt")
else:
    logging.info("\n--- All checks passed ---")
