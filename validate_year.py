import os
import logging
from datetime import datetime
import openpyxl
import config
from utility import Utility

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_filename = f"year_log_{timestamp}.txt"

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

def get_year_from_request_table(cell_value):
    if not cell_value:
        return None
    if isinstance(cell_value, datetime):
        return cell_value.year
    match = config.CELL_YEAR_PATTERN.search(str(cell_value))
    return int(match.group(1)) if match else None

def validate_excel_year(filepath, frequency):
    filename = os.path.basename(filepath)
    file_info = util.parse_filename(filename)

    if not file_info:
        return f"[SKIPPED] No year found or unmatched pattern in filename: {filename}"

    file_start_year = file_info["start_year"]
    file_end_year = file_info["end_year"] if file_info["end_year"] else file_start_year

    try:
        workbook = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        if 'REQUEST_TABLE' not in workbook.sheetnames:
            return f"[ERROR] Sheet 'REQUEST_TABLE' missing in {filename}"

        worksheet = workbook['REQUEST_TABLE']

        if frequency == 'annual':
            file_years_range = list(range(file_start_year, file_end_year + 1))
            for index, file_year in enumerate(file_years_range):
                cell_value = worksheet[f"G{7 + index}"].value
                request_table_year = get_year_from_request_table(cell_value)
                if request_table_year != file_year:
                    return f"[MISMATCH] {filename} row {7 + index}: Expected {file_year} in filename, found {request_table_year} in request table."
        else:
            current_row = 7
            
            while worksheet[f"G{current_row}"].value is not None:
                start_value = worksheet[f"G{current_row}"].value
                end_value = worksheet[f"H{current_row}"].value

                start_year = get_year_from_request_table(start_value)
                end_year = get_year_from_request_table(end_value)

                if not start_year or not end_year:
                    return f"[ERROR] Invalid dates in G{current_row}/H{current_row} for {filename}"

                # Checks if the Excel dates safely cover the filename dates
                if start_year > file_start_year or end_year < file_end_year:
                    return f"[MISMATCH] {filename} row {current_row} covers ({start_year}-{end_year}), but file requires ({file_start_year}-{file_end_year})"
                
                current_row += 1

        return f"[OK] {filename}"

    except Exception as exception:
        return f"[ERROR] Failed to read {filename}: {str(exception)}"
    finally:
        if 'workbook' in locals():
            workbook.close()

base_directory = '.' 
error_list = []
logging.info("Starting year validation process...")

for root_directory, directories, files in os.walk(base_directory):
    # Mutate directories in-place to skip excluded folders
    directories[:] = [directory for directory in directories if directory not in config.EXCLUDED_FOLDERS]
    folder_name = os.path.basename(root_directory).lower()

    logging.info(f"===== Scanning directory: {root_directory} =====")
    
    frequency = 'annual'
    if 'quarterly' in folder_name:
        frequency = 'quarterly'
    elif 'daily' in folder_name:
        frequency = 'daily'

    for file_name in files:
        if file_name.endswith('.xlsm') and not file_name.startswith('~'):
            filepath = os.path.join(root_directory, file_name)
            result = validate_excel_year(filepath, frequency)
            if result.startswith('[MISMATCH]') or result.startswith('[ERROR]'):
                logging.error(result)
                error_list.append(result)
            elif result.startswith('[SKIPPED]'):
                logging.warning(result)
            else:
                logging.info(result)

if error_list:
    logging.warning("--- Errors Found ---")
    with open('year_manual_error_list.txt', 'w', encoding='utf-8') as error_file:
        for error_message in error_list:
            error_file.write(error_message + '\n')
    logging.info("Errors saved to year_manual_error_list.txt")
else:
    logging.info("\n--- All checks passed ---")
