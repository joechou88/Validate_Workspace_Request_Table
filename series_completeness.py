import os
import re
import logging
import warnings
import openpyxl
import config
from utility import Utility

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

util = Utility()
base_directory = '.'
missing_files_list = []

logging.info("Starting insurance series completeness validation...")

# Group files by their base name (removing the -1, -2 parts)
file_groups = {}
category_mapping = {}

for root_directory, directories, files in os.walk(base_directory):
    if hasattr(config, 'EXCLUDED_FOLDERS'):
        directories[:] = [directory for directory in directories if directory not in config.EXCLUDED_FOLDERS]
        
    for file_name in files:
        if not file_name.endswith('.xlsm') or file_name.startswith('~'):
            continue
            
        if "life-insurance" not in file_name:
            continue

        # Create a unified group key by removing the trailing number from the category
        # Example: life-insurance-1_quarterly-financial_2005.xlsm -> life-insurance_quarterly-financial_2005.xlsm
        unified_file_name = re.sub(r'(life-insurance|non-life-insurance)-\d+', r'\1', file_name)
        group_key = os.path.join(root_directory, unified_file_name)
        
        if "non-life-insurance" in file_name:
            base_category = "non-life-insurance"
        else:
            base_category = "life-insurance"

        filepath = os.path.join(root_directory, file_name)
        
        if group_key not in file_groups:
            file_groups[group_key] = []
            category_mapping[group_key] = base_category
            
        file_groups[group_key].append(filepath)

for group_key, filepaths in file_groups.items():
    base_category = category_mapping[group_key]
    expected_series = set(config.EXPECTED_SERIES.get(base_category, []))
    
    if not expected_series:
        continue
        
    found_series = set()
    
    for filepath in filepaths:
        logging.info(f"Processing: {filepath}")
        try:
            workbook = openpyxl.load_workbook(filepath, data_only=True)
            sheet = workbook["REQUEST_TABLE"]

            # Read Column E (index 5) starting from row 7 downwards directly for all files
            for row in sheet.iter_rows(min_row=7, min_col=5, max_col=5, values_only=True):
                cell_value = row[0]
                if cell_value is not None:
                    for item in str(cell_value).split(','):
                        if item.strip():
                            found_series.add(item.strip())
                    
            workbook.close()
        except Exception as e:
            logging.error(f"Error reading {filepath}: {e}")

    missing_series = expected_series - found_series
    
    if missing_series:
        missing_record = f"{group_key} is missing {len(missing_series)} series: {list(missing_series)}"
        missing_files_list.append(missing_record)
        logging.warning(f"Missing in {group_key}: {len(missing_series)} series")

if missing_files_list:
    logging.warning(f"Found {len(missing_files_list)} incomplete groups. Saving to series_incomplete_list.txt")
    with open('series_incomplete_list.txt', 'w', encoding='utf-8') as output_file:
        for missing_record in missing_files_list:
            output_file.write(missing_record + '\n')
else:
    logging.info("\n--- All insurance files have complete series ---")
    