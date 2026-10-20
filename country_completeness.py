import os
import logging
import config
from utility import Utility

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

util = Utility()
base_directory = '.'
missing_files_list = []

logging.info("Starting country completeness validation...")

VALIDATE_START_YEAR = 2015
VALIDATE_END_YEAR = 2025

for root_directory, directories, files in os.walk(base_directory):
    # Exclude folders defined in config.EXCLUDED_FOLDERS
    directories[:] = [directory for directory in directories if directory not in config.EXCLUDED_FOLDERS]
    folder_name = os.path.basename(root_directory)
    
    # Track covered years for each country in this folder
    country_coverage = {country: set() for country in config.REQUIRED_COUNTRIES}
    has_country_files = False
    
    for file_name in files:
        if file_name.endswith('.xlsm') and not file_name.startswith('~'):
            filepath = os.path.join(root_directory, file_name)
            logging.info(f"Processing: {filepath}")
            file_info = util.parse_filename(file_name)

            if file_info:
                country = file_info.get("country")
                company = file_info.get("company")
                country = f"{country}{company}" if company else country

                if country in config.REQUIRED_COUNTRIES:
                    has_country_files = True
                    start_year = file_info.get("start_year")
                    end_year = file_info.get("end_year")
                    
                    for year in range(start_year, end_year + 1):
                        country_coverage[country].add(year)
    
    # Only validate folders that actually contain country files
    if has_country_files:
        for year in range(VALIDATE_START_YEAR, VALIDATE_END_YEAR + 1):
            expected_countries = config.REQUIRED_COUNTRIES.copy()
            
            # Specific rule: due to a Workspace series rename, daily_return_global before 2024 only needs Germany1 and Germany2 (split into 1-8 from 2024).
            if folder_name == "daily_return_global" and year <= 2023:
                expected_countries -= {"Germany3", "Germany4", "Germany5", "Germany6", "Germany7", "Germany8"}
            
            for country in sorted(expected_countries):
                if year not in country_coverage[country]:
                    # Construct missing record as "folder_name/Country_Year.xlsm"
                    missing_record = f"{folder_name}/{country}_{year}.xlsm"
                    missing_files_list.append(missing_record)
                    logging.warning(f"Missing: {missing_record}")

if missing_files_list:
    logging.warning(f"Found {len(missing_files_list)} missing files. Saving to country_incomplete_list.txt")
    with open('country_incomplete_list.txt', 'w', encoding='utf-8') as output_file:
        for missing_file in missing_files_list:
            output_file.write(missing_file + '\n')
else:
    logging.info("\n--- All folders have complete country files ---")
