## General guidelines
* **File Format:** The validation tools strictly process `.xlsm` files. Files downloaded from Workspace are saved as macro-enabled workbooks to retain the Workspace add-in functionality for future edits.
* **Filename Naming Convention:** All `.xlsm` files must adhere to a strict naming pattern defined in `config.py` using the structure: `[Country/Category][CompanyID]_[IntermediateText]_[StartYear]-[EndYear]_[Suffix].xlsm`
  * **Country / Category (Required):** The subject of the file (e.g., `Germany`, `New-Zealand`, `financial`, `life-insurance`).
  * **Company ID (Optional):** Numeric digits appended directly to the subject (e.g., the `1` in `Germany1`).
  * **Intermediate Text (Optional):** Descriptive text placed before the year (e.g., `quarterly-financial`). The script safely ignores this without confusing it for a year.
  * **Year Range (Required):** A 4-digit start year, optionally followed by a hyphen and a 4-digit end year (e.g., `2015`, `2005-2025`).
  * **Suffix (Optional):** An underscore followed by alphabetical characters (e.g., `_A`, `_BCD`).
  * **Examples Breakdown:**
    * `Germany1_2015_A.xlsm` *(Country: Germany | ID: 1 | Start: 2015 | Suffix: A)*
    * `New-Zealand_2015-2017.xlsm` *(Country: New-Zealand | Years: 2015-2017)*
    * `Switzerland_2015-2018_A.xlsm` *(Country: Switzerland | Years: 2015-2018 | Suffix: A)*
    * `life-insurance_annual-financial_2005-2025.xlsm` *(Category: life-insurance | Intermediate: annual-financial | Years: 2005-2025)*
    * `life-insurance-1_quarterly-financial_2005.xlsm` *(Category: life-insurance-1 | Intermediate: quarterly-financial | Start: 2005)*
* **Excluded Directories:** The following folders are skipped during validation:
  * `insurance_company_list/`: Contains intermediate helper files used to generate insurance company lists; no validation needed.
  * `annual_audit_fee/`: Does not use a standard `REQUEST_TABLE`. Parameters must be checked if updated correctly within the Excel formulas.
* **Logging & Error Tracking:** 
  * Full execution details are saved to timestamped log files depending on the script run (e.g., `year_log_{timestamp}.txt`, `series_log_{timestamp}.txt`, etc.).
  * Any mismatches or errors requiring human intervention are printed to the console and compiled into a dedicated summary file (e.g., `year_manual_error_list.txt`) for quick review.

## Validate year range
The script validates whether the requested time range in the `REQUEST_TABLE` (starting from row 7 downwards) correctly aligns with the years specified in the filename. The validation logic differs based on the data frequency:

* **Annual:** Validates sequentially row-by-row. For example, if the filename indicates `2015-2017`, the script strictly expects G7=2015, G8=2016, and G9=2017.
* **Quarterly / Daily:** Validates by range coverage. The requested dates in the Excel file (Column G for start date, Column H for end date) must fully cover or exceed the year range indicated in the filename. The script loops downwards to validate multiple subjects if present.
  * *Example:* If the filename specifies `2005-2025`, a request table with `G7=10/12/2004` and `H7=20/01/2026` is accepted. This inclusive logic allows us to pull a slightly wider data range when necessary.

## Validate series
The script validates whether the series mnemonics in the `REQUEST_TABLE` (starting from column E, row 7 downwards) match the expected mnemonic for the country specified in the filename.

* **Validation Logic:** The script extracts the country and optional company ID from the filename (e.g., `Germany` + `1` = `Germany1`), looks up the expected series in `config.COUNTRY_MNEMONIC`, and checks it against column E. It loops downwards until an empty cell is encountered.
* **Smart Skipping:** If the filename indicates a category (e.g., `financial_2005-2025.xlsm`) rather than a country, or if the country is not listed in the dictionary, the script safely logs a `[SKIPPED]` message and moves on without throwing an error.
* **Background Context (Germany):** The original general series for Germany (`FDEALL`) is no longer valid. As a result, Germany is now divided into 8 specific company groups (where `Germany1` maps to `FDEALL1`, up to `FDEALL8`). All 8 files must be processed and later combined to assemble the complete German dataset.

## Validate frequency
The script validates whether the data frequency specified in the `REQUEST_TABLE` (starting from column I, row 7 downwards) matches the frequency indicated by its parent folder's prefix (e.g., `annual_`, `quarterly_`, `daily_`).

* **Annual Normalization:** 
  * If a cell in column I is **blank**, the script defaults the frequency to `annual`.
  * If a cell contains **`Yearly`** (the standard option from the Excel dropdown), the script normalizes it to `annual`.
  * Both of these cases will successfully pass validation when the file is located inside an `annual_` folder.

## Validate datatype
The script verifies that the requested variables (datatypes) in column F of the `REQUEST_TABLE` exactly match the expected complete list for that specific dataset. You can flexibly add new folders and their corresponding datatype rules to `config.EXPECTED_DATATYPES` to expand validation capabilities.

* **Standard Validation:** For most folders, the script checks the datatypes row-by-row against the expected list. Any missing variables will be explicitly listed in the error log.
* **Multi-File Aggregation (`annual_financials_global`):** 
  * Because the number of active companies fluctuates by year, large datasets often exceed Workspace download limits. This requires variables to be split into multiple files (indicated by the filename suffix, e.g., `_A`, `_BCD`).
  * To handle this, the script intelligently groups files and **aggregates datatypes by Country and Year** before verifying completeness. As long as the combined set of variables across all split files is complete, the validation passes.
  * *Single-year split example:* Variables from `Vietnam_2023_ABC.xlsm` and `Vietnam_2023_D.xlsm` are automatically combined and validated as a single 2023 dataset.
  * *Multi-year split example:* For files like `Switzerland_2015-2018_A.xlsm` through `Switzerland_2015-2018_E.xlsm`, the script correctly maps rows to specific years (e.g., F7 for 2015, F8 for 2016, etc.), aggregates the datatypes across all suffix files year-by-year, and validates each year independently.
