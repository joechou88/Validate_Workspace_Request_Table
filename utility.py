import config
import os
import logging

class Utility:
    def parse_filename(self, fname):
        name = os.path.splitext(os.path.basename(fname))[0]
        m = config.FILENAME_PATTERN.fullmatch(name)
        if not m:
            return None
        
        info = m.groupdict()
        
        return {
            'country': info.get('country') or info.get('category'),
            'company': int(info['company']) if info.get('company') else None,
            'start_year': int(info['start_year']),
            'end_year': int(info['end_year']) if info.get('end_year') else None,
            'suffix': list(info['suffix']) if info.get('suffix') and len(info['suffix']) > 1 else info.get('suffix')
        }

    def actual_rows(self, worksheet):
        last_row = 0
        for i, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=1):     # min_row=2: skipping the row 1 header
            if any(cell is not None for cell in row):
                last_row = i
        return last_row

    def actual_columns(self, worksheet):
        last_column = 0      
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            width = len(row)
            while width > 0 and row[width - 1] is None:
                width -= 1

            if width > last_column:
                last_column = width
                
        return last_column

    def check_request_table(self, wb, fname, request_sheet="REQUEST_TABLE"):
        """
        Checks the validity of the request table. 
        If series is provided (e.g. FDEALL1), it will check the E column.
        """
        file = self.parse_filename(fname)
        if not file:
            logging.warning(f"⚠️ Skip: {fname} unable to parse filename")
            return False
        
        ws = wb[request_sheet]
        start_year = file["start_year"]
        end_year = file["end_year"] if file["end_year"] else start_year
        years_in_file_name = list(range(start_year, end_year + 1))
        country_in_file_name = file['country']
        series_in_file_name = config.COUNTRY_MNEMONIC.get(country_in_file_name)
        expected_series = f"FDEALL{file['company']}" if file['company'] else None

        row = 7
        years_in_request_table = []

        while ws[f"E{row}"].value not in (None, ""):
            if expected_series and ws[f"E{row}"].value != expected_series:
                logging.warning(f"⚠️ Skip: {fname} {request_sheet} E{row} = {ws[f'E{row}'].value}, expected {expected_series}")
                return False

            raw_year = ws[f"G{row}"].value
            try:
                years_in_request_table.append(int(str(raw_year).strip()))
            except Exception:
                logging.warning(f"⚠️ Skip: {fname} {request_sheet} G{row} = {raw_year}, cannot be parsed as a year.")
                return False
                
            row += 1

        if years_in_request_table != years_in_file_name:
            logging.warning(f"⚠️ Skip: {fname} {request_sheet} years mismatch. Expected {years_in_file_name} in file name, got {years_in_request_table} in G column of request table.")
            return False
        
        return True

    def validate_wb(self, wb, fname, start_year, end_year, years, expected_series=None, request_sheet="REQUEST_TABLE"):
        """
        Validates workbook structure: presence of request table, series match, and sheet count.
        """
        if request_sheet not in wb.sheetnames:
            raise ValueError(f"{fname} missing {request_sheet}")
        
        self.check_request_table(wb, fname, request_sheet)
        
        data_sheets = [s for s in wb.sheetnames if s != request_sheet]
        if len(data_sheets) < years:
            raise ValueError(
                f"{fname} insufficient sheets. Expected {years}, got {len(data_sheets)}."
            )

    def print_sheet_shapes(self, wb, fname, skip_sheet="REQUEST_TABLE"):
        """Prints the actual rows and cols of all sheets in a workbook."""
        for ws_name in wb.sheetnames:
            if ws_name == skip_sheet:
                continue
            ws = wb[ws_name]
            rows = self.actual_rows(ws)
            columns = self.actual_columns(ws)
            logging.info(f"{fname} 🔹 Sheet: {ws_name}, shape: {rows} rows x {columns} columns")

