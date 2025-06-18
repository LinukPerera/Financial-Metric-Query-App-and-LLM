import openpyxl
from openpyxl.utils import get_column_letter
import re
import sys
import logging
from contextlib import contextmanager
from DebugHeader import (
    detect_first_nonempty_row,
    detect_multirow_sheet,
    extract_header_signature,
    find_repeating_headers,
    detect_sectors,
    detect_glossary,
    get_merged_parent,
)

# Set up logging to a file
logging.basicConfig(filename='debug.log', level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

@contextmanager
def suppress_prints():
    """Suppress print statements from DebugHeader.py."""
    original_stdout = sys.stdout
    sys.stdout = open('/dev/null', 'w')
    try:
        yield
    finally:
        sys.stdout.close()
        sys.stdout = original_stdout

class FinancialQueryProcessor:
    def __init__(self, file_path):
        with suppress_prints():
            self.wb = openpyxl.load_workbook(file_path, data_only=True)
            self.sheet_info = self._parse_sheets()
            self.primary_sheet = self.sheet_info[detect_multirow_sheet(self.wb)]
            logging.debug(f"Initialized with primary sheet: {self.primary_sheet['sheet'].title}")
            logging.debug(f"Sheet info keys: {list(self.sheet_info.keys())}")

    def _parse_sheets(self):
        """Parse all sheets to extract headers, sectors, and data ranges."""
        sheet_info = {}
        with suppress_prints():
            for sheet_name in self.wb.sheetnames:
                sheet = self.wb[sheet_name]
                header_start = detect_first_nonempty_row(sheet)
                is_multirow = (sheet_name == detect_multirow_sheet(self.wb))
                header_rows = [header_start] if not is_multirow else [header_start, header_start + 1]
                logging.debug(f"Parsing sheet: {sheet_name}, header_rows: {header_rows}")

                # Extract header mappings (column letter to header name)
                headers = {}
                code_columns = set()  # Track columns with 'CODE'
                for col in range(1, sheet.max_column + 1):
                    hierarchy = []
                    has_code = False
                    for row in header_rows:
                        if row > sheet.max_row:
                            continue
                        value, _, _ = get_merged_parent(sheet, row, col)
                        if value not in (None, ''):
                            hierarchy.append(str(value).strip())
                            if str(value).strip().upper() == 'CODE':
                                has_code = True
                    if has_code:
                        code_columns.add(col)
                    headers[get_column_letter(col)] = ' > '.join(hierarchy) if hierarchy else None

                # Set primary key: Column B if two CODE columns, else Column A
                code_count = len(code_columns)
                primary_key_col = 'B' if code_count == 2 else 'A'
                logging.debug(f"Sheet {sheet_name}: code_columns={code_columns}, code_count={code_count}, primary_key_col={primary_key_col}")
                logging.debug(f"Headers: {headers}")

                # Detect sectors and repeating headers
                repeating_headers = find_repeating_headers(sheet, header_rows)
                glossary_start = detect_glossary(sheet, header_rows, repeating_headers, [])
                sectors = detect_sectors(sheet, glossary_start)
                sector_rows = [row for row, _ in sectors]
                logging.debug(f"Sectors: {sectors}, Repeating headers: {repeating_headers}")

                # Determine data rows (exclude headers, repeating headers, sectors, and glossary)
                exclude_rows = set(header_rows + repeating_headers + sector_rows)
                if glossary_start:
                    exclude_rows.update(range(glossary_start, sheet.max_row + 1))
                data_rows = [r for r in range(header_rows[-1] + 1, sheet.max_row + 1) if r not in exclude_rows]
                logging.debug(f"Data rows: {data_rows[:5]}... (total {len(data_rows)})")

                sheet_info[sheet_name] = {
                    'sheet': sheet,
                    'headers': headers,
                    'header_rows': header_rows,
                    'sectors': sectors,
                    'data_rows': data_rows,
                    'primary_key_col': primary_key_col,
                }
        return sheet_info

    def _find_column(self, sheet_name, keyword, single_match=True):
        """Find the column(s) matching a keyword in the headers."""
        headers = self.sheet_info[sheet_name]['headers']
        keyword = keyword.lower()
        matches = []
        for col, header in headers.items():
            if header and keyword in header.lower():
                matches.append(col)
        logging.debug(f"Finding column for keyword '{keyword}' in {sheet_name}: matches={matches}")
        return matches[:1] if single_match and matches else matches

    def _find_company(self, sheet_name, code):
        """Find a company's data row by its code, preserving exact match."""
        sheet = self.sheet_info[sheet_name]['sheet']
        primary_key_col = self.sheet_info[sheet_name]['primary_key_col']
        col_idx = openpyxl.utils.column_index_from_string(primary_key_col) - 1
        logging.debug(f"Searching for code '{code}' in {sheet_name}, column {primary_key_col}")
        for row in self.sheet_info[sheet_name]['data_rows']:
            cell_value = sheet.cell(row=row, column=col_idx + 1).value
            if cell_value and str(cell_value).strip().upper() == code.upper():
                logging.debug(f"Found code '{code}' at row {row}")
                return row
        logging.debug(f"Code '{code}' not found in {sheet_name}")
        return None

    def _get_sector_rows(self, sheet_name, sector_code):
        """Get all data rows for a given sector."""
        sectors = self.sheet_info[sheet_name]['sectors']
        data_rows = self.sheet_info[sheet_name]['data_rows']
        for i, (sector_row, sector_name) in enumerate(sectors):
            if sector_name.upper() == sector_code.upper():
                next_sector_row = sectors[i + 1][0] if i + 1 < len(sectors) else self.sheet_info[sheet_name]['sheet'].max_row + 1
                rows = [r for r in data_rows if sector_row < r < next_sector_row]
                logging.debug(f"Sector '{sector_code}' rows: {rows}")
                return rows
        logging.debug(f"Sector '{sector_code}' not found in {sheet_name}")
        return []

    def process_query(self, query):
        """Process a natural language query and return the result."""
        query = query.lower().strip()
        logging.debug(f"Processing query: {query}")
        primary_sheet_name = self.primary_sheet['sheet'].title

        # Simplified query parsing
        metric_pattern = r'(.+?)\s+for\s+([a-z0-9\.]+)'
        sector_pattern = r'average\s+(.+?)\s+for\s+sector\s+([a-z\s&]+)'
        general_pattern = r'average\s+(.+)'

        # Handle company-specific queries (e.g., "P/E for ABL")
        metric_match = re.search(metric_pattern, query)
        if metric_match:
            metric, code = metric_match.groups()
            logging.debug(f"Company query: metric={metric}, code={code}")
            return self._handle_company_metric(code, metric)

        # Handle sector-specific queries (e.g., "average P/E for sector BANKS")
        sector_match = re.search(sector_pattern, query)
        if sector_match:
            metric, sector = sector_match.groups()
            logging.debug(f"Sector query: metric={metric}, sector={sector}")
            return self._handle_sector_metric(sector, metric)

        # Handle general metric queries (e.g., "average P/E")
        general_match = re.search(general_pattern, query)
        if general_match:
            metric = general_match.group(1)
            logging.debug(f"General query: metric={metric}")
            return self._handle_general_metric(metric)

        logging.debug("Query not understood")
        return "Sorry, I couldn't understand your query. Please specify a company, sector, or metric (e.g., 'P/E for ABL' or 'average P/E for sector BANKS')."

    def _handle_company_metric(self, code, metric):
        """Retrieve a specific metric for a company."""
        sheet_name = detect_multirow_sheet(self.wb)
        sheet = self.sheet_info[sheet_name]['sheet']
        row = self._find_company(sheet_name, code)
        if not row:
            return f"Company {code} not found in {sheet_name}."

        columns = self._find_column(sheet_name, metric, single_match=True)
        if not columns:
            return f"Metric '{metric}' not found in {sheet_name} headers."

        results = []
        for col in columns:
            value = sheet.cell(row=row, column=openpyxl.utils.column_index_from_string(col)).value
            header = self.sheet_info[sheet_name]['headers'][col]
            results.append(f"{header}: {value if value is not None else 'N/A'}")
        result = f"Data for {code} in {sheet_name}:\n" + "\n".join(results)
        logging.debug(f"Company metric result: {result}")
        return result

    def _handle_sector_metric(self, sector, metric):
        """Retrieve a metric aggregated across a sector."""
        sheet_name = detect_multirow_sheet(self.wb)
        columns = self._find_column(sheet_name, metric, single_match=False)
        if not columns:
            return f"Metric '{metric}' not found in {sheet_name} headers."

        rows = self._get_sector_rows(sheet_name, sector)
        if not rows:
            return f"Sector {sector} not found in {sheet_name}."

        sheet = self.sheet_info[sheet_name]['sheet']
        values = []
        for row in rows:
            for col in columns:
                value = sheet.cell(row=row, column=openpyxl.utils.column_index_from_string(col)).value
                if isinstance(value, (int, float)):
                    values.append(value)

        if not values:
            return f"No numerical data found for {metric} in sector {sector}."

        avg_value = sum(values) / len(values)
        result = f"Average {metric} for sector {sector} in {sheet_name}: {avg_value:.2f}"
        logging.debug(f"Sector metric result: {result}")
        return result

    def _handle_general_metric(self, metric):
        """Retrieve a metric across all data in the primary sheet."""
        sheet_name = detect_multirow_sheet(self.wb)
        columns = self._find_column(sheet_name, metric, single_match=False)
        if not columns:
            return f"Metric '{metric}' not found in {sheet_name} headers."

        sheet = self.sheet_info[sheet_name]['sheet']
        values = []
        for row in self.sheet_info[sheet_name]['data_rows']:
            for col in columns:
                value = sheet.cell(row=row, column=openpyxl.utils.column_index_from_string(col)).value
                if isinstance(value, (int, float)):
                    values.append(value)

        if not values:
            return f"No numerical data found for {metric} in {sheet_name}."

        avg_value = sum(values) / len(values)
        result = f"Average {metric} across all companies in {sheet_name}: {avg_value:.2f}"
        logging.debug(f"General metric result: {result}")
        return result

def main():
    processor = FinancialQueryProcessor("financial_data.xlsx")
    while True:
        query = input("Enter your query (or 'exit' to quit): ")
        if query.lower() == 'exit':
            break
        print(processor.process_query(query))

if __name__ == "__main__":
    main()
