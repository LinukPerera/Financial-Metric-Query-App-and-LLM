import openpyxl
from openpyxl.utils import get_column_letter

def get_merged_parent(sheet, row, col):
    """Return the merged value and its full bounds if cell is in a merged range."""
    for merged_range in sheet.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        if min_row <= row <= max_row and min_col <= col <= max_col:
            value = sheet.cell(row=min_row, column=min_col).value
            return value, (min_row, max_row), (min_col, max_col)
    # Not merged
    value = sheet.cell(row=row, column=col).value
    return value, (row, row), (col, col)

def detect_first_nonempty_row(sheet, max_scan=10):
    """Find the first row that contains at least one non-empty cell."""
    for i in range(1, max_scan + 1):
        row = sheet[i]
        if any(cell.value not in (None, '') for cell in row):
            return i
    return 1  # fallback

def count_filled_header_cells(sheet, start_row, num_rows=2):
    """Count non-blank cells from start_row across `num_rows`."""
    filled = set()
    for row in range(start_row, start_row + num_rows):
        if row > sheet.max_row:
            continue
        for col in range(1, sheet.max_column + 1):
            val = sheet.cell(row=row, column=col).value
            if val not in (None, ''):
                filled.add((row, col))
    return len(filled)

def detect_multirow_sheet(wb):
    """Return the sheet with the most populated 2-row header area from first non-blank row."""
    counts = {}
    for name in wb.sheetnames[:3]:
        sheet = wb[name]
        start = detect_first_nonempty_row(sheet)
        counts[name] = count_filled_header_cells(sheet, start, num_rows=2)
    return max(counts, key=counts.get)

def print_sheet_headers(sheet, multirow_enabled):
    print(f"\n--- Headers from sheet: {sheet.title} ---")

    base_row = detect_first_nonempty_row(sheet)
    header_rows = [base_row]
    if multirow_enabled:
        header_rows.append(base_row + 1)

    for col in range(1, sheet.max_column + 1):
        hierarchy = []
        for row in header_rows:
            if row > sheet.max_row:
                continue
            value, (row_start, row_end), (col_start, col_end) = get_merged_parent(sheet, row, col)
            if value not in (None, ''):
                label = str(value).strip()
                if label not in hierarchy:
                    hierarchy.append(label)

        col_letter = get_column_letter(col)
        if hierarchy:
            print(f"{col_letter}: {' > '.join(hierarchy)}")
        else:
            print(f"{col_letter}: [Blank]")

def main():
    file_path = "financial_data.xlsx"  # Replace with your file path
    wb = openpyxl.load_workbook(file_path, data_only=True)

    multirow_sheet = detect_multirow_sheet(wb)

    for sheet_name in wb.sheetnames[:3]:
        sheet = wb[sheet_name]
        is_multirow = (sheet_name == multirow_sheet)
        print_sheet_headers(sheet, multirow_enabled=is_multirow)

if __name__ == "__main__":
    main()