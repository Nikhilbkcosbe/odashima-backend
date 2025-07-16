#!/usr/bin/env python3
from server.services.pdf_parser import PDFParser
import sys
sys.path.append('.')


def debug_column_mapping():
    """Debug the column mapping for page 22 subtable"""

    pdf_path = '../07_入札時（見積）積算参考資料.pdf'

    # Create parser instance
    pdf_parser = PDFParser()

    print('Debugging column mapping for page 22...')
    print(f'PDF: {pdf_path}')
    print()

    # Let's manually test the column mapping logic
    # Based on our debug, the header row is row 5:
    # ['名称・規格', None, '条件', '単位', '数量', '単価', '金額', '数量・金額増減', '摘要']

    header_row = ['名称・規格', None, '条件', '単位', '数量', '単価', '金額', '数量・金額増減', '摘要']

    print("Header row from debug:")
    print(header_row)
    print()

    # Test the column mapping function
    col_mapping = pdf_parser._get_subtable_column_mapping_improved(header_row)
    print("Column mapping:")
    for col_name, col_idx in col_mapping.items():
        print(f"  {col_name}: column {col_idx}")
    print()

    # Now let's test what happens when we extract from the problematic rows
    print("Testing extraction from problematic rows:")
    print()

    # Row 9: ['橋りょう特殊工', None, '', '', '', '', '', '', '']
    row_9 = ['橋りょう特殊工', None, '', '', '', '', '', '', '']
    print("Row 9 (橋りょう特殊工):", row_9)

    # Extract data manually
    raw_fields = {}
    quantity = 0.0
    unit = None
    item_name = None

    for col_name, col_idx in col_mapping.items():
        if col_idx < len(row_9) and row_9[col_idx]:
            cell_value = str(row_9[col_idx]).strip()
            if cell_value:
                print(
                    f"  Found data in {col_name} (col {col_idx}): '{cell_value}'")
                if col_name == "名称・規格":
                    item_name = cell_value
                    raw_fields[col_name] = cell_value
                elif col_name == "数量":
                    quantity = pdf_parser._extract_quantity(cell_value)
                    raw_fields[col_name] = cell_value
                elif col_name == "単位":
                    unit = cell_value
                    raw_fields[col_name] = cell_value

    print(
        f"  Extracted: item_name='{item_name}', quantity={quantity}, unit='{unit}'")
    print(f"  Raw fields: {raw_fields}")
    print()

    # Row 11: [None, None, None, '人', '', '', '', None, None]
    row_11 = [None, None, None, '人', '', '', '', None, None]
    print("Row 11 (unit row):", row_11)

    for col_name, col_idx in col_mapping.items():
        if col_idx < len(row_11) and row_11[col_idx]:
            cell_value = str(row_11[col_idx]).strip()
            if cell_value:
                print(
                    f"  Found data in {col_name} (col {col_idx}): '{cell_value}'")
    print()

    # Row 14: [None, None, None, '式', '1', '', '', None, None]
    row_14 = [None, None, None, '式', '1', '', '', None, None]
    print("Row 14 (wrong data):", row_14)

    for col_name, col_idx in col_mapping.items():
        if col_idx < len(row_14) and row_14[col_idx]:
            cell_value = str(row_14[col_idx]).strip()
            if cell_value:
                print(
                    f"  Found data in {col_name} (col {col_idx}): '{cell_value}'")


if __name__ == "__main__":
    debug_column_mapping()
