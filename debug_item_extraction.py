#!/usr/bin/env python3
from server.services.pdf_parser import PDFParser
import sys
sys.path.append('.')


def debug_item_extraction():
    """Debug the _extract_subtable_item_improved function with actual data"""

    # Create parser instance
    pdf_parser = PDFParser()

    print('Testing _extract_subtable_item_improved function...')
    print()

    # Simulate the exact data from page 22
    # Based on our debug, the data rows are:
    data_rows = [
        # row 0 (original row 6)
        ['橋りょう世話役', None, '', '', '', '', '', '', ''],
        # row 1 (original row 7)
        [None, None, None, '', '', '', '', None, None],
        [None, None, None, '人', '', '', '', None,
            None],       # row 2 (original row 8)
        # row 3 (original row 9)
        ['橋りょう特殊工', None, '', '', '', '', '', '', ''],
        # row 4 (original row 10)
        [None, None, None, '', '', '', '', None, None],
        [None, None, None, '人', '', '', '', None,
            None],       # row 5 (original row 11)
        ['ｸｲｯｸﾃﾞｯｷ\n賃料,120日', None, '', '', '', '',
            '', '', ''],  # row 6 (original row 12)
        # row 7 (original row 13)
        [None, None, None, '', '', '', '', None, None],
        [None, None, None, '式', '1', '', '', None,
            None],      # row 8 (original row 14)
    ]

    # Column mapping (corrected)
    col_mapping = {
        '名称・規格': 0,
        '単位': 3,
        '数量': 4,
        '摘要': 8
    }

    print("Data rows:")
    for i, row in enumerate(data_rows):
        print(f"Row {i}: {row}")
    print()

    print("Column mapping:")
    for col_name, col_idx in col_mapping.items():
        print(f"  {col_name}: column {col_idx}")
    print()

    # Test extraction for row 3 (橋りょう特殊工)
    print("Testing extraction for row 3 (橋りょう特殊工):")
    result = pdf_parser._extract_subtable_item_improved(
        data_rows, 3, col_mapping, 21, '内7号', 9)

    if result:
        item = result['item']
        print(f"  Item key: {item.item_key}")
        print(f"  Quantity: {item.quantity}")
        print(f"  Unit: {item.unit}")
        print(f"  Raw fields: {item.raw_fields}")
        print(f"  Next index: {result['next_index']}")
    else:
        print("  No item extracted")
    print()

    # Test extraction for row 0 (橋りょう世話役)
    print("Testing extraction for row 0 (橋りょう世話役):")
    result = pdf_parser._extract_subtable_item_improved(
        data_rows, 0, col_mapping, 21, '内7号', 6)

    if result:
        item = result['item']
        print(f"  Item key: {item.item_key}")
        print(f"  Quantity: {item.quantity}")
        print(f"  Unit: {item.unit}")
        print(f"  Raw fields: {item.raw_fields}")
        print(f"  Next index: {result['next_index']}")
    else:
        print("  No item extracted")
    print()

    # Test what happens if we process all rows sequentially
    print("Processing all rows sequentially:")
    i = 0
    items = []
    while i < len(data_rows):
        result = pdf_parser._extract_subtable_item_improved(
            data_rows, i, col_mapping, 21, '内7号', 6 + i)

        if result:
            item = result['item']
            items.append(item)
            print(
                f"  Row {i}: {item.item_key} - qty={item.quantity}, unit='{item.unit}'")
            i = result['next_index']
        else:
            print(f"  Row {i}: No item extracted")
            i += 1

    print(f"\nTotal items extracted: {len(items)}")


if __name__ == "__main__":
    debug_item_extraction()
