#!/usr/bin/env python3
from server.services.pdf_parser import PDFParser
import sys
sys.path.append('.')


def test_specific_issue():
    """Test the specific issue with page 22 and reference 内7号"""

    # Test with the specific PDF file
    pdf_path = '../07_入札時（見積）積算参考資料.pdf'

    # Create parser instance
    pdf_parser = PDFParser()

    print('Testing PDF subtable extraction with specific parameters...')
    print(f'PDF: {pdf_path}')
    print('Page: 22')
    print('Reference: 内7号')
    print()

    # Extract subtables from PDF using the correct method
    pdf_subtables = pdf_parser.extract_subtables_with_range(pdf_path)
    print(f'Total PDF subtables found: {len(pdf_subtables)}')

    # Filter for page 22 and reference 内7号
    target_subtables = []
    for subtable in pdf_subtables:
        if subtable.page_number == 22 and subtable.reference_number == '内7号':
            target_subtables.append(subtable)

    print(f'Subtables on page 22 with reference 内7号: {len(target_subtables)}')

    # Look for the specific items
    problem_items = []
    print(f'Items in subtable: {len(target_subtables)}')

    for item in target_subtables:
        item_key = item.item_key
        print(f'Item: {item_key}')
        print(f'  Quantity: {item.quantity}')
        print(f'  Unit: {item.unit}')
        print(f'  Raw fields: {item.raw_fields}')
        print()

        if '橋りょう特殊工' in item_key or '橋梁点検車' in item_key:
            print(f'*** FOUND TARGET ITEM: {item_key} ***')
            print(f'  Quantity: {item.quantity}')
            print(f'  Unit: {item.unit}')
            print(f'  Raw fields: {item.raw_fields}')
            print()

            # Check if this is a problem item
            quantity = item.quantity
            unit = item.unit

            if quantity == 1 and unit == '式':
                problem_items.append({
                    'item_key': item_key,
                    'quantity': quantity,
                    'unit': unit,
                    'expected_quantity': 0,
                    'expected_unit': '人' if '橋りょう特殊工' in item_key else '日'
                })

    print(f"Problem items found: {len(problem_items)}")
    for item in problem_items:
        print(f"  {item['item_key']}: qty={item['quantity']}, unit={item['unit']} (should be qty={item['expected_quantity']}, unit={item['expected_unit']})")

    return problem_items


if __name__ == "__main__":
    test_specific_issue()
