#!/usr/bin/env python3
from server.services.pdf_parser import PDFParser
import sys
sys.path.append('.')

def test_specific_issue():
    """Test subtable extraction with specific page range (13-82) and verify no subtables from page 2"""

    # Test with the specific PDF file
    pdf_path = '../07_入札時（見積）積算参考資料.pdf'

    # Create parser instance
    pdf_parser = PDFParser()

    print('\nTest 1: Extract subtables without page range (should return empty list)')
    subtables_no_range = pdf_parser.extract_subtables_with_range(pdf_path)
    print(f'Total subtables found (no range): {len(subtables_no_range)}')
    assert len(subtables_no_range) == 0, "Should return empty list when no page range specified"

    print('\nTest 2: Extract subtables from page 2 only')
    subtables_page_2 = pdf_parser.extract_subtables_with_range(pdf_path, 2, 2)
    print(f'Subtables found on page 2: {len(subtables_page_2)}')
    assert len(subtables_page_2) == 0, "Should not find any subtables on page 2"

    print('\nTest 3: Extract subtables from correct range (13-82)')
    subtables_correct_range = pdf_parser.extract_subtables_with_range(pdf_path, 13, 82)
    print(f'Total subtables found in range 13-82: {len(subtables_correct_range)}')
    assert len(subtables_correct_range) > 0, "Should find subtables in the correct range"

    # Verify page numbers are within range
    page_numbers = set(item.page_number for item in subtables_correct_range)
    print(f'Found subtables on pages: {sorted(page_numbers)}')
    assert all(13 <= page <= 82 for page in page_numbers), "All subtables should be within specified range"

    # Print some sample subtables for verification
    print('\nSample subtables found:')
    for i, item in enumerate(subtables_correct_range[:5]):
        print(f'\nItem {i+1}:')
        print(f'  Page: {item.page_number}')
        print(f'  Reference: {item.reference_number}')
        print(f'  Name: {item.item_key}')
        print(f'  Quantity: {item.quantity}')
        print(f'  Unit: {item.unit}')

if __name__ == '__main__':
    test_specific_issue()
