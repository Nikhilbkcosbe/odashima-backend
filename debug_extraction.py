#!/usr/bin/env python3
"""
Debug script to compare the corrected extraction with the original method.
"""

from server.services.excel_parser import ExcelParser
from server.services.excel_table_extractor_service import ExcelTableExtractorService
import sys
import os
from io import BytesIO

# Add the server directory to the Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'server'))


def debug_extraction():
    """Debug the extraction process"""

    # Test file path
    test_file_path = "水沢橋　積算書.xlsx"
    test_sheet_name = "52標準 15行本工事内訳書"

    if not os.path.exists(test_file_path):
        print(f"❌ Test file not found: {test_file_path}")
        return

    print(f"🔍 Testing extraction for sheet: {test_sheet_name}")

    # Test 1: Original method
    print("\n" + "="*80)
    print("TEST 1: ORIGINAL METHOD")
    print("="*80)

    try:
        with open(test_file_path, 'rb') as f:
            excel_buffer = BytesIO(f.read())

        excel_parser = ExcelParser()
        original_items = excel_parser.extract_items_from_buffer_with_sheet(
            excel_buffer, test_sheet_name)

        print(f"✅ Original method extracted: {len(original_items)} items")

        # Show first few items
        for i, item in enumerate(original_items[:5]):
            print(f"  {i+1}. {item.item_key} | {item.quantity} {item.unit}")

    except Exception as e:
        print(f"❌ Original method failed: {e}")
        import traceback
        traceback.print_exc()

    # Test 2: Corrected method
    print("\n" + "="*80)
    print("TEST 2: CORRECTED METHOD")
    print("="*80)

    try:
        with open(test_file_path, 'rb') as f:
            excel_buffer = BytesIO(f.read())

        service = ExcelTableExtractorService()
        corrected_items = service.extract_main_table_from_buffer(
            excel_buffer, test_sheet_name)

        print(f"✅ Corrected method extracted: {len(corrected_items)} items")

        # Show first few items
        for i, item in enumerate(corrected_items[:5]):
            print(f"  {i+1}. {item.item_key} | {item.quantity} {item.unit}")

    except Exception as e:
        print(f"❌ Corrected method failed: {e}")
        import traceback
        traceback.print_exc()

    # Test 3: Direct corrected extractor
    print("\n" + "="*80)
    print("TEST 3: DIRECT CORRECTED EXTRACTOR")
    print("="*80)

    try:
        from server.services.excel_table_extractor_service import ExcelTableExtractorCorrected

        extractor = ExcelTableExtractorCorrected(
            test_file_path, test_sheet_name)
        tables = extractor.extract_all_tables()

        print(f"✅ Direct extractor found: {len(tables)} tables")

        total_data_rows = 0
        for i, table in enumerate(tables):
            data_rows = len(table.get('data_rows', []))
            total_data_rows += data_rows
            print(f"  Table {i+1}: {data_rows} data rows")

            # Show first few data rows from first table
            if i == 0 and table.get('data_rows'):
                print("    First few data rows:")
                for j, row in enumerate(table['data_rows'][:3]):
                    print(f"      Row {j+1}: {row}")

        print(f"✅ Total data rows across all tables: {total_data_rows}")

    except Exception as e:
        print(f"❌ Direct extractor failed: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    debug_extraction()
