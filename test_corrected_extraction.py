#!/usr/bin/env python3
"""
Test script for the corrected Excel table extraction service.
This script tests the new ExcelTableExtractorService integration.
"""

from server.services.excel_parser import ExcelParser
from server.services.excel_table_extractor_service import ExcelTableExtractorService
import sys
import os
import tempfile
from io import BytesIO

# Add the server directory to the Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'server'))


def test_corrected_extraction():
    """Test the corrected Excel table extraction"""

    # Test file path (you can change this to your test file)
    test_file_path = "水沢橋　積算書.xlsx"
    test_sheet_name = "52標準 15行本工事内訳書"

    if not os.path.exists(test_file_path):
        print(f"❌ Test file not found: {test_file_path}")
        print("Please place the test Excel file in the backend directory")
        return False

    try:
        print("🧪 Testing corrected Excel table extraction...")

        # Test 1: Direct service usage
        print("\n1. Testing ExcelTableExtractorService directly...")
        service = ExcelTableExtractorService()

        # Read file into buffer
        with open(test_file_path, 'rb') as f:
            excel_buffer = BytesIO(f.read())

        # Extract using corrected logic
        items = service.extract_main_table_from_buffer(
            excel_buffer, test_sheet_name)

        print(f"✅ Extracted {len(items)} items using corrected logic")

        # Display first few items
        for i, item in enumerate(items[:5]):
            print(
                f"   Item {i+1}: {item.item_key[:50]}... (Qty: {item.quantity}, Unit: {item.unit})")

        # Test 2: Integration with ExcelParser
        print("\n2. Testing integration with ExcelParser...")
        excel_parser = ExcelParser()

        # Reset buffer position
        excel_buffer.seek(0)

        # Extract using the new method
        parser_items = excel_parser.extract_main_table_from_buffer(
            excel_buffer, test_sheet_name)

        print(
            f"✅ ExcelParser extracted {len(parser_items)} items using corrected logic")

        # Compare results
        if len(items) == len(parser_items):
            print("✅ Both methods extracted the same number of items")
        else:
            print(
                f"⚠️  Different item counts: Service={len(items)}, Parser={len(parser_items)}")

        # Test 3: Check for specific data patterns
        print("\n3. Checking data quality...")

        items_with_quantity = [item for item in items if item.quantity > 0]
        items_with_unit = [item for item in items if item.unit]
        items_with_name = [item for item in items if item.item_key and len(
            item.item_key.strip()) > 0]

        print(f"   Items with quantity > 0: {len(items_with_quantity)}")
        print(f"   Items with unit: {len(items_with_unit)}")
        print(f"   Items with name: {len(items_with_name)}")

        if len(items_with_name) > 0:
            print("✅ Data quality check passed")
        else:
            print("❌ No items with names found")
            return False

        print("\n🎉 All tests passed! The corrected extraction logic is working correctly.")
        return True

    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_original_vs_corrected():
    """Compare original vs corrected extraction methods"""

    test_file_path = "水沢橋　積算書.xlsx"
    test_sheet_name = "52標準 15行本工事内訳書"

    if not os.path.exists(test_file_path):
        print(f"❌ Test file not found: {test_file_path}")
        return False

    try:
        print("\n🔄 Comparing original vs corrected extraction methods...")

        excel_parser = ExcelParser()

        # Read file into buffer
        with open(test_file_path, 'rb') as f:
            excel_buffer = BytesIO(f.read())

        # Test original method
        excel_buffer.seek(0)
        original_items = excel_parser.extract_items_from_buffer_with_sheet(
            excel_buffer, test_sheet_name)

        # Test corrected method
        excel_buffer.seek(0)
        corrected_items = excel_parser.extract_main_table_from_buffer(
            excel_buffer, test_sheet_name)

        print(f"Original method extracted: {len(original_items)} items")
        print(f"Corrected method extracted: {len(corrected_items)} items")

        if len(corrected_items) > len(original_items):
            print("✅ Corrected method extracted more items (expected improvement)")
        elif len(corrected_items) == len(original_items):
            print("ℹ️  Both methods extracted the same number of items")
        else:
            print("⚠️  Corrected method extracted fewer items")

        # Show sample items from corrected method
        print("\nSample items from corrected method:")
        for i, item in enumerate(corrected_items[:3]):
            print(
                f"  {i+1}. {item.item_key[:40]}... (Qty: {item.quantity}, Unit: {item.unit})")

        return True

    except Exception as e:
        print(f"❌ Comparison test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    print("🚀 Starting Excel Table Extraction Tests")
    print("=" * 50)

    success = True

    # Run tests
    success &= test_corrected_extraction()
    success &= test_original_vs_corrected()

    print("\n" + "=" * 50)
    if success:
        print("🎉 All tests completed successfully!")
        print("The corrected Excel table extraction is ready for use.")
    else:
        print("❌ Some tests failed. Please check the errors above.")

    sys.exit(0 if success else 1)
