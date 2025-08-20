#!/usr/bin/env python3
"""
Test script for Excel verification API
"""

from excel_verification_api import verify_excel_file
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def test_api():
    """Test the Excel verification API with the test file"""

    print("=" * 80)
    print("TESTING EXCEL VERIFICATION API")
    print("=" * 80)

    # Test file path
    test_file = "../【修正】水沢橋　積算書.xlsx"
    test_sheet = "52標準 15行本工事内訳書"

    if not os.path.exists(test_file):
        print(f"❌ Test file not found: {test_file}")
        return False

    print(f"✅ Test file found: {test_file}")
    print(f"📋 Test sheet: {test_sheet}")

    try:
        # Test the verification function
        result = verify_excel_file(test_file, test_sheet)

        print(f"\n📊 VERIFICATION RESULTS:")
        print(f"Extraction Successful: {result.extraction_successful}")
        print(f"Business Logic Verified: {result.business_logic_verified}")
        print(f"Total Items: {result.total_items}")
        print(f"Verified Items: {result.verified_items}")
        print(f"Mismatched Items: {result.mismatched_items}")

        if result.error_message:
            print(f"❌ Error: {result.error_message}")
            return False

        if result.mismatches:
            print(f"\n⚠️  MISMATCHES FOUND:")
            for i, mismatch in enumerate(result.mismatches, 1):
                print(f"   {i}. {mismatch['item_name']}")
                print(f"      Path: {mismatch['path']}")
                print(f"      Level: {mismatch['level']}")
                print(f"      Actual:   ¥{mismatch['amount']:>15,.0f}")
                print(f"      Expected: ¥{mismatch['children_sum']:>15,.0f}")
                print(f"      Diff:     ¥{mismatch['difference']:>15,.0f}")

        print(f"\n🎉 API TEST SUCCESSFUL!")
        return True

    except Exception as e:
        print(f"❌ API test failed: {e}")
        return False


if __name__ == "__main__":
    success = test_api()
    sys.exit(0 if success else 1)
