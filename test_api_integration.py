#!/usr/bin/env python3
"""
Test script to verify API integration with the new standalone Excel extractor.
"""

import sys
import os
from io import BytesIO

# Add the server directory to the Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'server'))

from server.services.excel_table_extractor_service import ExcelTableExtractorService

def test_api_integration():
    """Test the API integration with the new standalone extractor"""
    
    # Test file path
    test_file_path = "水沢橋　積算書.xlsx"
    test_sheet_name = "52標準 15行本工事内訳書"
    
    if not os.path.exists(test_file_path):
        print(f"❌ Test file not found: {test_file_path}")
        return
    
    print(f"🔍 Testing API integration for sheet: {test_sheet_name}")
    print("=" * 80)
    
    try:
        # Read the Excel file into a buffer (simulating API upload)
        with open(test_file_path, 'rb') as f:
            excel_content = f.read()
        
        excel_buffer = BytesIO(excel_content)
        
        # Create the service instance (as the API would)
        excel_table_extractor = ExcelTableExtractorService()
        
        # Extract using the standalone logic (as the API would)
        tender_items = excel_table_extractor.extract_main_table_from_buffer(
            excel_buffer, test_sheet_name)
        
        print(f"✅ API integration successful!")
        print(f"📊 Extracted {len(tender_items)} items")
        
        # Show some sample items
        print("\n📋 Sample items:")
        for i, item in enumerate(tender_items[:5]):
            print(f"  {i+1}. {item.item_key} | {item.quantity} {item.unit}")
        
        # Check for the specific item mentioned in the user query
        target_item = "補強部材取付工 1部材当り平均質量G≦20kg"
        found_item = None
        
        for item in tender_items:
            if target_item in item.item_key:
                found_item = item
                break
        
        if found_item:
            print(f"\n🎯 Found target item: '{found_item.item_key}'")
            print(f"   Unit: '{found_item.unit}'")
            print(f"   Quantity: {found_item.quantity}")
        else:
            print(f"\n❌ Target item not found: {target_item}")
        
        # Close buffer
        excel_buffer.close()
        
        print("\n✅ API integration test completed successfully!")
        
    except Exception as e:
        print(f"❌ API integration test failed: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_api_integration() 