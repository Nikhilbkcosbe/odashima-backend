#!/usr/bin/env python3
"""
Debug script for Excel subtable extraction - specifically for 水沢橋　積算書.xlsx
Focus on reference number "57" subtable to check row spanning logic.
"""

from server.services.excel_parser import ExcelParser
import pandas as pd
import sys
import os
from io import BytesIO

# Add the backend directory to Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'backend'))


def debug_excel_subtable():
    """Debug the Excel subtable extraction for the specific file and reference."""

    # Path to the Excel file
    excel_path = "水沢橋　積算書.xlsx"

    if not os.path.exists(excel_path):
        print(f"Error: File {excel_path} not found in current directory")
        return

    print(f"=== Debugging Excel Subtable Extraction ===")
    print(f"File: {excel_path}")
    print(f"Target: Reference number '57' subtable")
    print()

    # Read the Excel file
    with open(excel_path, 'rb') as f:
        excel_buffer = BytesIO(f.read())

    # Initialize parser
    parser = ExcelParser()

    # Extract subtables using the simple extractor
    print("Extracting subtables...")
    try:
        # Assuming main sheet name - you may need to adjust this
        main_sheet_name = "積算書"  # Common name, adjust if needed

        # Get all sheet names first
        xl = pd.ExcelFile(excel_buffer)
        sheet_names = xl.sheet_names
        print(f"Available sheets: {sheet_names}")

        # Try to identify main sheet
        main_sheet = sheet_names[0]  # Default to first sheet
        for sheet in sheet_names:
            if any(keyword in sheet for keyword in ['積算', '標準', '本工事', 'メイン']):
                main_sheet = sheet
                break

        print(f"Using main sheet: {main_sheet}")

        # Reset buffer
        excel_buffer.seek(0)

        # Extract subtables
        subtables = parser.extract_subtables_from_buffer(
            excel_buffer, main_sheet_name=main_sheet)

        print(f"Total subtables extracted: {len(subtables)}")
        print()

        # Find reference number "57" subtable
        ref_57_items = [
            item for item in subtables if item.reference_number and "57" in item.reference_number]

        if not ref_57_items:
            print("No items found for reference number containing '57'")
            print("Available reference numbers:")
            unique_refs = set(
                item.reference_number for item in subtables if item.reference_number)
            for ref in sorted(unique_refs):
                print(f"  {ref}")
        else:
            print(
                f"Found {len(ref_57_items)} items for reference number containing '57':")
            print()

            for i, item in enumerate(ref_57_items, 1):
                print(f"Item {i}:")
                print(f"  Reference: {item.reference_number}")
                print(f"  Item Key: {item.item_key}")
                print(f"  Quantity: {item.quantity}")
                print(f"  Unit: {item.unit}")
                print(f"  Sheet: {item.sheet_name}")
                print(f"  Raw Fields: {item.raw_fields}")
                print()

        # Also show some context around reference 57
        print("=== Context Analysis ===")
        all_refs = [
            item.reference_number for item in subtables if item.reference_number]
        unique_refs = sorted(set(all_refs))

        if len(unique_refs) > 0:
            print(f"All reference numbers found: {unique_refs}")

            # Look for references close to 57
            target_refs = [ref for ref in unique_refs if any(
                num in ref for num in ['56', '57', '58'])]
            if target_refs:
                print(f"References near 57: {target_refs}")

                for ref in target_refs:
                    items = [
                        item for item in subtables if item.reference_number == ref]
                    print(f"\nReference {ref}: {len(items)} items")
                    for item in items[:3]:  # Show first 3 items
                        print(
                            f"  - {item.item_key} (qty: {item.quantity}, unit: {item.unit})")

    except Exception as e:
        print(f"Error during extraction: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    debug_excel_subtable()
