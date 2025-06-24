#!/usr/bin/env python3
"""
Debug script to test specific PDF and Excel file combination for quantity mismatches.
"""

import sys
import os
from io import BytesIO

from server.services.excel_parser import ExcelParser
from server.services.pdf_parser import PDFParser
from server.services.matcher import Matcher


def debug_specific_files():
    """Test the specific file combination mentioned by user."""

    # Set UTF-8 encoding for output
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')

    # Specific files mentioned by user
    pdf_path = "../07_入札時（見積）積算参考資料.pdf"
    excel_path = "../水沢橋　積算書.xlsx"

    print("=== Testing Specific File Combination ===")
    print(f"PDF: {pdf_path}")
    print(f"Excel: {excel_path}")
    print()

    # Target items to focus on
    target_items = [
        "曲面加工",
        "現場孔明",
        "防護柵設置",
        "地覆補修用足場"
    ]

    try:
        # Extract from PDF
        print("1. Extracting from PDF...")
        pdf_parser = PDFParser()
        pdf_items = pdf_parser.extract_tables(pdf_path)
        print(f"   Found {len(pdf_items)} PDF items")

        # Extract from Excel
        print("2. Extracting from Excel...")
        excel_parser = ExcelParser()
        with open(excel_path, 'rb') as f:
            excel_items = excel_parser.extract_items_from_buffer_with_sheet(
                BytesIO(f.read()),
                sheet_name=None
            )
        print(f"   Found {len(excel_items)} Excel items")

        # Focus on target items in results
        print(f"\n3. Looking for target items in PDF:")
        pdf_target_items = []
        for item in pdf_items:
            for target in target_items:
                if target in item.item_key:
                    pdf_target_items.append(item)
                    print(
                        f"   Found: {item.item_key[:80]}... (qty: {item.quantity})")

        print(f"\n4. Looking for target items in Excel:")
        excel_target_items = []
        for item in excel_items:
            for target in target_items:
                if target in item.item_key:
                    excel_target_items.append(item)
                    print(
                        f"   Found: {item.item_key[:80]}... (qty: {item.quantity})")

        # Run matching
        print(f"\n5. Running matcher...")
        matcher = Matcher()
        comparison_summary = matcher.compare_items(pdf_items, excel_items)

        # Check results for target items
        print(f"\n6. Checking match results for target items:")

        # Check if target items are in mismatches
        print(f"\nQuantity mismatches involving target items:")
        mismatch_count = 0
        for result in comparison_summary.results:
            if result.status == "QUANTITY_MISMATCH":
                # Check if this involves any target item
                is_target = any(target in result.pdf_item.item_key or target in result.excel_item.item_key
                                for target in target_items)
                if is_target:
                    mismatch_count += 1
                    diff = result.quantity_difference
                    print(f"   MISMATCH: {result.pdf_item.item_key[:60]}...")
                    print(f"     PDF qty: {result.pdf_item.quantity}")
                    print(f"     Excel qty: {result.excel_item.quantity}")
                    print(f"     Difference: {diff}")
                    print()

        # Check if target items are missing in Excel
        print(f"Target items missing in Excel:")
        missing_count = 0
        for result in comparison_summary.results:
            if result.status == "MISSING":
                is_target = any(
                    target in result.pdf_item.item_key for target in target_items)
                if is_target:
                    missing_count += 1
                    print(
                        f"   MISSING: {result.pdf_item.item_key[:60]}... (qty: {result.pdf_item.quantity})")

        print(f"\n=== SUMMARY FOR TARGET ITEMS ===")
        print(f"Target quantity mismatches: {mismatch_count}")
        print(f"Target items missing in Excel: {missing_count}")

        if mismatch_count == 0 and missing_count == 0:
            print("✅ SUCCESS: No quantity mismatches or missing items for target items!")
        else:
            print("❌ ISSUE: Still have mismatches or missing items for target items")

        # Overall summary
        print(f"\n=== OVERALL SUMMARY ===")
        print(f"Total matched pairs: {comparison_summary.matched_items}")
        print(
            f"Total quantity mismatches: {comparison_summary.quantity_mismatches}")
        print(
            f"PDF items not found in Excel: {comparison_summary.missing_items}")
        print(
            f"Excel items not found in PDF: {comparison_summary.extra_items}")

    except FileNotFoundError as e:
        print(f"File not found: {e}")
        print("Please check that both files exist in the parent directory.")
    except Exception as e:
        print(f"Error during processing: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    debug_specific_files()
