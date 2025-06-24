#!/usr/bin/env python3
"""
Debug script to compare PDF and Excel results and identify matching issues.
"""

import pandas as pd
import re
from typing import List, Dict, Optional, Tuple
from io import BytesIO

from server.services.excel_parser import ExcelParser
from server.services.pdf_parser import PDFParser
from server.services.matcher import Matcher
from server.services.normalizer import Normalizer
from server.schemas.tender import TenderItem


def debug_pdf_excel_matching():
    """Debug PDF vs Excel matching for the specific files."""

    # File paths
    pdf_path = "../sample.pdf"  # This should be 07_入札時（見積）積算参考資料.pdf
    excel_path = "../07_入札時（見積）積算参考資料.xlsx"

    print(f"=== Debugging PDF vs Excel Matching ===\n")
    print(f"PDF: {pdf_path}")
    print(f"Excel: {excel_path}\n")

    # Extract items from both files
    print("1. Extracting items from PDF...")
    pdf_parser = PDFParser()
    pdf_items = pdf_parser.extract_tables(pdf_path)

    print(f"Found {len(pdf_items)} PDF items\n")

    print("2. Extracting items from Excel...")
    excel_parser = ExcelParser()
    with open(excel_path, 'rb') as f:
        excel_items = excel_parser.extract_items_from_buffer_with_sheet(
            BytesIO(f.read()), sheet_name=None, item_name_column=None
        )

    print(f"Found {len(excel_items)} Excel items\n")

    # Target items to debug
    target_items = [
        "曲面加工",
        "現場孔明",
        "防護柵設置",
        "地覆補修用足場"
    ]

    print("3. Finding target items in PDF...")
    pdf_targets = []
    for item in pdf_items:
        for target in target_items:
            if target in item.item_key:
                pdf_targets.append((target, item))
                print(
                    f"PDF: '{target}' -> '{item.item_key}' (quantity: {item.quantity})")
                break

    print(f"\n4. Finding target items in Excel...")
    excel_targets = []
    for item in excel_items:
        for target in target_items:
            if target in item.item_key:
                excel_targets.append((target, item))
                print(
                    f"Excel: '{target}' -> '{item.item_key}' (quantity: {item.quantity})")
                break

    print(f"\n5. Testing matching logic...")
    matcher = Matcher()
    normalizer = Normalizer()

    # For each target item, try to match PDF with Excel
    for pdf_target, pdf_item in pdf_targets:
        print(
            f"\n--- Matching PDF item: '{pdf_item.item_key}' (qty: {pdf_item.quantity}) ---")

        # Find potential Excel matches
        potential_matches = []
        for excel_target, excel_item in excel_targets:
            if pdf_target in excel_target or excel_target in pdf_target:
                potential_matches.append((excel_target, excel_item))

        if not potential_matches:
            print(f"No potential Excel matches found for '{pdf_target}'")
            continue

        for excel_target, excel_item in potential_matches:
            print(
                f"\nTesting match with Excel: '{excel_item.item_key}' (qty: {excel_item.quantity})")

            # Test fuzzy matching
            similarity = matcher._calculate_similarity(
                pdf_item.item_key, excel_item.item_key)
            print(f"Similarity score: {similarity:.3f}")

            # Test if they are significantly different
            are_different = normalizer.are_items_significantly_different(
                pdf_item.item_key, excel_item.item_key
            )
            print(f"Are significantly different: {are_different}")

            # Test quantity comparison
            qty_diff = abs(pdf_item.quantity - excel_item.quantity)
            qty_ratio = max(pdf_item.quantity, excel_item.quantity) / min(pdf_item.quantity,
                                                                          excel_item.quantity) if min(pdf_item.quantity, excel_item.quantity) > 0 else float('inf')
            print(f"Quantity difference: {qty_diff}")
            print(f"Quantity ratio: {qty_ratio:.3f}")

            # Test normalized text comparison
            pdf_normalized = normalizer.normalize_text(pdf_item.item_key)
            excel_normalized = normalizer.normalize_text(excel_item.item_key)
            print(f"PDF normalized: '{pdf_normalized}'")
            print(f"Excel normalized: '{excel_normalized}'")

            # Test final matching decision
            match_result = matcher._is_good_match(
                pdf_item.item_key, excel_item.item_key, similarity)
            print(f"Would be matched: {match_result}")

    print(f"\n6. Full matching test...")
    # Run full matching process
    matched_items, pdf_only, excel_only = matcher.match_items(
        pdf_items, excel_items)

    print(f"Matched items: {len(matched_items)}")
    print(f"PDF only: {len(pdf_only)}")
    print(f"Excel only: {len(excel_only)}")

    # Check if our target items are in the mismatched lists
    print(f"\n7. Checking target items in mismatch lists...")

    target_keywords = ["曲面加工", "現場孔明", "防護柵設置", "地覆補修用足場"]

    print("Target items in PDF-only list:")
    for item in pdf_only:
        for keyword in target_keywords:
            if keyword in item.item_key:
                print(f"  PDF-only: '{item.item_key}' (qty: {item.quantity})")
                break

    print("Target items in Excel-only list:")
    for item in excel_only:
        for keyword in target_keywords:
            if keyword in item.item_key:
                print(
                    f"  Excel-only: '{item.item_key}' (qty: {item.quantity})")
                break

    print("Target items in matched list:")
    for pdf_item, excel_item in matched_items:
        for keyword in target_keywords:
            if keyword in pdf_item.item_key or keyword in excel_item.item_key:
                qty_match = "✅" if abs(
                    pdf_item.quantity - excel_item.quantity) < 0.01 else "❌"
                print(
                    f"  {qty_match} PDF: '{pdf_item.item_key}' (qty: {pdf_item.quantity})")
                print(
                    f"      Excel: '{excel_item.item_key}' (qty: {excel_item.quantity})")
                break


def debug_specific_item_matching():
    """Debug specific item matching with detailed analysis."""

    # Test the problematic items directly
    test_cases = [
        ("曲面加工", "曲面加工 + R=2mm"),
        ("現場孔明", "現場孔明"),
        ("防護柵設置", "防護柵設置 + 歩行者自転車柵兼用,B種,H=950mm"),
        ("地覆補修用足場", "地覆補修用足場")
    ]

    print(f"\n=== Testing Specific Item Matching ===\n")

    matcher = Matcher()
    normalizer = Normalizer()

    for pdf_name, excel_name in test_cases:
        print(f"Testing: PDF='{pdf_name}' vs Excel='{excel_name}'")

        # Test similarity
        similarity = matcher._calculate_similarity(pdf_name, excel_name)
        print(f"  Similarity: {similarity:.3f}")

        # Test if significantly different
        are_different = normalizer.are_items_significantly_different(
            pdf_name, excel_name)
        print(f"  Significantly different: {are_different}")

        # Test normalization
        pdf_norm = normalizer.normalize_text(pdf_name)
        excel_norm = normalizer.normalize_text(excel_name)
        print(f"  PDF normalized: '{pdf_norm}'")
        print(f"  Excel normalized: '{excel_norm}'")

        # Test final match decision
        is_match = matcher._is_good_match(pdf_name, excel_name, similarity)
        print(f"  Would match: {is_match}")
        print()


if __name__ == "__main__":
    import os
    if not os.path.exists("../sample.pdf"):
        print("PDF file not found: ../sample.pdf")
    elif not os.path.exists("../07_入札時（見積）積算参考資料.xlsx"):
        print("Excel file not found: ../07_入札時（見積）積算参考資料.xlsx")
    else:
        debug_pdf_excel_matching()
        debug_specific_item_matching()
