#!/usr/bin/env python3
"""
Debug script to analyze quantity extraction issues for specific items in Excel file.
"""

import pandas as pd
import re
from typing import List, Dict, Optional
from io import BytesIO

from server.services.excel_parser import ExcelParser
from server.schemas.tender import TenderItem


def analyze_excel_file(excel_path: str):
    """Analyze the Excel file to understand quantity extraction issues."""

    print(f"=== Analyzing Excel file: {excel_path} ===\n")

    # Initialize parser
    parser = ExcelParser()

    # Extract items using the current parser
    print("1. Extracting items using current parser...")
    with open(excel_path, 'rb') as f:
        items = parser.extract_items_from_buffer_with_sheet(
            BytesIO(f.read()),
            sheet_name=None,
            item_name_column=None
        )

    print(f"Found {len(items)} total items\n")

    # Target items to debug
    target_items = [
        "曲面加工 + R=2mm",
        "現場孔明",
        "防護柵設置 + 歩行者自転車柵兼用,B種,H=950mm",
        "地覆補修用足場"
    ]

    print("2. Searching for target items...")
    found_items = []

    for item in items:
        for target in target_items:
            if target in item.item_key or any(target_part in item.item_key for target_part in target.split(" + ")):
                found_items.append((target, item))
                print(
                    f"FOUND: '{target}' -> '{item.item_key}' (quantity: {item.quantity})")
                break

    if not found_items:
        print("No exact matches found. Searching for partial matches...")

        for item in items:
            for target in target_items:
                # Check for partial matches
                target_parts = target.replace(" + ", " ").split()
                if any(part in item.item_key for part in target_parts if len(part) > 2):
                    print(
                        f"PARTIAL: '{target}' -> '{item.item_key}' (quantity: {item.quantity})")

    print("\n3. Analyzing raw Excel structure...")

    # Read Excel file directly to analyze structure
    excel_file = pd.ExcelFile(excel_path)
    for sheet_name in excel_file.sheet_names:
        print(f"\n--- Sheet: {sheet_name} ---")
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

        # Search for target items in raw data
        for target in target_items:
            target_clean = target.replace(" + ", "").replace(",", "")

            for row_idx, row in df.iterrows():
                row_str = " ".join([str(val) for val in row if pd.notna(val)])

                if target_clean in row_str or any(part in row_str for part in target_clean.split() if len(part) > 2):
                    print(f"Found '{target}' in row {row_idx + 1}:")
                    print(
                        f"  Raw row data: {[str(val) if pd.notna(val) else 'NaN' for val in row[:10]]}")

                    # Check if there's quantity data in this row or nearby rows
                    print(f"  Checking for quantity in current row...")
                    for col_idx, val in enumerate(row):
                        if pd.notna(val):
                            val_str = str(val).replace(",", "").strip()
                            if re.match(r'^\d+\.?\d*$', val_str):
                                print(
                                    f"    Possible quantity in column {col_idx}: {val}")

                    # Check next few rows for quantity data (row spanning)
                    print(f"  Checking next 3 rows for quantity (row spanning)...")
                    for next_row_offset in range(1, 4):
                        if row_idx + next_row_offset < len(df):
                            next_row = df.iloc[row_idx + next_row_offset]
                            next_row_str = " ".join(
                                [str(val) for val in next_row if pd.notna(val)])

                            # Look for numeric values
                            numeric_values = []
                            for col_idx, val in enumerate(next_row):
                                if pd.notna(val):
                                    val_str = str(val).replace(",", "").strip()
                                    if re.match(r'^\d+\.?\d*$', val_str) and float(val_str) > 0:
                                        numeric_values.append(
                                            f"col{col_idx}:{val}")

                            if numeric_values:
                                print(
                                    f"    Row {row_idx + next_row_offset + 1}: {numeric_values}")

                    print()

    print("\n4. Analyzing extracted items with zero quantities...")
    zero_quantity_items = [item for item in items if item.quantity == 0]
    print(f"Found {len(zero_quantity_items)} items with zero quantity")

    for item in zero_quantity_items[:10]:  # Show first 10
        print(f"  '{item.item_key}' - Raw fields: {item.raw_fields}")

    print("\n5. Show all extracted items for reference...")
    print("All extracted items:")
    for i, item in enumerate(items):
        print(f"  {i+1:3d}: '{item.item_key}' (quantity: {item.quantity})")
        if i >= 20:  # Show first 20 items
            print(f"  ... and {len(items) - 20} more items")
            break


def debug_quantity_extraction():
    """Debug quantity extraction for the specific Excel file."""

    excel_path = "../07_入札時（見積）積算参考資料.xlsx"

    import os
    if not os.path.exists(excel_path):
        print(f"Excel file not found: {excel_path}")
        return

    analyze_excel_file(excel_path)


if __name__ == "__main__":
    debug_quantity_extraction()
