#!/usr/bin/env python3
"""
Debug script to examine Excel row spanning behavior for specific items.
"""

import sys
import os
from io import BytesIO

from server.services.excel_parser import ExcelParser


def debug_row_spanning():
    """Debug row spanning behavior for specific items."""

    # Set UTF-8 encoding for output
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')

    excel_path = "../水沢橋　積算書.xlsx"

    print("=== Excel Row Spanning Debug ===")
    print(f"File: {excel_path}")
    print()

    # Target items to focus on
    target_items = [
        "防護柵設置",
        "現場孔明",
        "曲面加工",
        "地覆補修用足場"
    ]

    try:
        # Extract from Excel with detailed logging
        print("Extracting from Excel with row spanning analysis...")
        excel_parser = ExcelParser()

        # Enable more detailed logging for row spanning
        import logging
        logging.getLogger(
            'server.services.excel_parser').setLevel(logging.DEBUG)

        with open(excel_path, 'rb') as f:
            excel_items = excel_parser.extract_items_from_buffer_with_sheet(
                BytesIO(f.read()),
                sheet_name="52標準 15行本工事内訳書"  # Focus on specific sheet
            )

        print(f"Total Excel items extracted: {len(excel_items)}")
        print()

        # Analyze target items
        for target in target_items:
            print(f"=== Analysis for '{target}' ===")

            # Find all items containing this target
            matching_items = []
            for item in excel_items:
                if target in item.item_key:
                    matching_items.append(item)

            if not matching_items:
                print(f"   No items found containing '{target}'")
                continue

            print(f"   Found {len(matching_items)} items:")
            for item in matching_items:
                print(f"     • {item.item_key} → quantity: {item.quantity}")

                # Check if this looks like a merged item (contains multiple components)
                if " + " in item.item_key:
                    base_name = item.item_key.split(" + ")[0]
                    specification = " + ".join(item.item_key.split(" + ")[1:])
                    print(f"       Base: '{base_name}'")
                    print(f"       Spec: '{specification}'")
                    print(f"       This appears to be a row-spanning result")

            print()

        # Now let's manually examine the raw Excel data for these items
        print("=== Raw Excel Sheet Analysis ===")
        print("Let's examine the actual Excel rows to understand the row spanning...")

        # Read the Excel file directly to see raw data
        import pandas as pd
        excel_file = pd.ExcelFile(excel_path)
        df = pd.read_excel(
            excel_file, sheet_name="52標準 15行本工事内訳書", header=None)

        # Search for target items in the raw data
        for target in target_items:
            print(f"\n--- Raw data search for '{target}' ---")

            found_rows = []
            for idx, row in df.iterrows():
                row_str = " ".join([str(val) for val in row if pd.notna(val)])
                if target in row_str:
                    found_rows.append((idx, row, row_str))

            if found_rows:
                print(f"Found {len(found_rows)} rows containing '{target}':")
                for idx, row, row_str in found_rows:
                    print(f"  Row {idx + 1}: {row_str}")

                    # Look at the next few rows to see if they contain quantities
                    for next_idx in range(idx + 1, min(idx + 4, len(df))):
                        next_row = df.iloc[next_idx]
                        next_str = " ".join(
                            [str(val) for val in next_row if pd.notna(val)])
                        if next_str.strip():
                            # Check if this row contains primarily numeric data (potential quantity row)
                            numeric_values = []
                            for val in next_row:
                                if pd.notna(val):
                                    try:
                                        numeric_val = float(
                                            str(val).replace(',', ''))
                                        numeric_values.append(numeric_val)
                                    except:
                                        pass

                            if numeric_values:
                                print(
                                    f"    Row {next_idx + 1}: {next_str} (numeric values: {numeric_values})")
                        else:
                            break  # Stop at empty row
            else:
                print(f"No raw rows found containing '{target}'")

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    debug_row_spanning()
