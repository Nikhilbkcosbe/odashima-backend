#!/usr/bin/env python3
"""
Simple debug script to compare PDF and Excel results.
"""

import sys
import os
from io import BytesIO

from server.services.excel_parser import ExcelParser
from server.services.pdf_parser import PDFParser
from server.services.matcher import Matcher

def debug_matching():
    """Simple matching debug."""
    
    # Set UTF-8 encoding for output
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    
    pdf_path = "../sample.pdf"
    excel_path = "../07_入札時（見積）積算参考資料.xlsx"
    
    print("=== PDF vs Excel Matching Debug ===")
    
    # Extract from PDF
    print("Extracting from PDF...")
    pdf_parser = PDFParser()
    pdf_items = pdf_parser.extract_tables(pdf_path)
    print(f"PDF items: {len(pdf_items)}")
    
    # Extract from Excel  
    print("Extracting from Excel...")
    excel_parser = ExcelParser()
    with open(excel_path, 'rb') as f:
        excel_items = excel_parser.extract_items_from_buffer_with_sheet(
            BytesIO(f.read()), sheet_name=None, item_name_column=None
        )
    print(f"Excel items: {len(excel_items)}")
    
    # Target keywords
    keywords = ["curvy", "onsite", "guard", "scaffold"]  # English approximations
    
    print("\nTarget items in PDF:")
    for item in pdf_items:
        for keyword in keywords:
            if any(char in item.item_key for char in ["曲", "現", "防", "地"]):
                print(f"  {item.item_key[:50]}... (qty: {item.quantity})")
                break
    
    print("\nTarget items in Excel:")
    for item in excel_items:
        for keyword in keywords:
            if any(char in item.item_key for char in ["曲", "現", "防", "地"]):
                print(f"  {item.item_key[:50]}... (qty: {item.quantity})")
                break
    
    # Run matching
    print("\nRunning matcher...")
    matcher = Matcher()
    result = matcher.compare_items(pdf_items, excel_items)
    
    print(f"Total items: {result.total_items}")
    print(f"Matched: {result.matched_items}")
    print(f"Quantity mismatches: {result.quantity_mismatches}")
    print(f"Missing in Excel: {result.missing_items}")
    print(f"Extra in Excel: {result.extra_items}")
    
    # Show mismatches
    print("\nQuantity mismatches:")
    for r in result.results:
        if r.status == "QUANTITY_MISMATCH":
            print(f"  PDF: {r.pdf_item.item_key[:30]}... ({r.pdf_item.quantity})")
            print(f"  Excel: {r.excel_item.item_key[:30]}... ({r.excel_item.quantity})")
            print(f"  Difference: {r.quantity_difference}")
            print()

if __name__ == "__main__":
    debug_matching() 