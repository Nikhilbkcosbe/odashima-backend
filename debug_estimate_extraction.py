#!/usr/bin/env python3
"""
Debug script to analyze PDF content and improve extraction patterns
"""

from server.services.estimate_extractor import EstimateReferenceExtractor
import sys
import os
import re
sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def debug_estimate_extraction():
    """Debug the estimate extraction by analyzing the actual PDF content"""

    test_pdf_path = "05.入札時（見積）積算参考資料.pdf"

    if not os.path.exists(test_pdf_path):
        print(f"Test PDF not found: {test_pdf_path}")
        return

    extractor = None
    try:
        print("=== DEBUGGING ESTIMATE EXTRACTION ===")
        extractor = EstimateReferenceExtractor(test_pdf_path)

        print(
            f"PDF loaded successfully. Number of pages: {len(extractor.pdf_pages)}")

        if len(extractor.pdf_pages) >= 2:
            page = extractor.pdf_pages[1]
            page_text = page.extract_text() or ""
            page_text = extractor._clean_text(page_text)

            print(f"\n=== FULL PAGE 2 TEXT ===")
            print(page_text)
            print(f"\n=== TEXT LENGTH: {len(page_text)} ===")

            # Search for specific patterns
            print("\n=== SEARCHING FOR PATTERNS ===")

            # Look for 省庁 patterns
            print("\n--- 省庁 Patterns ---")
            shocho_patterns = [
                r'([^\s]+[県都府市])',
                r'([^\s]+県[^\s]*)',
                r'([^\s]+都[^\s]*)',
                r'([^\s]+府[^\s]*)',
                r'([^\s]+市[^\s]*)',
            ]
            for i, pattern in enumerate(shocho_patterns):
                matches = re.findall(pattern, page_text)
                print(f"Pattern {i+1}: {pattern}")
                print(f"Matches: {matches}")

            # Look for 年度 patterns
            print("\n--- 年度 Patterns ---")
            nendo_patterns = [
                r'([令和平成昭和大正明治]\s*\d+\s*年度)',
                r'(\d+\s*年度)',
                r'([令和平成昭和大正明治]\s*\d+)',
            ]
            for i, pattern in enumerate(nendo_patterns):
                matches = re.findall(pattern, page_text)
                print(f"Pattern {i+1}: {pattern}")
                print(f"Matches: {matches}")

            # Look for 工種区分
            print("\n--- 工種区分 Patterns ---")
            keihi_patterns = [
                r'工種区分[^\n]*?([^\n]+)',
                r'工種区分\s*([^\n\s]+)',
            ]
            for i, pattern in enumerate(keihi_patterns):
                matches = re.findall(pattern, page_text)
                print(f"Pattern {i+1}: {pattern}")
                print(f"Matches: {matches}")

            # Look for 工事名
            print("\n--- 工事名 Patterns ---")
            kouji_patterns = [
                r'工\s*事\s*名[^\n]*?([^\n]+)',
                r'工\s*事\s*名\s*([^\n\s]+)',
                r'(一般国道\d+号[^\n]+工事)',
                r'([^\n]*橋[^\n]*工事)',
            ]
            for i, pattern in enumerate(kouji_patterns):
                matches = re.findall(pattern, page_text)
                print(f"Pattern {i+1}: {pattern}")
                print(f"Matches: {matches}")

            # Look for any text containing "一般国道"
            print("\n--- 一般国道 Search ---")
            highway_matches = re.findall(r'(一般国道[^\n]+)', page_text)
            print(f"Highway matches: {highway_matches}")

            # Look for any text containing "橋"
            print("\n--- 橋 Search ---")
            bridge_matches = re.findall(r'([^\n]*橋[^\n]*)', page_text)
            print(f"Bridge matches: {bridge_matches}")

        else:
            print("PDF has less than 2 pages")

    except Exception as e:
        print(f"Error during debugging: {str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        if extractor:
            try:
                extractor.close()
            except Exception:
                pass


if __name__ == "__main__":
    debug_estimate_extraction()
