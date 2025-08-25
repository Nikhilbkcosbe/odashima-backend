#!/usr/bin/env python3
"""
Test script to run the improved extraction and see the results
"""

from server.services.estimate_extractor import EstimateReferenceExtractor
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def test_improved_extraction():
    """Test the improved estimate extraction"""

    test_pdf_path = "05.入札時（見積）積算参考資料.pdf"

    if not os.path.exists(test_pdf_path):
        print(f"Test PDF not found: {test_pdf_path}")
        return

    extractor = None
    try:
        print("=== TESTING IMPROVED ESTIMATE EXTRACTION ===")
        extractor = EstimateReferenceExtractor(test_pdf_path)

        print(
            f"PDF loaded successfully. Number of pages: {len(extractor.pdf_pages)}")

        if len(extractor.pdf_pages) >= 2:
            print("Extracting information from page 2...")
            result = extractor.extract_estimate_info(1)  # page index 1 = page 2

            print("\n=== EXTRACTION RESULTS ===")
            for key, value in result.items():
                print(f"{key}: {value}")

            # Also show the raw text for debugging
            page = extractor.pdf_pages[1]
            page_text = page.extract_text() or ""
            page_text = extractor._clean_text(page_text)

            print(f"\n=== TEXT LENGTH: {len(page_text)} ===")
            print(f"=== FIRST 500 CHARS: ===")
            print(page_text[:500])
            print(f"\n=== LAST 500 CHARS: ===")
            print(page_text[-500:])

        else:
            print("PDF has less than 2 pages, cannot test extraction")

    except Exception as e:
        print(f"Error during testing: {str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        if extractor:
            try:
                extractor.close()
            except Exception:
                pass


if __name__ == "__main__":
    test_improved_extraction()
