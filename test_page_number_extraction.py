#!/usr/bin/env python3
"""
Test script to test the new page number functionality for estimate extraction
"""

from server.services.estimate_extractor import EstimateReferenceExtractor
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def test_page_number_extraction():
    """Test the estimate extraction with different page numbers"""

    test_pdf_path = "07_入札時（見積）積算参考資料.pdf"

    if not os.path.exists(test_pdf_path):
        print(f"Test PDF not found: {test_pdf_path}")
        print("Please ensure the test PDF file exists in the backend directory")
        return

    extractor = None
    try:
        print("=== TESTING PAGE NUMBER EXTRACTION ===")
        extractor = EstimateReferenceExtractor(test_pdf_path)

        print(
            f"PDF loaded successfully. Number of pages: {len(extractor.pdf_pages)}")

        # Test extraction from different pages
        # Test pages 1, 2, 3, 4, 5 (adjust based on your PDF)
        test_pages = [0, 1, 2, 3, 4]

        for page_index in test_pages:
            if page_index < len(extractor.pdf_pages):
                print(
                    f"\n--- Testing Page {page_index + 1} (index {page_index}) ---")
                try:
                    result = extractor.extract_estimate_info(page_index)

                    print("Extraction Results:")
                    for key, value in result.items():
                        print(f"  {key}: {value}")

                except Exception as e:
                    print(
                        f"Error extracting from page {page_index + 1}: {str(e)}")
            else:
                print(f"\n--- Page {page_index + 1} does not exist ---")

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
    test_page_number_extraction()
