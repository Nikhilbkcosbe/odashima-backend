#!/usr/bin/env python3
"""
Test script for the EstimateReferenceExtractor
"""

from server.services.estimate_extractor import EstimateReferenceExtractor
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def test_estimate_extractor():
    """Test the estimate extractor with a sample PDF"""

    # Test with one of the existing PDF files
    test_pdf_path = "05.入札時（見積）積算参考資料.pdf"

    if not os.path.exists(test_pdf_path):
        print(f"Test PDF not found: {test_pdf_path}")
        print("Please ensure the test PDF file exists in the backend directory")
        return

    try:
        print("Testing EstimateReferenceExtractor...")
        extractor = EstimateReferenceExtractor(test_pdf_path)

        print(
            f"PDF loaded successfully. Number of pages: {len(extractor.pdf_pages)}")

        if len(extractor.pdf_pages) >= 2:
            print("Extracting information from page 2...")
            result = extractor.extract_estimate_info()

            print("\n=== EXTRACTION RESULTS ===")
            for key, value in result.items():
                print(f"{key}: {value}")
        else:
            print("PDF has less than 2 pages, cannot test extraction")

    except Exception as e:
        print(f"Error during testing: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    test_estimate_extractor()
