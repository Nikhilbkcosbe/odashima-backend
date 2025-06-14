#!/usr/bin/env python3
"""
Test script to demonstrate the improved iterative table extraction
for the Construction Tender vs Proposal Reconciliation System.
"""

import sys
import os
import logging
from io import BytesIO

# Add the server directory to the Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'server'))

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def test_pdf_extraction():
    """Test the improved PDF extraction process."""
    try:
        from server.services.pdf_parser import PDFParser

        # Check if sample PDF exists
        pdf_path = "../sample.pdf"
        if not os.path.exists(pdf_path):
            logger.warning(f"Sample PDF not found at {pdf_path}")
            return

        logger.info("=== TESTING PDF EXTRACTION ===")
        parser = PDFParser()

        # Test iterative extraction
        items = parser.extract_tables(pdf_path)

        logger.info(f"Total items extracted: {len(items)}")

        # Show first few items
        for i, item in enumerate(items[:3]):
            logger.info(f"Item {i+1}: {item.item_key}")
            logger.info(f"  Quantity: {item.quantity}")
            logger.info(f"  Fields: {item.raw_fields}")

        if len(items) > 3:
            logger.info(f"... and {len(items) - 3} more items")

    except ImportError as e:
        logger.error(f"Import error (packages may need to be installed): {e}")
    except Exception as e:
        logger.error(f"Error testing PDF extraction: {e}")


def test_excel_extraction():
    """Test the improved Excel extraction process."""
    try:
        from server.services.excel_parser import ExcelParser

        # Check if sample Excel exists
        excel_path = "../水沢橋　積算書.xlsx"
        if not os.path.exists(excel_path):
            logger.warning(f"Sample Excel not found at {excel_path}")
            return

        logger.info("=== TESTING EXCEL EXTRACTION ===")
        parser = ExcelParser()

        # Test file-based extraction
        items = parser.extract_items(excel_path)

        logger.info(f"Total items extracted: {len(items)}")

        # Show first few items
        for i, item in enumerate(items[:3]):
            logger.info(f"Item {i+1}: {item.item_key}")
            logger.info(f"  Quantity: {item.quantity}")
            logger.info(f"  Fields: {item.raw_fields}")

        if len(items) > 3:
            logger.info(f"... and {len(items) - 3} more items")

        # Test buffer-based extraction
        logger.info("=== TESTING EXCEL BUFFER EXTRACTION ===")
        with open(excel_path, 'rb') as f:
            buffer = BytesIO(f.read())

        buffer_items = parser.extract_items_from_buffer(buffer)
        logger.info(f"Buffer extraction: {len(buffer_items)} items")

    except ImportError as e:
        logger.error(f"Import error (packages may need to be installed): {e}")
    except Exception as e:
        logger.error(f"Error testing Excel extraction: {e}")


def test_comparison():
    """Test the improved comparison process."""
    try:
        from server.services.pdf_parser import PDFParser
        from server.services.excel_parser import ExcelParser
        from server.services.matcher import Matcher

        pdf_path = "../sample.pdf"
        excel_path = "../水沢橋　積算書.xlsx"

        if not os.path.exists(pdf_path) or not os.path.exists(excel_path):
            logger.warning("Sample files not found, skipping comparison test")
            return

        logger.info("=== TESTING FULL COMPARISON PROCESS ===")

        # Extract items
        pdf_parser = PDFParser()
        excel_parser = ExcelParser()

        pdf_items = pdf_parser.extract_tables(pdf_path)
        excel_items = excel_parser.extract_items(excel_path)

        # Compare items
        matcher = Matcher()
        result = matcher.compare_items(pdf_items, excel_items)

        logger.info("=== COMPARISON RESULTS ===")
        logger.info(f"Total items: {result.total_items}")
        logger.info(f"Perfect matches: {result.matched_items}")
        logger.info(f"Quantity mismatches: {result.quantity_mismatches}")
        logger.info(f"Missing in Excel: {result.missing_items}")
        logger.info(f"Extra in Excel: {result.extra_items}")

        # Test missing items only
        logger.info("=== TESTING MISSING ITEMS EXTRACTION ===")
        missing_items = matcher.get_missing_items_only(pdf_items, excel_items)
        logger.info(f"Missing items count: {len(missing_items)}")

        # Test mismatched items only
        logger.info("=== TESTING MISMATCHED ITEMS EXTRACTION ===")
        mismatched_items = matcher.get_mismatched_items_only(
            pdf_items, excel_items)
        logger.info(f"Mismatched items count: {len(mismatched_items)}")

    except ImportError as e:
        logger.error(f"Import error (packages may need to be installed): {e}")
    except Exception as e:
        logger.error(f"Error testing comparison: {e}")


def main():
    """Run all tests."""
    logger.info("Starting extraction system tests...")

    test_pdf_extraction()
    test_excel_extraction()
    test_comparison()

    logger.info("Tests completed!")


if __name__ == "__main__":
    main()
