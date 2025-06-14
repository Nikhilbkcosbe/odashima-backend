#!/usr/bin/env python3
"""
Test script to verify row spanning and empty row handling functionality.
"""

import sys
import os
import logging

# Add the server directory to the Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'server'))

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def test_pdf_row_spanning():
    """Test PDF parser with row spanning logic."""
    try:
        from server.services.pdf_parser import PDFParser

        logger.info("=== TESTING PDF ROW SPANNING LOGIC ===")

        # Create a test PDF parser
        parser = PDFParser()

        # Test empty row detection
        empty_row = ["", None, "", ""]
        is_empty = parser._is_completely_empty_row(empty_row)
        logger.info(f"Empty row detection: {is_empty} (should be True)")

        # Test quantity-only row detection
        raw_fields = {}  # No meaningful fields
        quantity = 5.0
        is_quantity_only = parser._is_quantity_only_row(raw_fields, quantity)
        logger.info(
            f"Quantity-only row detection: {is_quantity_only} (should be True)")

        # Test meaningful key creation
        meaningful_fields = {"工種": "土工", "種別": "掘削", "名称・規格": "土砂掘削"}
        key = parser._create_item_key_from_fields(meaningful_fields)
        logger.info(f"Meaningful key created: '{key}' (should not be empty)")

        # Test empty fields key creation
        empty_fields = {}
        empty_key = parser._create_item_key_from_fields(empty_fields)
        logger.info(f"Empty key created: '{empty_key}' (should be empty)")

        logger.info("✅ PDF row spanning tests completed")

    except ImportError as e:
        logger.error(f"Import error (packages may need to be installed): {e}")
    except Exception as e:
        logger.error(f"Error testing PDF row spanning: {e}")


def test_excel_row_spanning():
    """Test Excel parser with row spanning logic."""
    try:
        import pandas as pd
        from server.services.excel_parser import ExcelParser

        logger.info("=== TESTING EXCEL ROW SPANNING LOGIC ===")

        # Create a test Excel parser
        parser = ExcelParser()

        # Test empty row detection
        empty_series = pd.Series([None, "", pd.NaN, ""])
        is_empty = parser._is_completely_empty_row(empty_series)
        logger.info(f"Empty row detection: {is_empty} (should be True)")

        # Test non-empty row detection
        non_empty_series = pd.Series([None, "", "some_value", ""])
        is_not_empty = parser._is_completely_empty_row(non_empty_series)
        logger.info(
            f"Non-empty row detection: {is_not_empty} (should be False)")

        # Test quantity-only row detection
        raw_fields = {}  # No meaningful fields
        quantity = 3.5
        is_quantity_only = parser._is_quantity_only_row(raw_fields, quantity)
        logger.info(
            f"Quantity-only row detection: {is_quantity_only} (should be True)")

        # Test meaningful key creation
        meaningful_fields = {"工種": "土工", "種別": "掘削", "名称": "土砂掘削"}
        key = parser._create_item_key_from_fields(meaningful_fields)
        logger.info(f"Meaningful key created: '{key}' (should not be empty)")

        # Test empty fields key creation
        empty_fields = {}
        empty_key = parser._create_item_key_from_fields(empty_fields)
        logger.info(f"Empty key created: '{empty_key}' (should be empty)")

        logger.info("✅ Excel row spanning tests completed")

    except ImportError as e:
        logger.error(f"Import error (packages may need to be installed): {e}")
    except Exception as e:
        logger.error(f"Error testing Excel row spanning: {e}")


def test_quantity_merging():
    """Test quantity merging functionality."""
    try:
        from server.schemas.tender import TenderItem
        from server.services.pdf_parser import PDFParser

        logger.info("=== TESTING QUANTITY MERGING LOGIC ===")

        parser = PDFParser()

        # Create a test item
        existing_items = [
            TenderItem(
                item_key="土工|掘削|土砂掘削",
                raw_fields={"工種": "土工", "種別": "掘削", "名称・規格": "土砂掘削"},
                quantity=10.0,
                source="PDF"
            )
        ]

        logger.info(f"Initial item quantity: {existing_items[0].quantity}")

        # Test merging additional quantity
        result = parser._merge_quantity_with_previous_item(existing_items, 5.0)
        logger.info(f"Merge result: {result}")
        logger.info(
            f"Updated item quantity: {existing_items[0].quantity} (should be 15.0)")

        # Test merging with no existing items
        empty_list = []
        result_empty = parser._merge_quantity_with_previous_item(
            empty_list, 3.0)
        logger.info(
            f"Merge with empty list result: {result_empty} (should be 'skipped')")

        logger.info("✅ Quantity merging tests completed")

    except ImportError as e:
        logger.error(f"Import error (packages may need to be installed): {e}")
    except Exception as e:
        logger.error(f"Error testing quantity merging: {e}")


def test_simulated_table_processing():
    """Test simulated table processing with various row types."""
    try:
        from server.services.pdf_parser import PDFParser
        from server.schemas.tender import TenderItem

        logger.info("=== TESTING SIMULATED TABLE PROCESSING ===")

        parser = PDFParser()
        items = []

        # Simulate different types of rows
        test_rows = [
            # Normal row with all data
            {
                "row": ["土工", "掘削", "土砂掘削", "10.0", "m3"],
                "col_indices": {"工種": 0, "種別": 1, "名称・規格": 2, "数量": 3, "単位": 4},
                "description": "Normal row with all data"
            },
            # Completely empty row
            {
                "row": ["", "", "", "", ""],
                "col_indices": {"工種": 0, "種別": 1, "名称・規格": 2, "数量": 3, "単位": 4},
                "description": "Completely empty row (should be skipped)"
            },
            # Quantity-only row (row spanning)
            {
                "row": ["", "", "", "5.0", ""],
                "col_indices": {"工種": 0, "種別": 1, "名称・規格": 2, "数量": 3, "単位": 4},
                "description": "Quantity-only row (should merge with previous)"
            },
            # Row with only non-meaningful data
            {
                "row": ["", "", "", "", "m3"],
                "col_indices": {"工種": 0, "種別": 1, "名称・規格": 2, "数量": 3, "単位": 4},
                "description": "Row with only unit (should be skipped)"
            }
        ]

        for i, test_case in enumerate(test_rows):
            logger.info(
                f"Processing test row {i+1}: {test_case['description']}")

            result = parser._process_single_row_with_spanning(
                test_case["row"],
                test_case["col_indices"],
                1, 1, i+1,
                items
            )

            if isinstance(result, TenderItem):
                items.append(result)
                logger.info(
                    f"  ✅ Created item: {result.item_key} (quantity: {result.quantity})")
            else:
                logger.info(f"  ⏭️  Result: {result}")

        logger.info(f"Final items count: {len(items)}")
        for item in items:
            logger.info(f"  - {item.item_key}: quantity={item.quantity}")

        logger.info("✅ Simulated table processing tests completed")

    except ImportError as e:
        logger.error(f"Import error (packages may need to be installed): {e}")
    except Exception as e:
        logger.error(f"Error testing simulated table processing: {e}")


def main():
    """Run all row spanning tests."""
    logger.info("Starting row spanning and empty row handling tests...")

    test_pdf_row_spanning()
    test_excel_row_spanning()
    test_quantity_merging()
    test_simulated_table_processing()

    logger.info("All tests completed!")


if __name__ == "__main__":
    main()
