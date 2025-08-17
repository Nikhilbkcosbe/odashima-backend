"""
Table Title Extractor
A standalone module for extracting table titles from both PDF and Excel documents.
This module provides functions to extract table title information that appears
between reference numbers and column headers.
"""

import re
import logging
from typing import List, Dict, Optional, Tuple
import pandas as pd

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def normalize_text(text: str) -> str:
    """
    Normalize text by removing spaces and converting full-width characters to half-width
    """
    if not text or pd.isna(text):
        return ""

    # Convert to string if not already
    text = str(text).strip()

    # Convert full-width characters to half-width
    full_to_half = str.maketrans(
        '０１２３４５６７８９ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ'
        'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ　',
        '0123456789abcdefghijklmnopqrstuvwxyz'
        'ABCDEFGHIJKLMNOPQRSTUVWXYZ '
    )
    text = text.translate(full_to_half)

    # Remove all spaces for comparison
    text = re.sub(r'\s+', '', text)

    return text


def extract_pdf_table_title_items(table: List[List[str]], reference_row_idx: int, header_row_idx: int) -> Optional[Dict[str, str]]:
    """
    Extract table title items from PDF table that appear between reference number and column headers.

    Args:
        table: The table data as list of lists
        reference_row_idx: Index of the row containing the reference number
        header_row_idx: Index of the row containing column headers

    Returns:
        Dictionary with table title items or None if not found
    """
    try:
        # Check if we have a reference row
        if reference_row_idx >= len(table):
            return None

        reference_row = table[reference_row_idx]

        # Look for table title structure in the reference row itself
        # The table title is embedded within the reference number row
        # Structure can be either 7 cells or 8 cells:
        # 7 cells: [Reference, Item Name, 単位, Unit, 単位数量, Quantity, 単価]
        # 8 cells: [Reference, Item Name, Specification, 単位, Unit, 単位数量, Quantity, 単価]

        if len(reference_row) < 7:
            return None

        # Find positions of "単位" and "単位数量"
        unit_pos = None
        unit_qty_pos = None

        for i, cell in enumerate(reference_row):
            if cell and "単位" in str(cell) and "単位数量" not in str(cell):
                unit_pos = i
            elif cell and "単位数量" in str(cell):
                unit_qty_pos = i

        if unit_pos is None or unit_qty_pos is None:
            return None

        # Extract item name from cells after reference number
        item_name_parts = []
        for i in range(1, unit_pos):  # Start from 1 to skip reference number
            if i < len(reference_row) and reference_row[i]:
                item_name_parts.append(str(reference_row[i]).strip())

        item_name = " ".join(item_name_parts) if item_name_parts else ""

        # Extract unit from cell after "単位"
        unit = ""
        if unit_pos + 1 < len(reference_row):
            unit = str(reference_row[unit_pos + 1]).strip()

        # Extract unit quantity from cell after "単位数量"
        unit_quantity = ""
        if unit_qty_pos + 1 < len(reference_row):
            unit_quantity = str(reference_row[unit_qty_pos + 1]).strip()

        # Validate that we have the required components
        if not unit or not unit_quantity:
            return None

        return {
            "item_name": item_name,
            "unit": unit,
            "unit_quantity": unit_quantity
        }

    except Exception as e:
        logger.error(f"Error extracting PDF table title: {e}")
        return None


def extract_excel_table_title_items(df: pd.DataFrame, reference_row: int, header_row: int, pdf_title_item_name: str = None) -> Optional[Dict[str, str]]:
    """
    Extract table title items from Excel subtable.

    Args:
        df: DataFrame containing the Excel data
        reference_row: Row index containing the reference number
        header_row: Row index containing the column headers
        pdf_title_item_name: Optional PDF title item name to use for intelligent selection

    Returns:
        Dictionary with item_name, unit, and unit_quantity, or None if no title found
    """
    try:
        # Find the previous table end (or start of sheet if no previous table)
        prev_table_end = find_previous_table_end(df, reference_row)

        # Collect sentences between previous table end and reference number
        sentences_before_ref = []

        # Helper function to check if text is meaningful
        def is_meaningful_text(text):
            if not text or len(text.strip()) < 3:
                return False
            # Check if it contains Japanese characters, numbers, or Latin letters
            return bool(re.search(r'[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF0-9A-Za-z]', text))

        # Helper function to normalize text for comparison
        def normalize_text(text):
            if not text:
                return ""
            # Convert to string and normalize
            text = str(text).strip()
            # Convert full-width to half-width
            full_to_half = str.maketrans(
                '０１２３４５６７８９ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ'
                'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ　',
                '0123456789abcdefghijklmnopqrstuvwxyz'
                'ABCDEFGHIJKLMNOPQRSTUVWXYZ '
            )
            text = text.translate(full_to_half)
            # Remove spaces for comparison
            return re.sub(r'\s+', '', text)

        # Helper function to extract 1st part of PDF title
        def extract_first_part_of_pdf_title(pdf_title):
            if not pdf_title:
                return ""
            # Split by common separators and take the first meaningful part
            parts = re.split(r'[、，\s]+', pdf_title)
            return parts[0] if parts else ""

        # Helper function to find best matching sentence
        def find_best_matching_sentence(sentences, pdf_first_part):
            if not sentences:
                return None
            if len(sentences) == 1:
                return sentences[0]['text']

            # Always select the 2nd sentence from the back (closest to reference number)
            # If there's only 1 sentence, select that one
            # If there are 2 or more sentences, select the 2nd from the end
            if len(sentences) >= 2:
                selected_sentence = sentences[-2]['text']  # 2nd from the back
                logger.debug(
                    f"Selected 2nd sentence from back: '{selected_sentence}'")
            else:
                selected_sentence = sentences[0]['text']  # Only 1 sentence
                logger.debug(f"Selected only sentence: '{selected_sentence}'")

            return selected_sentence

        # Collect sentences between previous table end and reference number
        prev_table_end = find_previous_table_end(df, reference_row)
        start_row = prev_table_end + 1 if prev_table_end is not None else 0
        logger.debug(
            f"Collecting sentences from row {start_row} to {reference_row}")

        for row_idx in range(start_row, reference_row):
            row_text = " ".join(
                [str(cell) for cell in df.iloc[row_idx] if pd.notna(cell) and str(cell).strip()])
            if is_meaningful_text(row_text):
                sentences_before_ref.append(
                    {'row': row_idx, 'text': row_text.strip()})
                logger.debug(
                    f"Added sentence from row {row_idx}: '{row_text.strip()}'")

        logger.debug(
            f"Collected {len(sentences_before_ref)} sentences: {[s['text'] for s in sentences_before_ref]}")

        # Extract 1st part of PDF title for matching
        pdf_first_part = extract_first_part_of_pdf_title(pdf_title_item_name)

        # Select title based on PDF title matching
        selected_title = None

        # Try sentences between previous table end and reference number
        if sentences_before_ref:
            selected_title = find_best_matching_sentence(
                sentences_before_ref, pdf_first_part)

        if selected_title:
            return {
                "item_name": selected_title,
                "unit": "",  # Excel titles don't have separate unit fields
                "unit_quantity": ""  # Excel titles don't have separate quantity fields
            }

        return None

    except Exception as e:
        logger.error(f"Error extracting Excel table title: {e}")
        return None


def find_previous_table_end(df: pd.DataFrame, current_reference_row: int) -> int:
    """
    Find the end of the previous table by looking for a table number (just a number, no prefix or suffix)
    before the current reference row.

    Args:
        df: The DataFrame containing the Excel data
        current_reference_row: Current reference row to search backwards from

    Returns:
        Row index where the previous table ends, or 0 if no previous table found
    """
    try:
        # Search backwards from the current reference row
        for row_idx in range(current_reference_row - 1, -1, -1):
            row_data = df.iloc[row_idx]

            # Check if this row contains just a number (table end marker)
            non_empty_cells = [str(cell).strip() for cell in row_data if pd.notna(
                cell) and str(cell).strip()]

            if len(non_empty_cells) == 1:
                cell_value = non_empty_cells[0]
                # Check if it's just a number (no prefix or suffix)
                if re.match(r'^\d+$', cell_value):
                    return row_idx

        # If no previous table end found, return 0 (start of sheet)
        return 0

    except Exception as e:
        logger.error(f"Error finding previous table end: {e}")
        return 0


def find_excel_table_end(df: pd.DataFrame, start_row: int) -> int:
    """
    Find the end of a table in Excel by looking for a table number (just a number, no prefix or suffix).

    Args:
        df: The DataFrame containing the Excel data
        start_row: Starting row to search from

    Returns:
        Row index where the table ends
    """
    try:
        for row_idx in range(start_row, len(df)):
            row_data = df.iloc[row_idx]

            # Check if this row contains just a number (table end marker)
            non_empty_cells = [str(cell).strip() for cell in row_data if pd.notna(
                cell) and str(cell).strip()]

            if len(non_empty_cells) == 1:
                cell_value = non_empty_cells[0]
                # Check if it's just a number (no prefix or suffix)
                if re.match(r'^\d+$', cell_value):
                    return row_idx

        # If no table end found, return the last row
        return len(df) - 1

    except Exception as e:
        logger.error(f"Error finding Excel table end: {e}")
        return len(df) - 1
