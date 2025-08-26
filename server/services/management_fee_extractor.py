import os
import re
import pdfplumber
import logging
from typing import List, Dict, Any, Optional
from io import BytesIO
import tempfile

logger = logging.getLogger(__name__)


class ManagementFeeExtractor:
    """
    Extract management fee subtables from PDF files.
    Specifically looks for rows where the "摘要" column contains "管理費区分:" followed by non-zero values.
    """

    def __init__(self, pdf_path: str):
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF not found: {pdf_path}")
        self.pdf_path = pdf_path

    def extract_management_fee_subtables(self, start_page: Optional[int] = None, end_page: Optional[int] = None) -> List[Dict[str, Any]]:
        """
        Extract management fee subtables from the specified page range.

        Args:
            start_page: Starting page number (1-based, None means start from page 1)
            end_page: Ending page number (1-based, None means extract all pages)

        Returns:
            List of dictionaries containing management fee subtable data
        """
        logger.info(f"=== EXTRACTING MANAGEMENT FEE SUBTABLES ===")
        logger.info(f"PDF file: {self.pdf_path}")
        logger.info(f"Page range: {start_page} to {end_page}")

        all_subtables = []

        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                # Determine page range
                if start_page is None:
                    start_page = 1
                if end_page is None:
                    end_page = len(pdf.pages)

                # Convert to 0-based indexing
                start_idx = start_page - 1
                end_idx = end_page

                logger.info(f"Processing pages {start_idx + 1} to {end_idx}")

                for page_num in range(start_idx, end_idx):
                    if page_num >= len(pdf.pages):
                        break

                    page = pdf.pages[page_num]
                    logger.info(f"Processing page {page_num + 1}")

                    # Extract tables from the page
                    tables = page.extract_tables()

                    for table_idx, table in enumerate(tables):
                        if not table or len(table) < 2:
                            continue

                        # Check if this table contains management fee data
                        management_fee_data = self._extract_management_fee_from_table(
                            table, page_num + 1, table_idx
                        )

                        if management_fee_data:
                            all_subtables.extend(management_fee_data)

        except Exception as e:
            logger.error(
                f"Error extracting management fee subtables: {str(e)}", exc_info=True)
            raise

        logger.info(
            f"Extracted {len(all_subtables)} management fee subtable rows")
        return all_subtables

    def _extract_management_fee_from_table(self, table: List[List[str]], page_num: int, table_idx: int) -> List[Dict[str, Any]]:
        """
        Extract management fee data from a single table.

        Args:
            table: Table data as list of rows
            page_num: Page number (1-based)
            table_idx: Table index on the page

        Returns:
            List of management fee data dictionaries
        """
        management_fee_data = []

        # Find the header row and column positions
        header_row_idx, column_mapping = self._find_header_and_columns(table)

        if header_row_idx is None or not column_mapping:
            logger.debug(
                f"No valid header found in table {table_idx} on page {page_num}")
            # Debug: Log the first few rows to see what we're working with
            for i, row in enumerate(table[:5]):
                logger.debug(f"Row {i}: {row}")
            return management_fee_data

        # Find reference numbers in the table
        current_reference = None

        # Look for reference numbers in the table
        for row_idx in range(len(table)):
            row = table[row_idx]
            if not row:
                continue

            # Check if this row contains a reference number
            reference_number = self._find_reference_in_row(row)
            if reference_number:
                current_reference = reference_number
                logger.debug(
                    f"Found reference number: {current_reference} at row {row_idx}")
                break

            # Extract data rows
        for row_idx in range(header_row_idx + 1, len(table)):
            row = table[row_idx]
            if not row:
                continue

            # Debug: Log rows that might contain management fee data
            notes_col = column_mapping.get('摘要')
            if notes_col is not None and notes_col < len(row):
                notes_text = str(row[notes_col]).strip()
                if '管理費区分' in notes_text:
                    logger.debug(
                        f"Found potential management fee row {row_idx}: {notes_text}")

            # Check if this row has management fee data
            management_fee_item = self._extract_management_fee_row(
                row, column_mapping, page_num, row_idx, current_reference
            )

            if management_fee_item:
                management_fee_data.append(management_fee_item)

        return management_fee_data

    def _find_header_and_columns(self, table: List[List[str]]) -> tuple[Optional[int], Dict[str, int]]:
        """
        Find the header row and map column positions.

        Args:
            table: Table data as list of rows

        Returns:
            Tuple of (header_row_index, column_mapping)
        """
        column_mapping = {}

        for row_idx, row in enumerate(table):
            if not row:
                continue

            # Look for header row with expected columns
            header_found = False
            for col_idx, cell in enumerate(row):
                if not cell:
                    continue

                cell_text = str(cell).strip()

                # Check for expected column headers - handle various formats
                # Normalize cell text for comparison
                normalized_cell = cell_text.replace('　', ' ').replace('：', ':')

                if ('名称' in normalized_cell and ('規格' in normalized_cell or '・' in normalized_cell)) or '名称・規格' in normalized_cell:
                    column_mapping['名称・規格'] = col_idx
                    header_found = True
                elif '単位' in normalized_cell or '単　位' in cell_text:
                    column_mapping['単位'] = col_idx
                    header_found = True
                elif '数量' in normalized_cell or '数　量' in cell_text:
                    column_mapping['数量'] = col_idx
                    header_found = True
                elif '単価' in normalized_cell or '単　価' in cell_text:
                    column_mapping['単価'] = col_idx
                    header_found = True
                elif '金額' in normalized_cell or '金　額' in cell_text:
                    column_mapping['金額'] = col_idx
                    header_found = True
                elif '摘要' in normalized_cell or '摘　要' in cell_text:
                    column_mapping['摘要'] = col_idx
                    header_found = True

            # At least name, unit, and notes columns
            if header_found and len(column_mapping) >= 3:
                logger.debug(
                    f"Found header row at index {row_idx}: {column_mapping}")
                logger.debug(f"Header row content: {row}")
                return row_idx, column_mapping

        return None, {}

    def _find_reference_in_row(self, row: List[str]) -> Optional[str]:
        """Find reference numbers like 内4号, 単3号, etc. in a row."""
        if not row:
            return None

        # Join all cells in the row to search for reference patterns
        row_text = ' '.join([str(cell) if cell else '' for cell in row])

        # Pattern: any kanji character(s) + digits + 号
        # Handles full-width and half-width characters, with optional spaces
        pattern = r'([一-龯々]+)\s*(\d+)\s*号'

        matches = re.findall(pattern, row_text)
        if matches:
            kanji, number = matches[0]  # Take the first match
            clean_ref = f"{kanji}{number}号"
            logger.debug(f"Found reference number in row: {clean_ref}")
            return clean_ref

        return None

    def _extract_management_fee_row(self, row: List[str], column_mapping: Dict[str, int],
                                    page_num: int, row_idx: int, reference_number: Optional[str] = None) -> Optional[Dict[str, Any]]:
        """
        Extract management fee data from a single row.

        Args:
            row: Row data
            column_mapping: Column position mapping
            page_num: Page number
            row_idx: Row index

        Returns:
            Management fee data dictionary or None if not a management fee row
        """
        if not row or len(row) == 0:
            return None

        # Check if this row has management fee data in the 摘要 column
        notes_col = column_mapping.get('摘要')
        if notes_col is None or notes_col >= len(row):
            return None

        notes_text = str(row[notes_col]).strip()

        # Check if this is a management fee row - handle various formats
        # Normalize the text to handle full-width and half-width characters
        normalized_text = notes_text.replace(
            '　', ' ')  # Full-width space to half-width
        normalized_text = normalized_text.replace(
            '：', ':')  # Full-width colon to half-width

        # Check for management fee pattern with various formats
        management_fee_patterns = [
            r'管理費区分\s*:\s*([^\s]+)',  # Standard format
            r'管理費区分\s*：\s*([^\s]+)',  # Full-width colon
            r'管理費区分\s*:\s*([^\s]+)',   # Half-width colon
            r'管理費区分\s*：\s*([^\s]+)',  # Full-width colon with spaces
            r'管理費区分\s*:\s*([^\s]+)',   # Half-width colon with spaces
        ]

        fee_value = None
        for pattern in management_fee_patterns:
            match = re.search(pattern, normalized_text)
            if match:
                fee_value = match.group(1).strip()
                break

        if not fee_value:
            return None

        # Check if the value is not zero (handle various zero formats)
        zero_patterns = ['0', '０', 'O', 'Ｏ', 'o', 'ｏ']
        if fee_value in zero_patterns:
            return None

        # Extract other column data
        item_name = ''
        unit = ''
        quantity = ''
        unit_price = ''
        amount = ''

        name_col = column_mapping.get('名称・規格')
        if name_col is not None and name_col < len(row):
            item_name = str(row[name_col]).strip()

        unit_col = column_mapping.get('単位')
        if unit_col is not None and unit_col < len(row):
            unit = str(row[unit_col]).strip()

        quantity_col = column_mapping.get('数量')
        if quantity_col is not None and quantity_col < len(row):
            quantity = str(row[quantity_col]).strip()

        price_col = column_mapping.get('単価')
        if price_col is not None and price_col < len(row):
            unit_price = str(row[price_col]).strip()

        amount_col = column_mapping.get('金額')
        if amount_col is not None and amount_col < len(row):
            amount = str(row[amount_col]).strip()

        # Create the management fee item
        management_fee_item = {
            'item_name': item_name,
            'unit': unit,
            'quantity': quantity,
            'unit_price': unit_price,
            'amount': amount,
            'management_fee_category': fee_value,
            'reference_number': reference_number or 'Unknown',
            'notes': notes_text,
            'page_number': page_num,
            'row_number': row_idx + 1,
            'raw_fields': {
                '名称・規格': item_name,
                '単位': unit,
                '数量': quantity,
                '単価': unit_price,
                '金額': amount,
                '摘要': notes_text
            }
        }

        logger.debug(f"Found management fee item: {item_name} - {fee_value}")
        logger.debug(f"Original notes text: '{notes_text}'")
        logger.debug(f"Normalized text: '{normalized_text}'")
        logger.debug(f"Extracted fee value: '{fee_value}'")
        return management_fee_item

    def close(self):
        """Clean up resources."""
        pass
