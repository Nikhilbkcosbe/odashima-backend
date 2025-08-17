import pdfplumber
import re
import json
import logging
from typing import List, Dict, Optional, Tuple, Any

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class SubtablePDFExtractor:
    def __init__(self):
        """Initialize the subtable extractor with flexible column patterns."""
        # Define flexible patterns for the 4 required columns
        self.column_patterns = {
            "名称・規格": [
                "名称・規格", "名称規格", "名　称　・　規　格", "名称　規格",
                "名称・規格　格", "名称・規格格", "名称", "規格", "項目", "品名"
            ],
            "単位": [
                "単位", "単　位", "単 位", "たんい", "たんい"
            ],
            "数量": [
                "数量", "数　量", "数 量", "すうりょう"
            ],
            "摘要": [
                "摘要", "摘　要", "摘 要", "備考", "てきよう"
            ]
        }

    def extract_subtables_from_pdf(self, pdf_path: str, start_page: int, end_page: int) -> Dict[str, Any]:
        """
        Extract subtables from PDF within specified page range.

        Args:
            pdf_path (str): Path to the PDF file
            start_page (int): Starting page number (1-based)
            end_page (int): Ending page number (1-based)

        Returns:
            Dict[str, Any]: JSON-like structure containing extracted subtables
        """
        logger.info(
            f"Starting subtable extraction from {pdf_path}, pages {start_page}-{end_page}")

        result = {
            "pdf_file": pdf_path,
            "page_range": {"start": start_page, "end": end_page},
            "subtables": [],
            "total_subtables": 0,
            "total_rows": 0
        }

        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)

                # Validate page range
                if start_page < 1 or end_page > total_pages or start_page > end_page:
                    logger.error(
                        f"Invalid page range: {start_page}-{end_page} (PDF has {total_pages} pages)")
                    return result

                # Process each page
                for page_num in range(start_page - 1, end_page):  # Convert to 0-based
                    logger.info(f"Processing page {page_num + 1}")
                    page = pdf.pages[page_num]
                    page_subtables = self._extract_subtables_from_page(
                        page, page_num + 1)
                    result["subtables"].extend(page_subtables)

                result["total_subtables"] = len(result["subtables"])
                result["total_rows"] = sum(len(subtable["rows"])
                                           for subtable in result["subtables"])

                logger.info(
                    f"Extraction complete: {result['total_subtables']} subtables, {result['total_rows']} total rows")

        except Exception as e:
            logger.error(f"Error during extraction: {e}")
            result["error"] = str(e)

        return result

    def _extract_subtables_from_page(self, page, page_num: int) -> List[Dict[str, Any]]:
        """Extract all subtables from a single page."""
        page_subtables = []

        try:
            # Get page text for reference number detection
            page_text = page.extract_text()
            if not page_text:
                logger.warning(f"No text found on page {page_num}")
                return page_subtables

            # Find all reference numbers on this page
            reference_numbers = self._find_reference_numbers(page_text)
            logger.info(
                f"Found reference numbers on page {page_num}: {reference_numbers}")

            if not reference_numbers:
                logger.info(f"No reference patterns found on page {page_num}")
                return page_subtables

            # Extract tables from page
            tables = page.extract_tables()
            if not tables:
                logger.info(f"No tables found on page {page_num}")
                return page_subtables

            logger.info(f"Found {len(tables)} tables on page {page_num}")

            # Process each table to find subtables
            for table_idx, table in enumerate(tables):
                if not table or len(table) < 2:
                    continue

                table_subtables = self._extract_subtables_from_table(
                    table, reference_numbers, page_num, table_idx
                )
                page_subtables.extend(table_subtables)

        except Exception as e:
            logger.error(f"Error processing page {page_num}: {e}")

        return page_subtables

    def _find_reference_numbers(self, text: str) -> List[str]:
        """Find reference numbers like 内4号, 単3号, etc. in text."""
        # Pattern: any kanji character(s) + digits + 号
        # Handles full-width and half-width characters, with optional spaces
        pattern = r'([一-龯々]+)\s*(\d+)\s*号'

        matches = re.findall(pattern, text)
        reference_numbers = []

        for kanji, number in matches:
            # Remove any spaces and create clean reference number
            clean_ref = f"{kanji}{number}号"
            if clean_ref not in reference_numbers:
                reference_numbers.append(clean_ref)

        logger.debug(f"Extracted reference numbers: {reference_numbers}")
        return reference_numbers

    def _extract_subtables_from_table(self, table: List[List[str]], reference_numbers: List[str],
                                      page_num: int, table_idx: int) -> List[Dict[str, Any]]:
        """Extract subtables from a single table."""
        if not table or len(table) < 5:
            return []

        subtables = []
        current_reference = None
        current_subtable_rows = []
        current_column_mapping = None
        current_table_title = None
        processed_rows = set()  # Track processed rows to avoid duplicates

        for row_idx, row in enumerate(table):
            # Skip rows that were already processed as part of multi-row extraction
            if row_idx in processed_rows:
                logger.info(
                    f"🎯 DEBUG: Skipping row {row_idx} (already processed): {row}")
                continue

            row_text = ' '.join([str(cell) if cell else '' for cell in row])

            # Debug logging for internal subtable rows
            if row[0] and ('発生品運搬' in str(row[0]) or '交通誘導警備員' in str(row[0])):
                logger.info(
                    f"🎯 DEBUG: Processing internal subtable row {row_idx}: {row}")

            # Debug logging for key rows
            if row[0] and ('3745' in str(row[0]) or '合計' in str(row[0])):
                logger.info(f"🎯 DEBUG: Processing row {row_idx}: {row}")

            # Process data rows for current subtable FIRST (before checking for references)
            if current_reference:
                # Check for total row (end of subtable)
                if self._is_total_row(row):
                    logger.info(
                        f"Found total row for {current_reference}, ending subtable")
                    logger.info(f"🎯 DEBUG: Total row content: {row}")
                    break

                # Skip header row - check if this row contains column headers
                row_text_check = ' '.join(
                    [str(cell) if cell else '' for cell in row])
                is_header_row = any(header in row_text_check for header in [
                                    '名称・規格', '単位', '数量', '摘要'])
                if is_header_row:
                    logger.info(
                        f"🎯 DEBUG: Skipping header row {row_idx}: {row}")
                    continue

                # Extract row data using multi-row logic (call once per subtable)
                if current_column_mapping and not current_subtable_rows:
                    # Debug logging for internal subtables and VP40*3745
                    if row_idx < len(table) and table[row_idx][0]:
                        row_0_text = str(table[row_idx][0])
                        if '3745' in row_0_text:
                            logger.info(
                                f"🎯 DEBUG: About to call multi-row extraction for VP40*3745 from row {row_idx}")
                        elif '発生品運搬' in row_0_text or '交通誘導警備員' in row_0_text:
                            logger.info(
                                f"🎯 DEBUG: About to call multi-row extraction for internal subtable from row {row_idx}: {row_0_text}")

                    # Call multi-row extraction once from the first data row
                    extracted_rows, processed_indices = self._extract_multirow_data(
                        table, row_idx, current_column_mapping, current_reference)
                    # Add all extracted logical rows
                    for row_data in extracted_rows:
                        if row_data and any(row_data.values()):
                            # Debug logging for result
                            if '3745' in row_data.get('名称・規格', ''):
                                logger.info(
                                    f"🎯 DEBUG: VP40*3745 row data result: {row_data}")
                            elif '発生品運搬' in row_data.get('名称・規格', '') or '交通誘導警備員' in row_data.get('名称・規格', ''):
                                logger.info(
                                    f"🎯 DEBUG: Internal subtable row data result: {row_data}")
                            current_subtable_rows.append(row_data)
                    # Update processed_rows
                    processed_rows.update(processed_indices)
                continue  # Skip reference checking for data rows

            # Check if this row contains a reference number (only when not processing a subtable)
            found_reference = self._find_reference_in_row(
                row_text, reference_numbers)

            if found_reference:
                logger.info(
                    f"Found reference {found_reference} at row {row_idx + 1}")

                # If we have an existing subtable, save it first
                if current_reference and current_subtable_rows:
                    subtable = self._create_subtable_dict(
                        current_reference, current_subtable_rows, page_num, table_idx, current_table_title
                    )
                    subtables.append(subtable)
                    current_subtable_rows = []
                    current_table_title = None  # Reset table title for next subtable

                # Check if this is a real subtable header by looking for column headers
                # in the next few rows (usually within 3 rows)
                column_mapping = None
                header_row_idx = None

                for check_idx in range(row_idx + 1, min(row_idx + 4, len(table))):
                    if check_idx < len(table):
                        potential_headers = self._find_column_headers(
                            table[check_idx])
                        if potential_headers:
                            column_mapping = potential_headers
                            header_row_idx = check_idx
                            logger.info(
                                f"Found column headers for {found_reference}: {column_mapping}")
                            break

                # Only treat as subtable header if we found valid column headers
                if column_mapping and header_row_idx is not None:
                    # Extract table title items between reference row and header row
                    table_title_items = self._extract_table_title_items(
                        table, row_idx, header_row_idx)

                    current_reference = found_reference
                    current_column_mapping = column_mapping
                    current_table_title = table_title_items  # Store table title items
                    # Skip to after the header row for data extraction
                    continue
                else:
                    # This is just a reference number in the data, not a subtable header
                    logger.info(
                        f"Reference {found_reference} found in data content, not treating as subtable header")
                    continue

        # Don't forget the last subtable if no total row was found
        if current_reference and current_subtable_rows:
            subtable = self._create_subtable_dict(
                current_reference, current_subtable_rows, page_num, table_idx, current_table_title
            )
            subtables.append(subtable)

        return subtables

    def _find_reference_in_row(self, row_text: str, reference_numbers: List[str]) -> Optional[str]:
        """Check if a row contains any of the reference numbers."""
        for ref_num in reference_numbers:
            # Extract kanji and number parts properly
            match = re.match(r'([一-龯々]+)(\d+)号', ref_num)
            if not match:
                continue

            kanji_part = match.group(1)
            number_part = match.group(2)

            # Create flexible pattern that handles spaces
            pattern = f"{kanji_part}\\s+{number_part}\\s*号"
            if re.search(pattern, row_text):
                return ref_num
        return None

    def _find_column_headers(self, row: List[str]) -> Optional[Dict[str, int]]:
        """Find column header positions in a row."""
        column_mapping = {}

        for col_idx, cell in enumerate(row):
            if not cell:
                continue

            cell_text = str(cell).strip()

            # Check this cell against each column pattern
            matched = False
            for col_name, patterns in self.column_patterns.items():
                if matched or col_name in column_mapping:
                    continue  # Skip if we already matched this cell or found this column

                for pattern in patterns:
                    # First try exact match
                    if pattern == cell_text:
                        column_mapping[col_name] = col_idx
                        matched = True
                        break

                    # Then try flexible matching - remove spaces and special characters
                    clean_cell = re.sub(r'[\s　・]', '', cell_text)
                    clean_pattern = re.sub(r'[\s　・]', '', pattern)

                    if clean_pattern in clean_cell or clean_cell in clean_pattern:
                        column_mapping[col_name] = col_idx
                        matched = True
                        break

                if matched:
                    break

        logger.info(f"Column mapping result: {column_mapping}")

        # Return mapping only if we found at least 2 of the 4 required columns
        if len(column_mapping) >= 2:
            logger.info(
                f"✓ Found valid column mapping with {len(column_mapping)} columns")
            return column_mapping
        else:
            logger.info(
                f"✗ Not enough columns found ({len(column_mapping)}/2 minimum required)")

        return None

    def _is_total_row(self, row: List[str]) -> bool:
        """Check if a row contains 合計 (total)."""
        row_text = " ".join([str(cell) if cell else "" for cell in row])
        return "合計" in row_text

    def _extract_row_data(self, row: List[str], column_mapping: Dict[str, int]) -> Dict[str, str]:
        """Extract data from a row based on column mapping."""
        row_data = {
            "reference_number": "",  # This will be set later
            "名称・規格": "",
            "単位": "",
            "数量": "",
            "摘要": ""
        }

        for col_name, col_idx in column_mapping.items():
            if col_idx < len(row) and row[col_idx]:
                cell_value = str(row[col_idx]).strip()
                row_data[col_name] = cell_value

        return row_data

    def _extract_multirow_data(self, table: List[List[str]], start_row_idx: int,
                               column_mapping: Dict[str, int], reference_number: str) -> Tuple[List[Dict[str, str]], List[int]]:
        """Extract data that may span multiple rows, creating separate logical rows for each item.
        Returns: (list_of_extracted_rows, list_of_processed_row_indices)
        """
        if start_row_idx >= len(table):
            return [], []

        extracted_rows = []
        processed_indices = []

        # Debug flags
        debug_subtables = ['単14号', '単30号', '単40号', '内6号', '内9号']
        is_debug = reference_number in debug_subtables

        if is_debug:
            logger.info(
                f"🎯 DEBUG: Starting multi-row extraction for {reference_number} at row {start_row_idx}")

        # Process the table starting from start_row_idx
        current_idx = start_row_idx

        while current_idx < len(table):
            current_row = table[current_idx]

            # Stop if we encounter a 合計 (total) row
            if current_row[0] and '合計' in str(current_row[0]):
                if is_debug:
                    logger.info(f"🎯 DEBUG: Stopping at 合計 row {current_idx}")
                break

            # Check if this row has an item name (名称・規格)
            item_name = ""
            if (column_mapping.get('名称・規格', 0) < len(current_row) and
                    current_row[column_mapping.get('名称・規格', 0)]):
                potential_item = str(
                    current_row[column_mapping.get('名称・規格', 0)]).strip()
                # Skip header values
                if potential_item and potential_item not in ["名称・規格", "単位", "数量", "摘要"]:
                    item_name = potential_item

            if item_name:
                # Found an item name - start a new logical row
                row_data = {
                    "reference_number": reference_number,
                    "名称・規格": item_name,
                    "単位": "",
                    "数量": "",
                    "摘要": ""
                }

                if is_debug:
                    logger.info(
                        f"🎯 DEBUG: Found item '{item_name}' at row {current_idx}")

                # Look ahead up to 4 rows to find unit/quantity for THIS specific item
                item_processed_indices = [current_idx]

                # First, check the current row for any additional data
                for col_name, col_idx in column_mapping.items():
                    if col_name != '名称・規格' and col_idx < len(current_row) and current_row[col_idx]:
                        cell_value = str(current_row[col_idx]).strip()
                        if cell_value and cell_value not in ["名称・規格", "単位", "数量", "摘要"]:
                            row_data[col_name] = cell_value
                            if is_debug:
                                logger.info(
                                    f"🎯 DEBUG: Found {col_name} = '{cell_value}' in same row")

                # Look ahead up to 4 rows for missing unit/quantity
                for lookahead in range(1, 5):  # Look ahead 1-4 rows
                    lookahead_idx = current_idx + lookahead
                    if lookahead_idx >= len(table):
                        break

                    lookahead_row = table[lookahead_idx]

                    # Stop if we hit another item name or 合計
                    if lookahead_row[0]:
                        cell_text = str(lookahead_row[0]).strip()
                        if ('合計' in cell_text or
                            (len(cell_text) > 2 and cell_text != item_name and
                             cell_text not in ["", "名称・規格", "単位", "数量", "摘要"])):
                            if is_debug:
                                logger.info(
                                    f"🎯 DEBUG: Stopping lookahead at row {lookahead_idx}: '{cell_text}'")
                            break

                    # Look for missing unit/quantity in this lookahead row
                    found_data = False
                    for col_name in ['単位', '数量', '摘要']:
                        col_idx = column_mapping.get(col_name, -1)
                        if (col_idx != -1 and col_idx < len(lookahead_row) and
                                lookahead_row[col_idx] and not row_data[col_name]):
                            cell_value = str(lookahead_row[col_idx]).strip()
                            if cell_value and cell_value not in ["名称・規格", "単位", "数量", "摘要"]:
                                row_data[col_name] = cell_value
                                item_processed_indices.append(lookahead_idx)
                                found_data = True
                                if is_debug:
                                    logger.info(
                                        f"🎯 DEBUG: Found {col_name} = '{cell_value}' at lookahead row {lookahead_idx}")

                    # If we found some data, we can continue looking for more

                # Add this logical row to results
                extracted_rows.append(row_data)
                processed_indices.extend(item_processed_indices)

                if is_debug:
                    logger.info(
                        f"🎯 DEBUG: Completed logical row: 名称='{row_data['名称・規格']}', 単位='{row_data['単位']}', 数量='{row_data['数量']}', 摘要='{row_data['摘要']}'")

                # Move to the next unprocessed row
                current_idx = max(
                    current_idx + 1, max(item_processed_indices) + 1)
            else:
                # No item name in this row, skip it
                current_idx += 1

        if is_debug:
            logger.info(
                f"🎯 DEBUG: Extracted {len(extracted_rows)} logical rows")

        return extracted_rows, processed_indices

    def _create_subtable_dict(self, reference_number: str, rows: List[Dict[str, str]],
                              page_num: int, table_idx: int, table_title: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
        """Create a subtable dictionary structure."""
        subtable_dict = {
            "reference_number": reference_number,
            "page_number": page_num,
            "table_index": table_idx,
            "row_count": len(rows),
            "rows": rows
        }

        # Add table title items if they exist
        if table_title:
            subtable_dict["table_title"] = table_title

        return subtable_dict

    def _extract_table_title_items(self, table: List[List[str]], reference_row_idx: int, header_row_idx: int) -> Optional[Dict[str, str]]:
        """
        Extract table title items from the reference row itself.
        The table title information is embedded in the reference row structure.

        Args:
            table: The table data
            reference_row_idx: Index of the row containing the reference number
            header_row_idx: Index of the row containing column headers

        Returns:
            Dictionary with table title items or None if not found
        """
        if reference_row_idx >= len(table):
            return None

        # Extract table title from the reference row itself
        reference_row = table[reference_row_idx]
        if not reference_row:
            return None

        # Get non-empty cells from the reference row
        non_empty_cells = [
            cell for cell in reference_row if cell and str(cell).strip()]

        # We need at least 6 items for a valid table title structure
        if len(non_empty_cells) >= 6:
            # Find the positions of "単位" and "単位数量" in the cells
            unit_pos = -1
            unit_qty_pos = -1

            for i, cell in enumerate(non_empty_cells):
                cell_str = str(cell).strip()
                if "単位" in cell_str and "単位数量" not in cell_str:
                    unit_pos = i
                elif "単位数量" in cell_str:
                    unit_qty_pos = i

            # Check if we found both "単位" and "単位数量"
            if unit_pos != -1 and unit_qty_pos != -1 and unit_pos < unit_qty_pos:
                # Extract the items based on the structure
                # Reference number (単1号, etc.)
                reference_number = str(non_empty_cells[0]).strip()

                # Extract item name from the cell after reference number (cell 1)
                item_name = str(non_empty_cells[1]).strip() if len(
                    non_empty_cells) > 1 else ""

                # Extract unit value - it's in the cell after "単位"
                unit_value = str(
                    non_empty_cells[unit_pos + 1]).strip() if unit_pos + 1 < len(non_empty_cells) else ""

                # Extract unit quantity value - it's in the cell after "単位数量"
                unit_quantity_value = str(
                    non_empty_cells[unit_qty_pos + 1]).strip() if unit_qty_pos + 1 < len(non_empty_cells) else ""

                # Create table title structure - item_name is the actual item name, not including reference number
                table_title = {
                    "item_name": item_name,
                    "unit": unit_value,
                    "unit_quantity": unit_quantity_value
                }

                logger.info(
                    f"Found table title items in reference row: {table_title}")
                return table_title

        return None

    def _extract_unit_value(self, cell_text: str) -> str:
        """
        Extract unit value from cell text that contains "単位".
        Returns everything after "単位" until the next meaningful separator.
        """
        if not cell_text or "単位" not in cell_text:
            return ""

        # Find the position of "単位"
        unit_pos = cell_text.find("単位")
        if unit_pos == -1:
            return ""

        # Extract everything after "単位"
        after_unit = cell_text[unit_pos + 2:].strip()

        # Clean up the unit value - remove extra spaces and common separators
        unit_value = after_unit.split()[0] if after_unit else ""

        return unit_value

    def _extract_unit_quantity_value(self, cell_text: str) -> str:
        """
        Extract unit quantity value from cell text that contains "単位数量".
        Returns everything after "単位数量" until the next meaningful separator.
        """
        if not cell_text or "単位数量" not in cell_text:
            return ""

        # Find the position of "単位数量"
        unit_qty_pos = cell_text.find("単位数量")
        if unit_qty_pos == -1:
            return ""

        # Extract everything after "単位数量"
        after_unit_qty = cell_text[unit_qty_pos + 4:].strip()

        # Clean up the quantity value - remove extra spaces and common separators
        quantity_value = after_unit_qty.split()[0] if after_unit_qty else ""

        return quantity_value


def extract_subtables_api(pdf_path: str, start_page: int, end_page: int) -> str:
    """
    API-ready function to extract subtables from PDF.

    Args:
        pdf_path (str): Path to the PDF file
        start_page (int): Starting page number (1-based)
        end_page (int): Ending page number (1-based)

    Returns:
        str: JSON string containing extracted subtables
    """
    extractor = SubtablePDFExtractor()
    result = extractor.extract_subtables_from_pdf(
        pdf_path, start_page, end_page)
    return json.dumps(result, ensure_ascii=False, indent=2)


# Test function to demonstrate usage
if __name__ == "__main__":
    # Test with the full specified range
    pdf_path = "../07_入札時（見積）積算参考資料.pdf"
    start_page = 13
    end_page = 82

    try:
        result_json = extract_subtables_api(pdf_path, start_page, end_page)
        print("Extraction completed successfully!")
        print(f"Result preview (first 500 characters):")
        print(result_json[:500] + "..." if len(result_json)
              > 500 else result_json)

        # Save to file for inspection
        with open("subtable_extraction_result.json", "w", encoding="utf-8") as f:
            f.write(result_json)
        print("Full result saved to 'subtable_extraction_result.json'")

    except Exception as e:
        print(f"Error during extraction: {e}")
