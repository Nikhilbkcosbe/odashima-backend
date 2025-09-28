import pdfplumber
import re
import json
import logging
from typing import List, Dict, Optional, Tuple, Any
from table_title_extractor import extract_pdf_table_title_items

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class SubtablePDFExtractor:
    def __init__(self):
        """Initialize the subtable extractor with flexible column patterns."""
        # Define flexible patterns for the 4 required columns
        self.column_patterns = {
            # Kitakami header variants are space-tolerant; we normalize before matching
            "名称・規格": [
                "名称・規格", "名称規格", "名　称　・　規　格", "名称　規格", "名 称 ・ 規 格",
                "名称・規格　格", "名称・規格格", "名称", "規格", "項目", "品名"
            ],
            "単位": [
                "単位", "単　位", "単 位", "たんい"
            ],
            "数量": [
                "数量", "数　量", "数 量", "すうりょう"
            ],
            # Optional columns frequently present in Kitakami subtables
            "単価": [
                "単価", "単 価", "たんか"
            ],
            "金額": [
                "金額", "金 額"
            ],
            "明細単価番号": [
                "明細単価番号", "明 細 単 価 番 号"
            ],
            "摘要": [
                "摘要", "摘　要", "摘 要", "備考", "てきよう"
            ]
        }

        # Global stop flag: set to True when "入力データ一覧表" is encountered
        self.stop_all_extraction = False

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

                    # Stop all extraction if global stop marker was encountered
                    if self.stop_all_extraction:
                        logger.info(
                            "Stop marker '入力データ一覧表' encountered. Halting further subtable extraction.")
                        break

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

                # Check global stop marker within this table before any processing
                if self._table_has_global_stop_marker(table):
                    self.stop_all_extraction = True
                    logger.info(
                        "Detected global stop row '入力データ一覧表' within table. Stopping now.")
                    return page_subtables

                table_subtables = self._extract_subtables_from_table(
                    table, reference_numbers, page_num, table_idx
                )
                page_subtables.extend(table_subtables)

                if self.stop_all_extraction:
                    logger.info(
                        "Global stop flag set during table processing. Exiting page early.")
                    return page_subtables

        except Exception as e:
            logger.error(f"Error processing page {page_num}: {e}")

        return page_subtables

    def _find_reference_numbers(self, text: str) -> List[str]:
        """Find reference numbers in text.

        Supports:
        - Standard: <Kanji><digits>号 (e.g., 内4号, 単3号)
        - Kitakami style: 第<digits>号<Kanji> (e.g., 第12号施)
        Both tolerate arbitrary spaces and full-width digits.
        """
        if not text:
            return []

        refs: List[str] = []

        # Kitakami style FIRST: 第 + digits + 号 + one Kanji (prefer the longer, more specific form)
        kita_pattern = r'第\s*([0-9０-９]+)\s*号\s*([一-龯])'
        for num, tail in re.findall(kita_pattern, text):
            try:
                num_norm = str(num).translate(
                    str.maketrans('０１２３４５６７８９', '0123456789'))
            except Exception:
                num_norm = str(num)
            value = f"第{num_norm}号{tail}"
            if value not in refs:
                refs.append(value)

        # Standard pattern: Kanji + digits + 号 (avoid overshadowing Kitakami-specific matches)
        std_pattern = r'([一-龯々]+)\s*([0-9０-９]+)\s*号'
        for kanji, num in re.findall(std_pattern, text):
            try:
                num_norm = str(num).translate(
                    str.maketrans('０１２３４５６７８９', '0123456789'))
            except Exception:
                num_norm = str(num)
            value = f"{kanji}{num_norm}号"
            # If this is a bare 第N号 and we already captured 第N号X, skip the shorter one
            if value.startswith("第"):
                if any(r.startswith(value) and len(r) > len(value) for r in refs):
                    continue
            if value not in refs:
                refs.append(value)

        logger.debug(f"Extracted reference numbers: {refs}")
        return refs

    def _extract_subtables_from_table(self, table: List[List[str]], reference_numbers: List[str],
                                      page_num: int, table_idx: int) -> List[Dict[str, Any]]:
        """Extract subtables from a single table."""
        if not table or len(table) < 5:
            return []

        subtables = []
        current_reference = None
        current_subtable_rows = []
        current_column_mapping = None
        current_is_kitakami = False  # Track if current subtable uses Kitakami-style headers
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

            # If we encounter the global stop marker anywhere, halt all extraction
            if self._row_is_global_stop_marker(row):
                self.stop_all_extraction = True
                logger.info(
                    "Encountered '入力データ一覧表' row. Halting all subtable extraction.")
                break

            # Process data rows for current subtable FIRST (before checking for references)
            if current_reference:
                # Check for total row (end of current subtable)
                if self._is_total_row(row, current_is_kitakami):
                    logger.info(
                        f"Found total row for {current_reference}, ending current subtable and searching for next reference")
                    logger.info(f"🎯 DEBUG: Total row content: {row}")
                    # Finalize current subtable
                    if current_subtable_rows:
                        subtable = self._create_subtable_dict(
                            current_reference, current_subtable_rows, page_num, table_idx, current_table_title
                        )
                        subtables.append(subtable)
                    # Reset state to look for the next subtable within the same table
                    current_reference = None
                    current_subtable_rows = []
                    current_column_mapping = None
                    current_table_title = None
                    # Continue scanning remaining rows for next reference
                    continue

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
                        table, row_idx, current_column_mapping, current_reference, current_is_kitakami)
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
                    current_reference = found_reference
                    current_column_mapping = column_mapping
                    # Determine Kitakami-like context for this subtable
                    current_is_kitakami = '明細単価番号' in (
                        current_column_mapping or {})

                    # Extract table title from the reference row
                    table_title = extract_pdf_table_title_items(
                        table, row_idx, header_row_idx)
                    if table_title:
                        current_table_title = table_title
                        logger.info(
                            f"Extracted table title for {found_reference}: {table_title}")

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
        """Return the first reference from reference_numbers that appears in this row.

        Matches both standard and Kitakami styles, tolerating spaces/width differences.
        """
        if not row_text:
            return None

        # Normalized containment check as a robust fallback
        norm_row = self._normalize_simple(row_text)

        # Prefer longer references first (e.g., 第12号施 over 第12号)
        try:
            sorted_refs = sorted(reference_numbers, key=lambda r: len(
                self._normalize_simple(r)), reverse=True)
        except Exception:
            sorted_refs = reference_numbers

        for ref_num in sorted_refs:
            # 1) Exact normalized containment
            if self._normalize_simple(ref_num) in norm_row:
                return ref_num

            # 2) Standard style regex (Kanji + digits + 号) with spaces
            m_std = re.match(r'([一-龯々]+)(\d+)号', ref_num)
            if m_std:
                kanji_part, number_part = m_std.group(1), m_std.group(2)
                pattern = f"{kanji_part}\\s*{number_part}\\s*号"
                if re.search(pattern, row_text):
                    return ref_num

            # 3) Kitakami style: 第 + digits + 号 + one Kanji
            m_kita = re.match(r'第(\d+)号([一-龯])', ref_num)
            if m_kita:
                num, tail = m_kita.group(1), m_kita.group(2)
                pattern = f"第\\s*{num}\\s*号\\s*{tail}"
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

    def _is_total_row(self, row: List[str], is_kitakami: bool = False) -> bool:
        """Return True if this row represents the total separator for the current subtable.

        - For all areas: rows containing '合計' anywhere end the subtable.
        - For Kitakami only: an exact '計' row (ignoring spaces/width) ends the subtable.
        """
        row_text = " ".join([str(cell) if cell else "" for cell in row])
        norm = self._normalize_simple(row_text)
        if "合計" in row_text:
            return True
        if is_kitakami and norm == "計":
            return True
        return False

    def _row_is_global_stop_marker(self, row: List[str]) -> bool:
        """Return True if row text equals '入力データ一覧表' ignoring width/spaces."""
        row_text = " ".join([str(cell) if cell else "" for cell in row])
        norm = self._normalize_simple(row_text)
        return "入力データ一覧表" in norm

    def _table_has_global_stop_marker(self, table: List[List[str]]) -> bool:
        for r in table:
            if self._row_is_global_stop_marker(r):
                return True
        return False

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
                               column_mapping: Dict[str, int], reference_number: str,
                               kitakami_mode: bool = False) -> Tuple[List[Dict[str, str]], List[int]]:
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
            if current_row[0] and ('合計' in str(current_row[0]) or (kitakami_mode and self._normalize_simple(str(current_row[0])) == '計')):
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
                        if col_name == '数量':
                            cell_value = self._merge_quantity_with_adjacent(
                                current_row, col_idx) or str(current_row[col_idx]).strip()
                        else:
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
                            (kitakami_mode and self._normalize_simple(cell_text) == '計') or
                            (len(cell_text) > 2 and cell_text != item_name and
                             cell_text not in ["", "名称・規格", "単位", "数量", "摘要"])):
                            if is_debug:
                                logger.info(
                                    f"🎯 DEBUG: Stopping lookahead at row {lookahead_idx}: '{cell_text}'")
                            break

                    # Look for missing unit/quantity/remarks/code in this lookahead row
                    found_data = False
                    for col_name in ['単位', '数量', '摘要', '明細単価番号']:
                        col_idx = column_mapping.get(col_name, -1)
                        if (col_idx != -1 and col_idx < len(lookahead_row) and
                                lookahead_row[col_idx] and not row_data.get(col_name)):
                            if col_name == '数量':
                                cell_value = self._merge_quantity_with_adjacent(
                                    lookahead_row, col_idx) or str(lookahead_row[col_idx]).strip()
                            else:
                                cell_value = str(
                                    lookahead_row[col_idx]).strip()
                            if cell_value and cell_value not in ["名称・規格", "単位", "数量", "摘要"]:
                                row_data[col_name] = cell_value
                                item_processed_indices.append(lookahead_idx)
                                found_data = True
                                if is_debug:
                                    logger.info(
                                        f"🎯 DEBUG: Found {col_name} = '{cell_value}' at lookahead row {lookahead_idx}")

                    # If we found some data, we can continue looking for more

                # If quantity is still empty but we have a quantity column, try merging using current row
                qty_col_idx = column_mapping.get('数量', -1)
                if not row_data['数量'] and qty_col_idx != -1:
                    merged_here = self._merge_quantity_with_adjacent(
                        current_row, qty_col_idx)
                    if merged_here:
                        row_data['数量'] = merged_here

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

        # Add table title if available
        if table_title:
            subtable_dict["table_title"] = table_title

        return subtable_dict

    def _normalize_simple(self, text: str) -> str:
        """Normalize text using NFKC and remove all spaces (ASCII and full-width)."""
        try:
            import unicodedata
            normalized = unicodedata.normalize('NFKC', text)
        except Exception:
            normalized = text
        return re.sub(r"[\s\u3000]+", "", normalized)

    # --- Kitakami quantity merge helpers ---
    def _merge_quantity_with_adjacent(self, row: List[str], qty_idx: int) -> str:
        """Return merged quantity string by combining integer in qty cell with adjacent decimal part if present."""
        try:
            if qty_idx < 0 or qty_idx >= len(row):
                return ""
            cell_text = self._normalize_simple(
                str(row[qty_idx])) if row[qty_idx] else ""
            if not cell_text:
                return ""

            main_num = self._extract_first_number(cell_text)
            if main_num is None:
                return ""
            if "." in main_num:
                return main_num

            # Look for decimal part in immediate neighbors
            neighbor = self._find_adjacent_decimal_part(row, qty_idx)
            if neighbor is not None:
                try:
                    # neighbor may be (digits, is_negative)
                    if isinstance(neighbor, tuple):
                        digits, neg_from_neighbor = neighbor
                    else:
                        digits, neg_from_neighbor = neighbor, False

                    integer_part = str(int(float(main_num)))
                    is_negative = integer_part.startswith(
                        '-') or neg_from_neighbor
                    integer_abs = integer_part.lstrip('-')
                    sign = '-' if is_negative else ''
                    return f"{sign}{integer_abs}.{digits}"
                except Exception:
                    return main_num
            return main_num
        except Exception:
            return ""

    def _find_adjacent_decimal_part(self, row: List[str], qty_idx: int) -> Optional[tuple[str, bool]]:
        for offset in [-1, 1]:
            check_idx = qty_idx + offset
            if 0 <= check_idx < len(row) and row[check_idx]:
                t = self._normalize_simple(str(row[check_idx]))
                # If looks like description with letters/kanji/units, skip
                if re.search(r"[A-Za-z一-龯号mktN明]", t):
                    continue
                # .xx or 0.xx (allow optional leading minus)
                m = re.search(r"^0\.(\d+)$", t)
                if m:
                    return (m.group(1), t.startswith('-'))
                m = re.search(r"^\.(\d+)$", t)
                if m:
                    return (m.group(1), t.startswith('-'))
                # -0.xx
                m = re.search(r"^-0\.(\d+)$", t)
                if m:
                    return (m.group(1), True)
                # Pure digits up to 3 chars (allow optional leading '-') → decimal digits
                if re.match(r"^-?\d{1,3}$", t):
                    return (t.lstrip('-'), t.startswith('-'))
        return None

    def _extract_first_number(self, text: str) -> Optional[str]:
        m = re.search(r"-?\d+(?:\.\d+)?", text)
        return m.group(0) if m else None

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
