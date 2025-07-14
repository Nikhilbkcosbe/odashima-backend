import pdfplumber
import os
import re
import logging
from typing import List, Dict, Tuple, Optional, Union
from ..schemas.tender import TenderItem, SubtableItem

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class PDFParser:
    def __init__(self):
        # Updated column patterns to match the specific PDF structure
        self.default_column_patterns = {
            "工事区分・工種・種別・細別": ["費 目 ・ 工 種 ・ 種 別 ・ 細 目", "費目・工種・種別・細別・規格", "工事区分・工種・種別・細別", "工事区分", "工種", "種別", "細別", "費目"],
            "規格": ["規格", "規 格", "名称・規格", "名称", "項目", "品名"],
            "単位": ["単位", "単 位"],
            "数量": ["数量", "数 量"],
            "単価": ["単価", "単 価"],
            "金額": ["金額", "金 額"],
            "数量・金額増減": ["数量・金額増減", "増減", "変更"],
            "摘要": ["摘要", "備考", "摘 要"]
        }

        self.column_patterns = self.default_column_patterns.copy()

    def extract_tables(self, pdf_path: str) -> List[TenderItem]:
        """
        Extract tables from PDF iteratively, page by page and table by table.
        """
        return self.extract_tables_with_range(pdf_path, None, None)

    def extract_tables_with_range(self, pdf_path: str, start_page: Optional[int] = None, end_page: Optional[int] = None) -> List[TenderItem]:
        """
        Extract tables from PDF iteratively with specified page range.

        Args:
            pdf_path: Path to the PDF file
            start_page: Starting page number (1-based, None means start from page 1)
            end_page: Ending page number (1-based, None means extract all pages)
        """

        all_items = []

        logger.info(f"Starting PDF extraction from: {pdf_path}")
        logger.info(
            f"Page range: {start_page or 'start'} to {end_page or 'end'}")

        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                logger.info(f"PDF has {total_pages} pages total")

                # Determine actual page range
                actual_start = (
                    start_page - 1) if start_page is not None else 0
                actual_end = (
                    end_page - 1) if end_page is not None else total_pages - 1

                # Validate page range
                actual_start = max(0, actual_start)
                actual_end = min(total_pages - 1, actual_end)

                if actual_start > actual_end:
                    logger.warning(
                        f"Invalid page range: start={actual_start+1}, end={actual_end+1}")
                    return all_items

                pages_to_process = actual_end - actual_start + 1
                logger.info(
                    f"Processing pages {actual_start + 1} to {actual_end + 1} ({pages_to_process} pages)")

                # Process specified page range iteratively
                for page_num in range(actual_start, actual_end + 1):
                    page = pdf.pages[page_num]
                    logger.info(
                        f"Processing page {page_num + 1}/{total_pages}")

                    page_items = self._extract_tables_from_page(page, page_num)

                    logger.info(
                        f"Extracted {len(page_items)} items from page {page_num + 1}")

                    # Join items from this page to the total collection
                    all_items.extend(page_items)

                logger.info(
                    f"Total items extracted from PDF (pages {actual_start + 1}-{actual_end + 1}): {len(all_items)}")

        except Exception as e:
            logger.error(f"Error processing PDF: {e}")
            raise

        return all_items

    def _extract_tables_from_page(self, page, page_num: int) -> List[TenderItem]:
        """
        Extract all tables from a single page iteratively.
        """
        page_items = []

        try:
            tables = page.extract_tables()

            if not tables:
                logger.info(f"No tables found on page {page_num + 1}")
                return page_items

            logger.info(f"Found {len(tables)} tables on page {page_num + 1}")

            # Process each table iteratively
            for table_num, table in enumerate(tables):
                logger.info(
                    f"Processing table {table_num + 1}/{len(tables)} on page {page_num + 1}")

                table_items = self._process_single_table(
                    table, page_num, table_num)

                logger.info(
                    f"Extracted {len(table_items)} items from table {table_num + 1}")

                # Join table items to page items
                page_items.extend(table_items)

        except Exception as e:
            logger.error(f"Error processing page {page_num + 1}: {e}")

        return page_items

    def _should_process_table(self, table: List[List], page_num: int, table_num: int) -> bool:
        """
        Check if a table should be processed.
        """
        return True  # Process all tables

    def _process_single_table(self, table: List[List], page_num: int, table_num: int) -> List[TenderItem]:
        """
        Process a single table and extract all valid items from it.
        """
        items = []

        if not table or len(table) < 2:
            logger.warning(
                f"Table {table_num + 1} on page {page_num + 1} is too small (less than 2 rows)")
            return items

        # Check if this table should be processed based on custom column requirements
        if not self._should_process_table(table, page_num, table_num):
            return items

        # Find header row
        header_row, header_idx = self._find_header_row(table)

        if header_row is None:
            logger.warning(
                f"No header row found in table {table_num + 1} on page {page_num + 1}")
            return items

        logger.info(
            f"Found header at row {header_idx + 1} in table {table_num + 1}")

        # Get column mapping
        col_indices = self._get_column_mapping(header_row)

        if not col_indices:
            logger.warning(
                f"No recognizable columns found in table {table_num + 1} on page {page_num + 1}")
            return items

        logger.info(f"Column mapping for table {table_num + 1}: {col_indices}")

        # Process data rows iteratively with enhanced row spanning logic
        data_rows = table[header_idx + 1:]
        logger.info(
            f"Processing {len(data_rows)} data rows in table {table_num + 1}")

        for row_idx, row in enumerate(data_rows):
            try:
                result = self._process_single_row_with_spanning(
                    row, col_indices, page_num, table_num, header_idx + 1 + row_idx, items
                )

                if result == "merged":
                    logger.debug(
                        f"Row {row_idx + 1} merged with previous item (quantity-only row)")
                elif result == "skipped":
                    logger.debug(f"Row {row_idx + 1} skipped (empty row)")
                elif result:
                    items.append(result)

            except Exception as e:
                logger.error(
                    f"Error processing row {row_idx + 1} in table {table_num + 1}: {e}")
                continue

        return items

    def _process_single_row_with_spanning(self, row: List, col_indices: Dict[str, int],
                                          page_num: int, table_num: int, row_num: int,
                                          existing_items: List) -> Union[TenderItem, str, None]:
        """
        Rewritten row spanning logic to properly handle:
        - Row 1: Has item name but no quantity (or quantity = 0)
        - Row 2: Has quantity and unit but no item name
        - Result: Combine into single item with name, quantity, and unit
        """
        # First check if row is completely empty
        if self._is_completely_empty_row(row):
            return "skipped"

        # Extract fields from row
        raw_fields = {}
        quantity = 0.0
        unit = None

        for col_name, col_idx in col_indices.items():
            if col_idx < len(row) and row[col_idx]:
                cell_value = str(row[col_idx]).strip()
                if cell_value:
                    if col_name == "数量":
                        quantity = self._extract_quantity(cell_value)
                    elif col_name == "単位":
                        unit = cell_value
                        # Also store in raw_fields for completion logic
                        raw_fields[col_name] = cell_value
                    else:
                        raw_fields[col_name] = cell_value

        # Check what type of data this row contains
        has_item_fields = self._has_item_identifying_fields(raw_fields)
        has_quantity_data = quantity > 0 or "単位" in raw_fields

        logger.debug(
            f"PDF Row {row_num}: has_item_fields={has_item_fields}, has_quantity_data={has_quantity_data}, quantity={quantity}")

        # Case 1: Row has item fields but no quantity data (incomplete item - needs spanning)
        if has_item_fields and not has_quantity_data:
            logger.debug(
                f"PDF Row {row_num}: Creating incomplete item (name only) - expecting quantity in next row")
            item_key = self._create_item_key_from_fields(raw_fields)
            if not item_key:
                return "skipped"

            return TenderItem(
                item_key=item_key,
                raw_fields=raw_fields,
                quantity=0.0,  # Will be updated when quantity row is found
                unit=unit,
                source="PDF",
                page_number=page_num + 1  # Convert 0-based to 1-based page number
            )

        # Case 2: Row has quantity data but no item fields (completion row for spanning)
        elif has_quantity_data and not has_item_fields:
            logger.debug(
                f"PDF Row {row_num}: Found completion row with quantity data")
            return self._complete_previous_item_with_quantity_data(existing_items, raw_fields, quantity)

        # Case 3: Row has both item fields and quantity data (complete item)
        elif has_item_fields and has_quantity_data:
            logger.debug(f"PDF Row {row_num}: Creating complete item")
            item_key = self._create_item_key_from_fields(raw_fields)
            if not item_key:
                return "skipped"

            return TenderItem(
                item_key=item_key,
                raw_fields=raw_fields,
                quantity=quantity,
                unit=unit,
                source="PDF",
                page_number=page_num + 1  # Convert 0-based to 1-based page number
            )

        # Case 4: Row has neither meaningful item fields nor quantity data
        else:
            logger.debug(f"PDF Row {row_num}: Skipping - no meaningful data")
            return "skipped"

    def _has_item_identifying_fields(self, raw_fields: Dict[str, str]) -> bool:
        """
        Check if the row contains fields that identify an item (name, classification, specification).
        """
        identifying_fields = [
            "工事区分・工種・種別・細別",
            "規格",
            "摘要"
        ]

        for field in identifying_fields:
            if field in raw_fields and raw_fields[field] and raw_fields[field].strip():
                return True

        return False

    def _complete_previous_item_with_quantity_data(self, existing_items: List[TenderItem],
                                                   raw_fields: Dict[str, str], quantity: float) -> str:
        """
        Complete the previous incomplete item with quantity and unit data.
        """
        if not existing_items:
            logger.warning(
                "PDF: Found quantity completion row but no previous item to complete")
            return "skipped"

        last_item = existing_items[-1]

        # Check if the last item needs completion (has quantity = 0)
        if last_item.quantity > 0:
            logger.warning(
                f"PDF: Previous item '{last_item.item_key}' already has quantity {last_item.quantity}")
            return "skipped"

        # Update the last item with quantity and any additional fields
        old_quantity = last_item.quantity
        last_item.quantity = quantity

        # Update unit if available in completion row
        if "単位" in raw_fields and raw_fields["単位"] and raw_fields["単位"].strip():
            last_item.unit = raw_fields["単位"].strip()
            logger.debug(
                f"PDF: Updated unit to '{last_item.unit}' for item '{last_item.item_key}'")

        # Merge additional fields (like unit) from the completion row
        for field_name, field_value in raw_fields.items():
            if field_name not in last_item.raw_fields or not last_item.raw_fields[field_name]:
                last_item.raw_fields[field_name] = field_value
                logger.debug(
                    f"PDF: Added field '{field_name}' = '{field_value}' to item '{last_item.item_key}'")

        logger.info(
            f"PDF row spanning completed: '{last_item.item_key}' quantity {old_quantity} -> {quantity}")
        if "単位" in raw_fields:
            logger.info(
                f"PDF row spanning: Added unit '{raw_fields['単位']}' to '{last_item.item_key}'")

        return "merged"

    def _is_completely_empty_row(self, row: List) -> bool:
        """
        Check if all cells in the row are empty or contain only whitespace.
        """
        if not row:
            return True

        for cell in row:
            if cell and str(cell).strip() and str(cell).strip() not in ["", "None", "nan", "NaN"]:
                return False

        return True

    def _is_quantity_only_row(self, raw_fields: Dict[str, str], quantity: float) -> bool:
        """
        Enhanced check for quantity-only rows (indicating row spanning).
        """
        # Must have quantity but no other meaningful fields
        if quantity <= 0:
            return False

        # Check if we have any meaningful non-quantity fields
        meaningful_fields = 0
        for field_name, field_value in raw_fields.items():
            if field_value and field_value.strip():
                meaningful_fields += 1

        # If we have quantity but no other fields, it's a spanned row
        is_quantity_only = meaningful_fields == 0

        if is_quantity_only:
            logger.debug(
                f"Detected quantity-only row: quantity={quantity}, fields={meaningful_fields}")

        return is_quantity_only

    def _merge_quantity_with_previous_item(self, existing_items: List[TenderItem], quantity: float) -> str:
        """
        Enhanced quantity merging with better error handling and logging.
        """
        if not existing_items:
            logger.warning(
                "Quantity-only row found but no previous item to merge with")
            return "skipped"

        # Get the last item for merging
        last_item = existing_items[-1]
        old_quantity = last_item.quantity
        new_quantity = old_quantity + quantity

        # Update the last item's quantity
        last_item.quantity = new_quantity

        logger.info(
            f"Row spanning merge: '{last_item.item_key}' quantity {old_quantity} + {quantity} = {new_quantity}")

        return "merged"

    def _create_item_key_from_fields(self, raw_fields: Dict[str, str]) -> str:
        """
        Create item key by concatenating item name with specification, then remove specification from raw_fields.
        """
        # Priority order for creating base item key
        base_key = ""

        # Use default fields to create base key
        key_fields = [
            "工事区分・工種・種別・細別",
            "摘要"
        ]

        # Get the base item name
        for field in key_fields:
            if field in raw_fields and raw_fields[field] and raw_fields[field].strip():
                base_key = raw_fields[field].strip()
                break

        # If no base key found, try other available fields
        if not base_key:
            for field_name, field_value in raw_fields.items():
                if field_value and field_value.strip() and field_name not in ["単位", "数量", "単価", "金額", "規格"]:
                    base_key = field_value.strip()
                    break

        # If still no key, return empty
        if not base_key:
            return ""

        # Check if specification exists and concatenate it with the base key
        if "規格" in raw_fields and raw_fields["規格"] and raw_fields["規格"].strip():
            specification = raw_fields["規格"].strip()
            concatenated_key = f"{base_key} + {specification}"

            # Remove specification from raw_fields since it's now part of the item key
            del raw_fields["規格"]

            logger.debug(
                f"PDF: Concatenated item key: '{base_key}' + '{specification}' = '{concatenated_key}'")
            logger.debug(f"PDF: Removed specification column from raw_fields")

            return concatenated_key

        # No specification to concatenate, return base key as is
        return base_key

    def _find_header_row(self, table: List[List]) -> Tuple[Optional[List], int]:
        """
        Find the header row in the table.
        Enhanced to search through multiple rows and use flexible pattern matching.
        """
        # Check up to 10 rows for headers (more flexible)
        for i, row in enumerate(table[:10]):
            if row:
                # Check for standard column indicators
                standard_indicators = ["名称", "工種", "数量", "単位", "単価", "金額"]
                has_standard = any(cell and any(indicator in str(cell)
                                   for indicator in standard_indicators) for cell in row)

                # Enhanced item name header detection using core words
                has_item_header = self._is_item_name_header_row(row)

                if has_standard or has_item_header:
                    logger.info(
                        f"Found header row at index {i} with indicators: standard={has_standard}, item_header={has_item_header}")
                    return row, i

        # If no clear header found, use first row
        logger.warning(
            "No clear header row found, using first row as fallback")
        return table[0] if table else None, 0

    def _is_item_name_header_row(self, row: List) -> bool:
        """
        Check if this row contains item name header based on core words.
        Removes special characters and looks for combinations of: 費目, 工種, 種別, 細目
        """
        if not row:
            return False

        core_words = ["費目", "工種", "種別", "細目"]

        for cell in row:
            if cell:
                # Remove all special characters and spaces for pattern matching
                cleaned_cell = self._clean_text_for_matching(str(cell))

                # Count how many core words are present in this cell
                word_count = sum(
                    1 for word in core_words if word in cleaned_cell)

                # If we have 2 or more core words, this is likely an item name header
                if word_count >= 2:
                    logger.info(
                        f"Item name header detected in cell: '{cell}' (cleaned: '{cleaned_cell}', words found: {word_count})")
                    return True

        return False

    def _clean_text_for_matching(self, text: str) -> str:
        """
        Clean text by removing special characters for flexible pattern matching.
        Keeps only Japanese characters and basic alphanumeric.
        """
        import re
        # Remove common separators and special characters, keep Japanese and alphanumeric
        cleaned = re.sub(r'[・・/\-\s\(\)（）\[\]【】「」『』\|｜\.。、，]', '', text)
        return cleaned

    def _get_column_mapping(self, header_row: List) -> Dict[str, int]:
        """
        Map column names to indices based on header row.
        Updated to handle the specific 8-column structure.
        """
        col_indices = {}

        # Check for all column types in the specific order
        column_names = [
            "工事区分・工種・種別・細別",
            "規格",
            "単位",
            "数量",
            "単価",
            "金額",
            "数量・金額増減",
            "摘要"
        ]

        for col_name in column_names:
            idx = self._find_column_index(header_row, col_name)
            if idx != -1:
                col_indices[col_name] = idx

        return col_indices

    def _find_column_index(self, header_row: List[str], column_name: str) -> int:
        """
        Find column index by matching patterns flexibly
        """
        if not header_row:
            return -1

        patterns = self.column_patterns.get(column_name, [column_name])

        for i, cell in enumerate(header_row):
            if not cell:
                continue

            cell_clean = str(cell).strip()
            for pattern in patterns:
                if pattern in cell_clean:
                    return i
        return -1

    def _extract_quantity(self, cell_value) -> float:
        """
        Extract numeric quantity from cell value
        """
        if not cell_value:
            return 0.0

        # Clean the value - remove commas, spaces, and convert to string
        value_str = str(cell_value).replace(
            ",", "").replace(" ", "").replace("　", "")

        # Try to extract number using regex
        number_match = re.search(r'[\d.]+', value_str)
        if number_match:
            try:
                return float(number_match.group())
            except ValueError:
                pass
        return 0.0

    def _process_single_row(self, row: List, col_indices: Dict[str, int],
                            page_num: int, table_num: int, row_num: int) -> Optional[TenderItem]:
        """
        Legacy method - kept for backward compatibility but now uses new spanning logic.
        """
        # Use new spanning logic with empty list (no previous items to merge with)
        result = self._process_single_row_with_spanning(
            row, col_indices, page_num, table_num, row_num, [])

        if isinstance(result, TenderItem):
            return result
        else:
            return None

    def extract_subtables_with_range(self, pdf_path: str, start_page: Optional[int] = None, end_page: Optional[int] = None, reference_numbers: Optional[List[str]] = None) -> List[SubtableItem]:
        """
        Extract subtables from PDF with specified page range using reference numbers from main table.

        Subtable extraction requirements:
        1. Use reference numbers from main table's 摘要 column
        2. Look for pattern: reference number → column headers → subtable data
        3. Column headers: 名称・規格, 単位, 数量, 摘要
        4. Only extract actual subtable data, not other elements

        Args:
            pdf_path: Path to the PDF file
            start_page: Starting page number (1-based, None means start from page 1)
            end_page: Ending page number (1-based, None means extract all pages)
            reference_numbers: List of reference numbers to look for (from main table's 摘要 column)
        """
        all_subtable_items = []

        logger.info(f"Starting PDF subtable extraction from: {pdf_path}")
        logger.info(
            f"Subtable page range: {start_page or 'start'} to {end_page or 'end'}")
        logger.info(f"Looking for reference numbers: {reference_numbers}")

        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                logger.info(f"PDF has {total_pages} pages total")

                # Use provided reference numbers or discover from main table
                if reference_numbers:
                    target_references = reference_numbers
                    logger.info(
                        f"Using provided reference numbers: {target_references}")
                else:
                    logger.info(
                        "No reference numbers provided, discovering from main table...")
                    target_references = self._discover_reference_patterns_from_main_table(
                        pdf)
                    if not target_references:
                        logger.warning(
                            "No reference patterns discovered from main table, using fallback patterns")
                        target_references = ['単', '道', '内']

                # Determine actual page range
                actual_start = (
                    start_page - 1) if start_page is not None else 0
                actual_end = (
                    end_page - 1) if end_page is not None else total_pages - 1

                # Validate page range
                actual_start = max(0, actual_start)
                actual_end = min(total_pages - 1, actual_end)

                if actual_start > actual_end:
                    logger.warning(
                        f"Invalid page range: start={actual_start+1}, end={actual_end+1}")
                    return all_subtable_items

                pages_to_process = actual_end - actual_start + 1
                logger.info(
                    f"Processing pages {actual_start + 1} to {actual_end + 1} ({pages_to_process} pages) for subtables")

                # Process specified page range for subtables
                for page_num in range(actual_start, actual_end + 1):
                    page = pdf.pages[page_num]
                    logger.info(
                        f"Processing page {page_num + 1}/{total_pages} for subtables")

                    page_subtables = self._extract_subtables_from_page_improved(
                        page, page_num, target_references)

                    logger.info(
                        f"Extracted {len(page_subtables)} subtable items from page {page_num + 1}")
                    all_subtable_items.extend(page_subtables)

                logger.info(
                    f"Total subtable items extracted from PDF (pages {actual_start + 1}-{actual_end + 1}): {len(all_subtable_items)}")

        except Exception as e:
            logger.error(f"Error processing PDF for subtables: {e}")
            raise

        return all_subtable_items

    def _extract_subtables_from_page_improved(self, page, page_num: int, reference_numbers: List[str]) -> List[SubtableItem]:
        """
        Simplified subtable extraction that works with the actual PDF structure.

        Pattern we're looking for:
        - Row 1: 【第 X 号 明細書】 (reference)
        - Row 2: 名称・規格 数量 単位... (headers)
        - Row 3+: actual data rows
        """
        page_subtable_items = []

        try:
            tables = page.extract_tables()

            if not tables:
                logger.info(
                    f"No tables found on page {page_num + 1} for subtables")
                return page_subtable_items

            logger.info(
                f"Found {len(tables)} tables on page {page_num + 1} for subtable extraction")

            # Process each table
            for table_num, table in enumerate(tables):
                # Need at least: reference, header, data
                if not table or len(table) < 3:
                    continue

                logger.info(
                    f"Processing table {table_num + 1}/{len(tables)} on page {page_num + 1}")

                # Look for subtables in this table using simple pattern matching
                for i in range(len(table) - 2):  # -2 because we need at least 3 rows
                    row = table[i]
                    if not row:
                        continue

                    # Check if this row contains a reference number
                    reference_found = self._extract_reference_number_simple(
                        row)
                    if not reference_found:
                        continue

                    logger.info(
                        f"Found reference '{reference_found}' at row {i + 1}")

                    # Check if the next row contains headers
                    if i + 1 < len(table):
                        header_row = table[i + 1]
                        if self._is_subtable_header_row_simple(header_row):
                            logger.info(f"Found headers at row {i + 2}")

                            # Extract data rows starting from row i+2
                            data_start_idx = i + 2
                            data_rows = []

                            # Collect data rows until we hit another reference or end of table
                            for j in range(data_start_idx, len(table)):
                                data_row = table[j]
                                if not data_row:
                                    continue

                                # Stop if we find another reference
                                if self._extract_reference_number_simple(data_row):
                                    break

                                # Stop if we find summary rows
                                if self._is_summary_row(data_row):
                                    break

                                # Add this data row
                                if self._has_meaningful_data(data_row):
                                    data_rows.append(data_row)

                            logger.info(
                                f"Found {len(data_rows)} data rows for reference '{reference_found}'")

                            # Create SubtableItems from data rows using multi-row logic
                            subtable_items = self._create_subtable_items_from_multirow_data(
                                data_rows, reference_found, page_num)
                            page_subtable_items.extend(subtable_items)

                            logger.info(
                                f"Created {len([item for item in page_subtable_items if item.reference_number == reference_found])} items for reference '{reference_found}'")

                logger.info(
                    f"Extracted {len(page_subtable_items)} subtable items from table {table_num + 1}")

        except Exception as e:
            logger.error(
                f"Error processing page {page_num + 1} for subtables: {e}")

        return page_subtable_items

    def _extract_reference_number_simple(self, row: List) -> Optional[str]:
        """
        Simple reference number extraction from a row.
        Looks for patterns like "内 1号", "内 2号", ..., "内 82号" etc.
        Ignores spaces and handles full-width/half-width characters.
        """
        for cell in row:
            if not cell:
                continue
            
            cell_str = str(cell).strip()
            
            # Remove all spaces and normalize full-width/half-width characters
            normalized = cell_str.replace(' ', '').replace('　', '').replace('１', '1').replace('２', '2').replace('３', '3').replace('４', '4').replace('５', '5').replace('６', '6').replace('７', '7').replace('８', '8').replace('９', '9').replace('０', '0')
            
            # Pattern: kanji + number + 号 (e.g., "内1号", "内2号", ..., "内82号")
            match = re.search(r'([内単道諸雑材工機労共設備運管電水土建構橋舗])\s*(\d+)\s*号', normalized)
            if match:
                kanji = match.group(1)
                number = match.group(2)
                return f"{kanji} {number}号"
        
        return None

    def _is_subtable_header_row_simple(self, row: List) -> bool:
        """
        Simple check for subtable header row.
        """
        if not row:
            return False

        row_text = ' '.join(str(cell) if cell else '' for cell in row)

        # Remove spaces for flexible matching
        row_text_clean = row_text.replace(' ', '').replace('　', '')

        # Must contain key headers (with flexible spacing)
        required_headers = ['名称', '単位', '数量']
        found_count = sum(
            1 for header in required_headers if header in row_text_clean)

        return found_count >= 2

    def _is_summary_row(self, row: List) -> bool:
        """
        Check if this is a summary row (計, 合計, etc.)
        """
        if not row:
            return False

        row_text = ' '.join(str(cell) if cell else '' for cell in row)
        summary_keywords = ['計', '合計', '小計', '総計']

        return any(keyword in row_text for keyword in summary_keywords)

    def _has_meaningful_data(self, row: List) -> bool:
        """
        Check if row has meaningful data (not just empty cells).
        For PDF subtables, even a single meaningful cell can be important.
        """
        if not row:
            return False

        # Count non-empty cells
        non_empty_count = sum(1 for cell in row if cell and str(cell).strip())

        # Must have at least 1 non-empty cell (relaxed from 2)
        return non_empty_count >= 1

    def _create_subtable_item_from_row(self, row: List, reference_number: str, page_num: int) -> Optional[SubtableItem]:
        """
        Create a SubtableItem from a data row.
        Expected columns: 名称・規格, 条件, 単位, 数量, 単価, 金額, 数量・金額増減, 摘要
        """
        try:
            # Extract data based on expected column positions
            name = ""
            unit = ""
            quantity = 0.0

            # Clean the row - remove None values and convert to strings
            clean_row = [str(cell).strip()
                         if cell is not None else "" for cell in row]

            # Column mapping based on the header structure we saw:
            # 名称・規格(0), 条件(1), 単位(2), 数量(3), 単価(4), 金額(5), 数量・金額増減(6), 摘要(7)

            # Extract name (column 0)
            if len(clean_row) > 0 and clean_row[0]:
                name = clean_row[0]

            # Extract unit (column 2)
            if len(clean_row) > 2 and clean_row[2]:
                unit = clean_row[2]

            # Extract quantity (column 3)
            if len(clean_row) > 3 and clean_row[3]:
                try:
                    quantity_str = clean_row[3].replace(
                        ',', '').replace('、', '')
                    if quantity_str and (quantity_str.replace('.', '').isdigit()):
                        quantity = float(quantity_str)
                except (ValueError, TypeError):
                    pass

            # Skip if no meaningful name
            if not name or len(name.strip()) == 0:
                return None

            # Create the SubtableItem
            return SubtableItem(
                item_key=name,
                raw_fields={
                    "名称・規格": name,
                    "単位": unit,
                    "数量": str(quantity),
                    "条件": clean_row[1] if len(clean_row) > 1 else "",
                    "単価": clean_row[4] if len(clean_row) > 4 else "",
                    "金額": clean_row[5] if len(clean_row) > 5 else "",
                    "摘要": clean_row[7] if len(clean_row) > 7 else ""
                },
                quantity=quantity,
                unit=unit,
                source="PDF",
                page_number=page_num + 1,
                reference_number=reference_number,
                sheet_name=None
            )

        except Exception as e:
            logger.error(f"Error creating subtable item from row: {e}")
            return None

    def _create_subtable_items_from_multirow_data(self, data_rows: List[List], reference_number: str, page_num: int) -> List[SubtableItem]:
        """
        Create SubtableItems from multi-row data where each item spans multiple rows.
        
        Pattern observed:
        - Row 1: Item name in column 0
        - Row 2: Empty row
        - Row 3: Unit in column 3, quantity in column 4
        """
        items = []
        i = 0
        
        while i < len(data_rows):
            try:
                # Look for item name row
                name_row = data_rows[i]
                clean_name_row = [str(cell).strip() if cell is not None else "" for cell in name_row]
                
                # Check if this row has an item name (column 0)
                if len(clean_name_row) > 0 and clean_name_row[0]:
                    item_name = clean_name_row[0]
                    
                    # Look for unit/quantity in the next few rows
                    unit = ""
                    quantity = 0.0
                    
                    # Check next 3 rows for unit and quantity data
                    for j in range(i + 1, min(i + 4, len(data_rows))):
                        if j < len(data_rows):
                            data_row = data_rows[j]
                            clean_data_row = [str(cell).strip() if cell is not None else "" for cell in data_row]
                            
                            # Look for unit in column 3 and quantity in column 4
                            if len(clean_data_row) > 3 and clean_data_row[3]:
                                unit = clean_data_row[3]
                            
                            if len(clean_data_row) > 4 and clean_data_row[4]:
                                try:
                                    quantity_str = clean_data_row[4].replace(',', '').replace('、', '')
                                    if quantity_str and (quantity_str.replace('.', '').isdigit()):
                                        quantity = float(quantity_str)
                                except (ValueError, TypeError):
                                    pass
                            
                            # If we found both unit and quantity, we can create the item
                            if unit and quantity > 0:
                                break
                    
                    # Create the SubtableItem
                    if item_name:
                        item = SubtableItem(
                            item_key=item_name,
                            raw_fields={
                                "名称・規格": item_name,
                                "単位": unit,
                                "数量": str(quantity)
                            },
                            quantity=quantity,
                            unit=unit,
                            source="PDF",
                            page_number=page_num + 1,
                            reference_number=reference_number,
                            sheet_name=None
                        )
                        items.append(item)
                        logger.debug(f"Created item: {item_name[:30]}... (unit: {unit}, qty: {quantity})")
                
                i += 1
                
            except Exception as e:
                logger.error(f"Error processing multi-row data at index {i}: {e}")
                i += 1
        
        return items

    def _extract_subtables_from_table_improved(self, table: List[List], page_num: int, table_num: int, reference_numbers: List[str]) -> List[SubtableItem]:
        """
        Extract subtables from a single table using the improved logic.
        Look for: reference number → column headers → subtable data
        """
        subtable_items = []
        i = 0

        while i < len(table):
            row = table[i]
            if not row:
                i += 1
                continue

            # Step 1: Look for reference number
            found_reference = self._find_reference_in_row(
                row, reference_numbers)
            if not found_reference:
                i += 1
                continue

            logger.info(
                f"Found reference '{found_reference}' at row {i + 1} in table {table_num + 1}")

            # Step 2: Look for column headers in the next few rows
            header_row_idx = None
            col_mapping = None

            # Check next 5 rows for headers
            for j in range(i + 1, min(i + 6, len(table))):
                if self._is_subtable_header_row_improved(table[j]):
                    header_row_idx = j
                    col_mapping = self._get_subtable_column_mapping_improved(
                        table[j])
                    logger.info(
                        f"Found subtable headers at row {j + 1}: {col_mapping}")
                    break

            if not header_row_idx or not col_mapping:
                logger.warning(
                    f"No valid headers found after reference '{found_reference}'")
                i += 1
                continue

            # Step 3: Extract subtable data rows
            data_start = header_row_idx + 1
            data_end = self._find_subtable_data_end(
                table, data_start, reference_numbers)

            if data_start < len(table):
                data_rows = table[data_start:data_end]
                logger.info(
                    f"Processing {len(data_rows)} data rows for reference '{found_reference}'")

                # Process data rows with spanning logic
                subtable_data = self._process_subtable_data_rows_improved(
                    data_rows, col_mapping, page_num, found_reference, data_start)

                subtable_items.extend(subtable_data)
                logger.info(
                    f"Extracted {len(subtable_data)} items for reference '{found_reference}'")

            # Move to the end of this subtable
            i = data_end

        return subtable_items

    def _find_reference_in_row(self, row: List, reference_numbers: List[str]) -> Optional[str]:
        """
        Find a reference number from the provided list in the given row.
        """
        for cell in row:
            if not cell:
                continue

            cell_str = str(cell).strip()

            # Check if this cell contains any of our target reference numbers
            for ref_num in reference_numbers:
                if ref_num in cell_str:
                    # Extract the full reference (e.g., "単1号", "単 1号")
                    import re
                    # Look for patterns like "単1号", "単 1号", etc.
                    pattern = r'([一-龯A-Za-z]*)\s*(\d+)\s*号'
                    match = re.search(pattern, cell_str)
                    if match:
                        prefix = match.group(1).strip()
                        number = match.group(2)
                        return f"{prefix}{number}号"

        return None

    def _is_subtable_header_row_improved(self, row: List) -> bool:
        """
        Check if a row contains subtable column headers.
        Looking for: 名称・規格, 単位, 数量, 摘要
        """
        if not row:
            return False

        # Required headers for subtables
        required_headers = ["名称", "単位", "数量"]
        found_headers = 0

        for cell in row:
            if not cell:
                continue

            cell_str = str(cell).strip()

            # Check for each required header
            for header in required_headers:
                if header in cell_str:
                    found_headers += 1
                    break

        # Need at least 2 of the 3 required headers
        return found_headers >= 2

    def _get_subtable_column_mapping_improved(self, header_row: List) -> Dict[str, int]:
        """
        Get column mapping for subtable headers.
        Looking for: 名称・規格, 単位, 数量, 摘要
        """
        col_mapping = {}

        for i, cell in enumerate(header_row):
            if not cell:
                continue

            cell_str = str(cell).strip()

            # Map columns based on content
            if "名称" in cell_str or "規格" in cell_str:
                col_mapping["名称・規格"] = i
            elif "単位" in cell_str:
                col_mapping["単位"] = i
            elif "数量" in cell_str:
                col_mapping["数量"] = i
            elif "摘要" in cell_str or "備考" in cell_str:
                col_mapping["摘要"] = i

        return col_mapping

    def _find_subtable_data_end(self, table: List[List], start_idx: int, reference_numbers: List[str]) -> int:
        """
        Find where the subtable data ends.
        Stops at: next reference number, summary row, or end of table.
        """
        for i in range(start_idx, len(table)):
            row = table[i]
            if not row:
                continue

            # Check if this row contains a new reference number
            if self._find_reference_in_row(row, reference_numbers):
                return i

            # Check for summary/total rows
            for cell in row:
                if not cell:
                    continue
                cell_str = str(cell).strip()
                if any(word in cell_str for word in ["合計", "小計", "総計", "計"]):
                    return i

        return len(table)

    def _process_subtable_data_rows_improved(self, data_rows: List[List], col_mapping: Dict[str, int],
                                             page_num: int, reference_number: str, start_row_num: int) -> List[SubtableItem]:
        """
        Process subtable data rows with improved row spanning logic.
        """
        subtable_items = []
        i = 0

        while i < len(data_rows):
            row = data_rows[i]

            if not row or self._is_empty_row(row):
                i += 1
                continue

            # Extract item data with spanning logic
            item_data = self._extract_subtable_item_improved(
                data_rows, i, col_mapping, page_num, reference_number, start_row_num + i)

            if item_data:
                subtable_items.append(item_data['item'])
                i = item_data['next_index']
            else:
                i += 1

        return subtable_items

    def _extract_subtable_item_improved(self, data_rows: List[List], start_idx: int,
                                        col_mapping: Dict[str, int], page_num: int,
                                        reference_number: str, row_num: int) -> Optional[Dict]:
        """
        Extract a single subtable item with row spanning logic.
        """
        if start_idx >= len(data_rows):
            return None

        # Collect data from current and next few rows
        item_name = None
        quantity = 0.0
        unit = None
        remarks = None
        raw_fields = {}
        rows_processed = 1

        # Look ahead up to 3 rows for spanning data
        for offset in range(min(3, len(data_rows) - start_idx)):
            row = data_rows[start_idx + offset]

            if not row:
                continue

            # Extract data from this row
            for col_name, col_idx in col_mapping.items():
                if col_idx < len(row) and row[col_idx]:
                    cell_value = str(row[col_idx]).strip()
                    if cell_value:
                        if col_name == "名称・規格" and not item_name:
                            item_name = cell_value
                            raw_fields[col_name] = cell_value
                        elif col_name == "数量" and quantity == 0.0:
                            quantity = self._extract_quantity(cell_value)
                            raw_fields[col_name] = cell_value
                        elif col_name == "単位" and not unit:
                            unit = cell_value
                            raw_fields[col_name] = cell_value
                        elif col_name == "摘要" and not remarks:
                            remarks = cell_value
                            raw_fields[col_name] = cell_value

            # Stop if we have complete data
            if item_name and (quantity > 0 or unit):
                rows_processed = offset + 1
                break

        # Validate the extracted item
        if not item_name:
            return None

        # Skip non-meaningful items
        if any(word in item_name for word in ["合計", "小計", "総計"]):
            return None

        # Create the subtable item
        item_key = item_name

        subtable_item = SubtableItem(
            item_key=item_key,
            raw_fields=raw_fields,
            quantity=quantity,
            unit=unit or "",
            source="PDF",
            page_number=page_num + 1,
            reference_number=reference_number
        )

        return {
            'item': subtable_item,
            'next_index': start_idx + rows_processed
        }

    def _is_empty_row(self, row: List) -> bool:
        """
        Check if a row is empty or contains only whitespace.
        """
        for cell in row:
            if cell and str(cell).strip():
                return False
        return True

    def _discover_reference_patterns_from_main_table(self, pdf) -> List[str]:
        """
        Dynamically discover all reference patterns from the main table's 摘要 column.
        This makes the extraction work with any PDF file, not just hardcoded patterns.

        Args:
            pdf: Opened PDF file object

        Returns:
            List of discovered reference prefixes (e.g., ['単', '内', '道', '諸'])
        """
        discovered_patterns = set()

        # Assume main table is on pages 4-12 (this could be made configurable)
        main_table_start = 3  # 0-based page index for page 4
        main_table_end = 11   # 0-based page index for page 12

        logger.info(
            f"Scanning main table pages {main_table_start + 1} to {main_table_end + 1} for reference patterns")

        for page_num in range(main_table_start, min(main_table_end + 1, len(pdf.pages))):
            page = pdf.pages[page_num]
            tables = page.extract_tables()

            if not tables:
                continue

            for table in tables:
                if not table:
                    continue

                # Look for 摘要 column and extract reference patterns
                for row in table:
                    if not row:
                        continue

                    for cell in row:
                        if not cell:
                            continue

                        cell_str = str(cell).strip()

                        # Skip very long text to avoid false positives
                        if len(cell_str) > 100:
                            continue

                        # Find all reference patterns in the cell
                        patterns = self._extract_all_reference_patterns_from_text(
                            cell_str)
                        discovered_patterns.update(patterns)

        # Convert to sorted list for consistent ordering
        pattern_list = sorted(list(discovered_patterns))

        logger.info(
            f"Discovered {len(pattern_list)} unique reference patterns: {pattern_list}")

        return pattern_list

    def _extract_all_reference_patterns_from_text(self, text: str) -> List[str]:
        """
        Extract all possible reference patterns from a text string.

        Args:
            text: Text to search for reference patterns

        Returns:
            List of reference prefixes found (e.g., ['単', '内'])
        """
        patterns = []

        # Comprehensive regex to match various Japanese reference patterns
        # Matches: [Japanese character][optional space][number][optional space]号
        reference_regex = r'([一-龯\w])\s*(\d+)\s*号'

        matches = re.findall(reference_regex, text)

        for match in matches:
            if len(match) >= 2:
                prefix = match[0]
                number = match[1]

                # Validate that it looks like a reference pattern
                if self._is_valid_reference_prefix(prefix) and number.isdigit():
                    patterns.append(prefix)

        return patterns

    def _is_valid_reference_prefix(self, prefix: str) -> bool:
        """
        Check if a prefix is a valid reference prefix.

        Args:
            prefix: Potential reference prefix

        Returns:
            True if it's a valid reference prefix
        """
        # Single character prefixes are most common
        if len(prefix) != 1:
            return False

        # Common Japanese reference prefixes
        # This list can be expanded based on observations
        common_prefixes = {
            '単', '内', '道', '諸', '雑', '材', '工', '機', '労', '共', '設',
            '備', '運', '管', '設', '電', '水', '土', '建', '構', '橋', '舗'
        }

        # Allow common prefixes or single Japanese characters
        return prefix in common_prefixes or '\u3040' <= prefix <= '\u309f' or '\u30a0' <= prefix <= '\u30ff' or '\u4e00' <= prefix <= '\u9faf'

    def _extract_subtables_from_page(self, page, page_num: int, discovered_patterns: List[str]) -> List[SubtableItem]:
        """
        Extract subtables from a single page with specific subtable logic.
        Updated to use dynamically discovered patterns.
        """
        page_subtable_items = []
        current_reference_number = None

        try:
            tables = page.extract_tables()

            if not tables:
                logger.info(
                    f"No tables found on page {page_num + 1} for subtables")
                return page_subtable_items

            logger.info(
                f"Found {len(tables)} tables on page {page_num + 1} for subtable extraction")

            # Process each table for subtables
            for table_num, table in enumerate(tables):
                logger.info(
                    f"Processing table {table_num + 1}/{len(tables)} on page {page_num + 1} for subtables")

                # Look for reference numbers and subtable data in this table
                table_subtables, reference_number = self._process_single_table_for_subtables(
                    table, page_num, table_num, current_reference_number, discovered_patterns)

                # Update current reference number if found
                if reference_number:
                    current_reference_number = reference_number

                logger.info(
                    f"Extracted {len(table_subtables)} subtable items from table {table_num + 1}")

                page_subtable_items.extend(table_subtables)

        except Exception as e:
            logger.error(
                f"Error processing page {page_num + 1} for subtables: {e}")

        return page_subtable_items

    def _process_single_table_for_subtables(self, table: List[List], page_num: int, table_num: int, current_reference: Optional[str], discovered_patterns: List[str]) -> Tuple[List[SubtableItem], Optional[str]]:
        """
        Process a single table for subtable extraction with reference number detection.

        Returns:
            Tuple of (subtable_items, detected_reference_number)
        """
        subtable_items = []
        detected_reference = current_reference

        if not table or len(table) < 2:
            logger.warning(
                f"Table {table_num + 1} on page {page_num + 1} is too small for subtable processing")
            return subtable_items, detected_reference

        # Subtable-specific column patterns - updated based on actual PDF structure
        subtable_column_patterns = {
            "名称・規格": ["名称・規格", "名称", "規格", "名 称 ・ 規 格", "名 称", "規 格"],
            "条件": ["条件", "条 件"],
            "単位": ["単位", "単 位"],
            "数量": ["数量", "数 量"],
            "単価": ["単価", "単 価"],
            "金額": ["金額", "金 額"],
            "数量・金額増減": ["数量・金額増減", "増減", "数量増減", "金額増減"],
            "摘要": ["摘要", "備考", "摘 要", "備 考"]
        }

        # Process each row to find reference numbers and subtable data
        for row_idx, row in enumerate(table):
            if not row:
                continue

            # Check for reference numbers like "単 3号"
            reference_found = self._extract_reference_number(
                row, discovered_patterns)
            if reference_found:
                detected_reference = reference_found
                logger.info(
                    f"Found reference number '{detected_reference}' in table {table_num + 1}, row {row_idx + 1}")
                continue

            # Check if this row is a subtable header
            if self._is_subtable_header_row(row, subtable_column_patterns):
                logger.info(
                    f"Found subtable header at row {row_idx + 1} in table {table_num + 1}")
                # Process data rows following this header
                data_rows = table[row_idx + 1:]
                col_mapping = self._get_subtable_column_mapping(
                    row, subtable_column_patterns)

                if col_mapping:
                    # Process data rows with row spanning logic for subtables
                    subtable_items_from_rows = self._process_subtable_data_rows_with_spanning(
                        data_rows, col_mapping, page_num, detected_reference, row_idx + 1
                    )
                    subtable_items.extend(subtable_items_from_rows)

                break  # Stop processing this table after finding subtable header

        return subtable_items, detected_reference

    def _extract_reference_number(self, row: List, discovered_patterns: List[str]) -> Optional[str]:
        """
        Extract reference numbers from a row using both discovered patterns and common PDF patterns.
        This works with any PDF file by using patterns found in the main table plus common patterns.

        Args:
            row: Table row to search
            discovered_patterns: List of reference prefixes found in main table

        Returns:
            Reference number string (e.g., "第 1 号") or None
        """
        for cell in row:
            if not cell:
                continue

            cell_str = str(cell).strip()

            # Pattern 1: 【第 X 号 明細書】 format
            match = re.search(r'【第\s*(\d+)\s*号\s*明細書】', cell_str)
            if match:
                number = match.group(1)
                return f"第 {number} 号"

            # Pattern 2: 第 X 号 明細書 format (without brackets)
            match = re.search(r'第\s*(\d+)\s*号\s*明細書', cell_str)
            if match:
                number = match.group(1)
                return f"第 {number} 号"

            # Pattern 3: Ｐ X 号 format
            match = re.search(r'Ｐ\s*(\d+)\s*号', cell_str)
            if match:
                number = match.group(1)
                return f"Ｐ {number} 号"

            # Pattern 4: 単X号 format (original pattern)
            match = re.search(r'単\s*(\d+)\s*号', cell_str)
            if match:
                number = match.group(1)
                return f"単 {number} 号"

            # Pattern 5: Use discovered patterns if provided
            if discovered_patterns:
                escaped_patterns = [re.escape(pattern)
                                    for pattern in discovered_patterns]
                pattern_group = '|'.join(escaped_patterns)

                # Main pattern: (discovered_prefix) + optional space + number + 号
                reference_pattern = f'({pattern_group})\\s*(\\d+)\\s*号'
                match = re.search(reference_pattern, cell_str)

                if match:
                    prefix = match.group(1)
                    number = match.group(2)
                    return f"{prefix} {number}号"

        return None

    def _is_subtable_header_row(self, row: List, column_patterns: Dict[str, List[str]]) -> bool:
        """
        Check if a row is a subtable header based on the specific column patterns.
        """
        if not row:
            return False

        # Check if row contains required subtable column headers
        required_patterns = ["名称", "単位", "数量"]  # Core required columns
        found_patterns = 0

        for cell in row:
            if not cell:
                continue

            cell_str = str(cell).strip()

            for pattern in required_patterns:
                if pattern in cell_str:
                    found_patterns += 1
                    break

        # Consider it a header if we find at least 2 of the 3 required patterns
        return found_patterns >= 2

    def _get_subtable_column_mapping(self, header_row: List, column_patterns: Dict[str, List[str]]) -> Dict[str, int]:
        """
        Map subtable column names to indices based on header row.
        """
        col_indices = {}

        for col_name, patterns in column_patterns.items():
            for i, cell in enumerate(header_row):
                if not cell:
                    continue

                cell_str = str(cell).strip()

                for pattern in patterns:
                    if pattern in cell_str:
                        col_indices[col_name] = i
                        break

                if col_name in col_indices:
                    break

        return col_indices

    def _is_subtable_end_row(self, row: List) -> bool:
        """
        Check if a row indicates the end of a subtable.
        """
        if not row:
            return True

        # Check for typical end indicators
        for cell in row:
            if not cell:
                continue

            cell_str = str(cell).strip()

            # End indicators
            end_patterns = ["合計", "小計", "総計", "計"]

            for pattern in end_patterns:
                if pattern in cell_str and len(cell_str) <= 10:
                    return True

        return False

    def _process_subtable_data_row(self, row: List, col_mapping: Dict[str, int],
                                   page_num: int, reference_number: Optional[str],
                                   row_num: int) -> Optional[SubtableItem]:
        """
        Process a single subtable data row and create SubtableItem if valid.
        """
        if not row:
            return None

        # Extract fields based on column mapping
        raw_fields = {}
        quantity = 0.0
        unit = None
        item_name = None

        for col_name, col_idx in col_mapping.items():
            if col_idx < len(row) and row[col_idx]:
                cell_value = str(row[col_idx]).strip()
                if cell_value:
                    if col_name == "数量":
                        quantity = self._extract_quantity(cell_value)
                    elif col_name == "単位":
                        unit = cell_value
                    elif col_name == "名称・規格":
                        item_name = cell_value

                    raw_fields[col_name] = cell_value

        # Apply subtable-specific filtering rules
        # Must have item name
        if not item_name:
            logger.debug(f"Skipping row {row_num}: missing item name")
            return None

        # Accept items with either:
        # 1. Valid quantity (> 0), OR
        # 2. Item name + unit (even if quantity is 0 or missing)
        has_valid_quantity = quantity > 0
        has_name_and_unit = item_name and unit

        if not (has_valid_quantity or has_name_and_unit):
            logger.debug(
                f"Skipping row {row_num}: no valid quantity and no unit")
            return None

        # Ignore rows with only "合計" and "単価" without meaningful data
        if ("合計" in item_name or "単価" in item_name) and not has_valid_quantity and not unit:
            logger.debug(
                f"Skipping row {row_num}: contains '合計' or '単価' without meaningful data")
            return None

        # Create item key
        item_key = self._create_subtable_item_key(raw_fields, reference_number)

        if not item_key:
            return None

        return SubtableItem(
            item_key=item_key,
            raw_fields=raw_fields,
            quantity=quantity,
            unit=unit,
            source="PDF",
            page_number=page_num + 1,  # Convert 0-based to 1-based
            reference_number=reference_number
        )

    def _process_subtable_data_rows_with_spanning(self, data_rows: List[List], col_mapping: Dict[str, int],
                                                  page_num: int, reference_number: Optional[str],
                                                  start_row_num: int) -> List[SubtableItem]:
        """
        Process subtable data rows with row spanning logic.

        In subtable structure, data often spans multiple rows:
        - Row 1: Item name only
        - Row 2: Empty
        - Row 3: Unit and quantity
        """
        subtable_items = []
        i = 0

        while i < len(data_rows):
            row = data_rows[i]

            if not row:
                i += 1
                continue

            # Check if this is the end of the subtable
            if self._is_subtable_end_row(row):
                break

            # Try to process this row as a potential subtable item
            item_data = self._extract_subtable_item_data_with_spanning(
                data_rows, i, col_mapping, page_num, reference_number, start_row_num + i
            )

            if item_data:
                subtable_items.append(item_data['item'])
                i = item_data['next_index']  # Skip processed rows
            else:
                i += 1

        return subtable_items

    def _extract_subtable_item_data_with_spanning(self, data_rows: List[List], start_idx: int,
                                                  col_mapping: Dict[str, int], page_num: int,
                                                  reference_number: Optional[str], row_num: int) -> Optional[Dict]:
        """
        Extract subtable item data with row spanning logic.

        Returns dict with 'item' and 'next_index' or None if no valid item found.
        """
        if start_idx >= len(data_rows):
            return None

        # Look ahead up to 3 rows to gather complete item information
        item_name = None
        quantity = 0.0
        unit = None
        raw_fields = {}
        rows_processed = 1

        # Check the next few rows for spanning data
        for offset in range(min(3, len(data_rows) - start_idx)):
            row = data_rows[start_idx + offset]

            if not row:
                continue

            # Extract data from this row
            for col_name, col_idx in col_mapping.items():
                if col_idx < len(row) and row[col_idx]:
                    cell_value = str(row[col_idx]).strip()
                    if cell_value:
                        if col_name == "名称・規格" and not item_name:
                            item_name = cell_value
                            raw_fields[col_name] = cell_value
                        elif col_name == "数量" and quantity == 0.0:
                            quantity = self._extract_quantity(cell_value)
                            raw_fields[col_name] = cell_value
                        elif col_name == "単位" and not unit:
                            unit = cell_value
                            raw_fields[col_name] = cell_value
                        elif col_name not in raw_fields:  # Store other fields
                            raw_fields[col_name] = cell_value

            # If we find a new item name on a later row, stop here
            if offset > 0 and col_mapping.get("名称・規格", -1) < len(row):
                cell_value = row[col_mapping["名称・規格"]
                                 ] if col_mapping["名称・規格"] < len(row) else None
                if cell_value and str(cell_value).strip() and item_name and str(cell_value).strip() != item_name:
                    # This is a new item, don't include this row
                    break

            rows_processed = offset + 1

        # Apply subtable filtering rules
        # Must have item name
        if not item_name:
            return None

        # Accept items with either:
        # 1. Valid quantity (> 0), OR
        # 2. Item name + unit (even if quantity is 0 or missing)
        has_valid_quantity = quantity > 0
        has_name_and_unit = item_name and unit

        if not (has_valid_quantity or has_name_and_unit):
            return None

        # Ignore rows with only "合計" and "単価" without meaningful data
        if ("合計" in item_name or "単価" in item_name) and not has_valid_quantity and not unit:
            return None

        # Create item key
        item_key = self._create_subtable_item_key(raw_fields, reference_number)

        if not item_key:
            return None

        # Create SubtableItem
        subtable_item = SubtableItem(
            item_key=item_key,
            raw_fields=raw_fields,
            quantity=quantity,
            unit=unit,
            source="PDF",
            page_number=page_num + 1,  # Convert 0-based to 1-based
            reference_number=reference_number,
            sheet_name=None  # PDF doesn't have sheet names
        )

        return {
            'item': subtable_item,
            'next_index': start_idx + rows_processed
        }

    def _create_subtable_item_key(self, raw_fields: Dict[str, str], reference_number: Optional[str]) -> str:
        """
        Create item key for subtable items.
        For subtables, the item_key should just be the item name (名称・規格).
        The reference number is stored separately in the reference_number field.
        """
        # Use 名称・規格 as primary identifier (without reference number prefix)
        base_key = ""

        if "名称・規格" in raw_fields and raw_fields["名称・規格"]:
            base_key = raw_fields["名称・規格"].strip()

        return base_key
