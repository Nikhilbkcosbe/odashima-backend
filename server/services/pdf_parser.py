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

    def extract_subtables_with_range(self, pdf_path: str, start_page: Optional[int] = None, end_page: Optional[int] = None) -> List[SubtableItem]:
        """
        Extract subtables from PDF with specified page range using different extraction logic.

        Subtable extraction requirements:
        1. Ignore rows with only "合計" and "単価" without quantities
        2. Column headers: 名称・規格, 単位, 単数, 摘要
        3. Find reference numbers like "単 3号" and associate table data with them

        Args:
            pdf_path: Path to the PDF file
            start_page: Starting page number (1-based, None means start from page 1)
            end_page: Ending page number (1-based, None means extract all pages)
        """
        all_subtable_items = []

        logger.info(f"Starting PDF subtable extraction from: {pdf_path}")
        logger.info(
            f"Subtable page range: {start_page or 'start'} to {end_page or 'end'}")

        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                logger.info(f"PDF has {total_pages} pages total")

                # STEP 1: Dynamically discover all reference patterns from main table
                logger.info(
                    "Step 1: Discovering reference patterns from main table...")
                discovered_patterns = self._discover_reference_patterns_from_main_table(
                    pdf)

                if not discovered_patterns:
                    logger.warning(
                        "No reference patterns discovered from main table, using fallback patterns")
                    # Fallback to known patterns
                    discovered_patterns = ['単', '道', '内']

                logger.info(
                    f"Discovered reference patterns: {discovered_patterns}")

                # STEP 2: Extract subtables using discovered patterns
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

                    page_subtables = self._extract_subtables_from_page(
                        page, page_num, discovered_patterns)

                    logger.info(
                        f"Extracted {len(page_subtables)} subtable items from page {page_num + 1}")

                    all_subtable_items.extend(page_subtables)

                logger.info(
                    f"Total subtable items extracted from PDF (pages {actual_start + 1}-{actual_end + 1}): {len(all_subtable_items)}")

        except Exception as e:
            logger.error(f"Error processing PDF for subtables: {e}")
            raise

        return all_subtable_items

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
        Extract reference numbers from a row using dynamically discovered patterns.
        This works with any PDF file by using patterns found in the main table.
        
        Args:
            row: Table row to search
            discovered_patterns: List of reference prefixes found in main table
            
        Returns:
            Reference number string (e.g., "単 45号") or None
        """
        if not discovered_patterns:
            return None
            
        for cell in row:
            if not cell:
                continue

            cell_str = str(cell).strip()

            # Create dynamic pattern based on discovered prefixes
            # Join all discovered patterns with | for regex OR
            escaped_patterns = [re.escape(pattern) for pattern in discovered_patterns]
            pattern_group = '|'.join(escaped_patterns)
            
            # Main pattern: (discovered_prefix) + optional space + number + 号
            reference_pattern = f'({pattern_group})\\s*(\\d+)\\s*号'
            match = re.search(reference_pattern, cell_str)

            if match:
                prefix = match.group(1)
                number = match.group(2)
                return f"{prefix} {number}号"

            # Alternative pattern without 号 but with discovered prefixes (for short strings only)
            reference_pattern_alt = f'({pattern_group})\\s*(\\d+)'
            match_alt = re.search(reference_pattern_alt, cell_str)

            # Only for short strings to avoid false positives
            if match_alt and len(cell_str.strip()) <= 10:
                prefix = match_alt.group(1)
                number = match_alt.group(2)
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
