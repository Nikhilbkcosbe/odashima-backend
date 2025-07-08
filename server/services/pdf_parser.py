import pdfplumber
import os
import re
import logging
from typing import List, Dict, Tuple, Optional, Union
from ..schemas.tender import TenderItem

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
