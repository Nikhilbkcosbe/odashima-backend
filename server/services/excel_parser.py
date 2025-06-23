import pandas as pd
import re
import logging
from typing import List, Dict, Optional, Union, Tuple
from io import BytesIO
from ..schemas.tender import TenderItem

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ExcelParser:
    def __init__(self):
        # Updated column patterns to match the PDF parser structure
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
        self.custom_item_name_column = None

    def set_custom_item_name_column(self, custom_column_name: Optional[str]):
        """
        Set a custom column name for item identification and update patterns accordingly.

        Args:
            custom_column_name: The custom column name to use for item identification
        """
        if custom_column_name:
            self.custom_item_name_column = custom_column_name
            # Create updated patterns with the custom column name as the primary item identifier
            self.column_patterns = self.default_column_patterns.copy()
            self.column_patterns["工事区分・工種・種別・細別"] = [
                custom_column_name] + self.default_column_patterns["工事区分・工種・種別・細別"]
            logger.info(
                f"Excel Parser: Set custom item name column to '{custom_column_name}'")
        else:
            self.custom_item_name_column = None
            self.column_patterns = self.default_column_patterns.copy()
            logger.info("Excel Parser: Using default column patterns")

    def extract_items_from_buffer(self, excel_buffer: BytesIO) -> List[TenderItem]:
        """
        Extract items from Excel buffer iteratively, sheet by sheet.
        """
        return self.extract_items_from_buffer_with_sheet(excel_buffer, None, None)

    def extract_items_from_buffer_with_sheet(self, excel_buffer: BytesIO, sheet_name: Optional[str] = None, item_name_column: Optional[str] = None) -> List[TenderItem]:
        """
        Extract items from Excel buffer with optional sheet filtering.

        Args:
            excel_buffer: BytesIO buffer containing Excel data
            sheet_name: Specific sheet name to extract from (None means all sheets)
            item_name_column: Custom column name to use for item identification (optional)
        """
        # Set custom item name column if provided
        self.set_custom_item_name_column(item_name_column)

        all_items = []
        excel_file = None

        logger.info("Starting Excel extraction from buffer")
        logger.info(f"Target sheet: {sheet_name or 'All sheets'}")

        try:
            # Read from buffer
            excel_file = pd.ExcelFile(excel_buffer)

            # Filter sheets based on sheet_name parameter
            if sheet_name:
                # Extract from specific sheet
                if sheet_name in excel_file.sheet_names:
                    target_sheets = [sheet_name]
                    logger.info(f"Found target sheet: '{sheet_name}'")
                else:
                    logger.error(
                        f"Sheet '{sheet_name}' not found. Available sheets: {excel_file.sheet_names}")
                    raise ValueError(
                        f"Sheet '{sheet_name}' not found in Excel file. Available sheets: {excel_file.sheet_names}")
            else:
                # Extract from all sheets
                target_sheets = excel_file.sheet_names

            total_sheets = len(target_sheets)
            logger.info(
                f"Excel file has {len(excel_file.sheet_names)} total sheets, processing {total_sheets} sheets: {target_sheets}")

            # Process target sheets iteratively
            for sheet_idx, current_sheet_name in enumerate(target_sheets):
                logger.info(
                    f"Processing sheet {sheet_idx + 1}/{total_sheets}: '{current_sheet_name}'")

                try:
                    sheet_items = self._process_single_sheet(
                        excel_file, current_sheet_name, sheet_idx)

                    logger.info(
                        f"Extracted {len(sheet_items)} items from sheet '{current_sheet_name}'")

                    # Join items from this sheet to the total collection
                    all_items.extend(sheet_items)

                except Exception as e:
                    logger.error(
                        f"Error processing sheet '{current_sheet_name}': {e}")
                    continue

            logger.info(
                f"Total items extracted from Excel ({len(target_sheets)} sheets): {len(all_items)}")

        except Exception as e:
            logger.error(f"Error reading Excel buffer: {e}")
            raise
        finally:
            # Ensure proper cleanup
            if excel_file is not None:
                try:
                    excel_file.close()
                except Exception as e:
                    logger.error(f"Error closing Excel file: {e}")

        return all_items

    def _process_single_sheet(self, excel_file: pd.ExcelFile, sheet_name: str, sheet_idx: int) -> List[TenderItem]:
        """
        Process a single Excel sheet and extract all valid items from it.
        Enhanced to detect and process multiple subtables within the sheet.
        """
        items = []

        try:
            # Read sheet data
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            logger.info(f"Sheet '{sheet_name}' has shape {df.shape}")

            if df.empty:
                logger.warning(f"Sheet '{sheet_name}' is empty")
                return items

            # Detect multiple subtables within the sheet
            subtables = self._detect_subtables_in_sheet(df, sheet_name)

            if not subtables:
                logger.warning(
                    f"No valid subtables found in sheet '{sheet_name}'")
                return items

            logger.info(
                f"Found {len(subtables)} subtables in sheet '{sheet_name}'")

            # Process each subtable iteratively
            for subtable_idx, subtable_info in enumerate(subtables):
                logger.info(
                    f"Processing subtable {subtable_idx + 1}/{len(subtables)} in sheet '{sheet_name}'")

                subtable_items = self._process_single_subtable(
                    df, subtable_info, sheet_name, sheet_idx, subtable_idx)

                logger.info(
                    f"Extracted {len(subtable_items)} items from subtable {subtable_idx + 1}")

                # Join subtable items to sheet items
                items.extend(subtable_items)

        except Exception as e:
            logger.error(f"Error processing sheet '{sheet_name}': {e}")

        return items

    def _detect_subtables_in_sheet(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        """
        Detect multiple subtables within a single Excel sheet.
        Returns list of subtable information with start/end rows and header positions.
        """
        subtables = []
        max_rows_to_scan = len(df)

        logger.info(
            f"Scanning {max_rows_to_scan} rows for subtables in sheet '{sheet_name}'")

        i = 0
        while i < max_rows_to_scan:
            # Look for potential header row
            header_row_idx, header_row = self._find_next_header_row(
                df, i, sheet_name)

            if header_row_idx is None:
                break  # No more headers found

            # Check if this subtable should be processed based on custom column requirements
            if not self._should_process_subtable(header_row, sheet_name, len(subtables)):
                i = header_row_idx + 1
                continue

            # Find the end of this subtable (next header or end of sheet)
            subtable_end = self._find_subtable_end(
                df, header_row_idx + 1, max_rows_to_scan)

            subtable_info = {
                'start_row': header_row_idx,
                'end_row': subtable_end,
                'header_row_idx': header_row_idx,
                'header_row': header_row
            }

            subtables.append(subtable_info)
            logger.info(
                f"Detected subtable at rows {header_row_idx + 1}-{subtable_end + 1} in sheet '{sheet_name}'")

            # Move to next potential subtable location
            i = subtable_end + 1

        return subtables

    def _find_next_header_row(self, df: pd.DataFrame, start_row: int, sheet_name: str) -> Tuple[Optional[int], Optional[pd.Series]]:
        """
        Find the next header row starting from start_row.
        Enhanced to use flexible pattern matching for core words.
        """
        max_rows_to_check = min(start_row + 20, len(df)
                                )  # Check next 20 rows max

        for i in range(start_row, max_rows_to_check):
            row = df.iloc[i]

            # Convert row to string for analysis
            row_str = " ".join([str(val) for val in row if pd.notna(val)])
            row_str_clean = row_str.replace(
                '\u3000', '').replace(' ', '').replace('　', '')

            # Check for construction document header patterns
            header_patterns = ["工種", "種別", "細別", "数量", "単位", "費目", "名称", "規格"]
            matches = sum(
                1 for pattern in header_patterns if pattern in row_str_clean)

            # Enhanced item name header detection using core words
            has_item_header = self._is_item_name_header_row(row)

            # Check for custom column if specified
            has_custom = False
            if self.custom_item_name_column:
                has_custom = any(self.custom_item_name_column in str(val)
                                 for val in row if pd.notna(val))

            # If we find multiple matching patterns or item header, it's likely a header
            if matches >= 2 or has_item_header or has_custom:
                logger.info(
                    f"Header found at row {i + 1} with {matches} patterns, item_header={has_item_header}, custom={has_custom}")
                return i, row

            # Special handling for combined header patterns
            if ("費目/工種/種別/細別/規格" in row_str or
                    ("費目" in row_str and "工種" in row_str and "種別" in row_str)):
                logger.info(f"Combined header pattern found at row {i + 1}")
                return i, row

        return None, None

    def _should_process_subtable(self, header_row: pd.Series, sheet_name: str, subtable_num: int) -> bool:
        """
        Check if a subtable should be processed based on custom column requirements.
        Enhanced to use flexible pattern matching for core words, similar to PDF parser.
        """
        if not self.custom_item_name_column:
            return True  # Process all subtables if no custom column specified

        # Check if any cell in the header row contains the custom column name (exact match)
        for cell in header_row:
            if pd.notna(cell) and self.custom_item_name_column in str(cell):
                logger.info(
                    f"Subtable {subtable_num + 1} in sheet '{sheet_name}' contains required column '{self.custom_item_name_column}' (exact match) - will process")
                return True

        # If no exact match, check for flexible matching with core words
        cleaned_custom_column = self._clean_text_for_matching(
            self.custom_item_name_column)
        core_words = ["費目", "工種", "種別", "細目"]
        custom_word_count = sum(
            1 for word in core_words if word in cleaned_custom_column)

        if custom_word_count >= 2:  # If custom column contains core words, use flexible matching
            for cell in header_row:
                if pd.notna(cell):
                    cleaned_cell = self._clean_text_for_matching(str(cell))
                    cell_word_count = sum(
                        1 for word in core_words if word in cleaned_cell)

                    # If both custom column and cell contain 2+ core words, consider it a match
                    if cell_word_count >= 2:
                        logger.info(
                            f"Subtable {subtable_num + 1} in sheet '{sheet_name}' contains matching core words (flexible match) - will process")
                        logger.info(
                            f"  Custom column: '{self.custom_item_name_column}' -> '{cleaned_custom_column}'")
                        logger.info(
                            f"  Found cell: '{cell}' -> '{cleaned_cell}'")
                        return True

        logger.info(
            f"Subtable {subtable_num + 1} in sheet '{sheet_name}' does not contain required column '{self.custom_item_name_column}' - skipping")
        return False

    def _find_subtable_end(self, df: pd.DataFrame, start_row: int, max_row: int) -> int:
        """
        Find the end row of a subtable by looking for the next header or empty section.
        """
        current_row = start_row
        consecutive_empty_rows = 0

        while current_row < max_row:
            row = df.iloc[current_row]

            # Check if this looks like a new header row
            if self._is_potential_header_row(row):
                # Found potential next header, end current subtable here
                return current_row - 1

            # Check if row is empty
            if self._is_completely_empty_row(row):
                consecutive_empty_rows += 1
                # If we find 3+ consecutive empty rows, consider it end of subtable
                if consecutive_empty_rows >= 3:
                    return current_row - consecutive_empty_rows
            else:
                consecutive_empty_rows = 0

            current_row += 1

        # Reached end of sheet
        return max_row - 1

    def _is_potential_header_row(self, row: pd.Series) -> bool:
        """
        Quick check if a row might be a header row (used for subtable boundary detection).
        """
        if row is None or len(row) == 0:
            return False

        # Convert row to string for analysis
        row_str = " ".join([str(val) for val in row if pd.notna(val)])
        row_str_clean = row_str.replace(
            '\u3000', '').replace(' ', '').replace('　', '')

        # Check for header indicators
        header_patterns = ["工種", "種別", "細別", "数量", "単位", "費目", "名称", "規格"]
        matches = sum(
            1 for pattern in header_patterns if pattern in row_str_clean)

        # If we have 2+ header patterns, it's likely a header
        return matches >= 2

    def _process_single_subtable(self, df: pd.DataFrame, subtable_info: Dict, sheet_name: str, sheet_idx: int, subtable_idx: int) -> List[TenderItem]:
        """
        Process a single subtable within an Excel sheet.
        """
        items = []

        try:
            header_row_idx = subtable_info['header_row_idx']
            header_row = subtable_info['header_row']
            end_row = subtable_info['end_row']

            logger.info(
                f"Processing subtable {subtable_idx + 1} (rows {header_row_idx + 1}-{end_row + 1}) in sheet '{sheet_name}'")

            # Get column mapping from header row
            col_mapping = self._get_column_mapping_from_header(
                header_row, f"{sheet_name}_subtable_{subtable_idx + 1}")

            if not col_mapping:
                logger.warning(
                    f"No recognizable columns found in subtable {subtable_idx + 1} of sheet '{sheet_name}'")
                return items

            logger.info(
                f"Column mapping for subtable {subtable_idx + 1}: {col_mapping}")

            # Process data rows within subtable boundaries
            data_rows_start = header_row_idx + 1
            data_rows_end = min(end_row + 1, len(df))
            total_rows = data_rows_end - data_rows_start

            logger.info(
                f"Processing {total_rows} data rows in subtable {subtable_idx + 1}")

            for row_idx in range(data_rows_start, data_rows_end):
                try:
                    result = self._process_single_row_with_spanning(
                        df.iloc[row_idx], col_mapping, f"{sheet_name}_subtable_{subtable_idx + 1}", sheet_idx, row_idx, items
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
                        f"Error processing row {row_idx + 1} in subtable {subtable_idx + 1}: {e}")
                    continue

        except Exception as e:
            logger.error(
                f"Error processing subtable {subtable_idx + 1} in sheet '{sheet_name}': {e}")

        return items

    def _find_header_row_iteratively(self, df: pd.DataFrame, sheet_name: str) -> Tuple[Optional[int], Optional[pd.Series]]:
        """
        Find the header row by checking each row iteratively.
        Enhanced to use flexible pattern matching for core words.
        """
        max_rows_to_check = min(15, len(df))  # Check first 15 rows

        logger.info(
            f"Searching for header row in first {max_rows_to_check} rows of sheet '{sheet_name}'")

        for i in range(max_rows_to_check):
            row = df.iloc[i]

            # Convert row to string for analysis
            row_str = " ".join([str(val) for val in row if pd.notna(val)])
            row_str_clean = row_str.replace(
                '\u3000', '').replace(' ', '').replace('　', '')

            # Check for construction document header patterns
            header_patterns = ["工種", "種別", "細別", "数量", "単位", "費目", "名称", "規格"]
            matches = sum(
                1 for pattern in header_patterns if pattern in row_str_clean)

            logger.debug(
                f"Row {i + 1} matches {matches} header patterns: '{row_str[:50]}...'")

            # Enhanced item name header detection using core words
            has_item_header = self._is_item_name_header_row(row)

            # Check for custom column if specified
            has_custom = False
            if self.custom_item_name_column:
                has_custom = any(
                    self.custom_item_name_column in str(val) for val in row if pd.notna(val))

            # If we find multiple matching patterns or item header, it's likely a header
            if matches >= 2 or has_item_header or has_custom:
                logger.info(
                    f"Header found at row {i + 1} with {matches} patterns, item_header={has_item_header}, custom={has_custom}")
                return i, row

            # Special handling for combined header patterns
            if ("費目/工種/種別/細別/規格" in row_str or
                    ("費目" in row_str and "工種" in row_str and "種別" in row_str)):
                logger.info(f"Combined header pattern found at row {i + 1}")
                return i, row

        # If no clear header found, try to find data-like rows and use previous row as header
        for i in range(1, max_rows_to_check):
            row = df.iloc[i]
            non_null_values = [str(val) for val in row if pd.notna(val)]

            if len(non_null_values) >= 3:
                # Check if it looks like a data row (mixed text and numbers)
                has_text = any(not self._is_numeric(val)
                               for val in non_null_values)
                has_numbers = any(self._is_numeric(val)
                                  for val in non_null_values)

                if has_text and has_numbers:
                    # Use previous row as header
                    header_idx = max(0, i - 1)
                    logger.info(
                        f"Inferred header at row {header_idx + 1} based on data pattern at row {i + 1}")
                    return header_idx, df.iloc[header_idx]

        return None, None

    def _is_item_name_header_row(self, row: pd.Series) -> bool:
        """
        Check if this row contains item name header based on core words.
        Removes special characters and looks for combinations of: 費目, 工種, 種別, 細目
        """
        if row is None or len(row) == 0:
            return False

        core_words = ["費目", "工種", "種別", "細目"]

        for cell in row:
            if pd.notna(cell):
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
        cleaned = re.sub(r'[・・/\-\s\(\)（）\[\]【】「」『』\|｜\.。、，\u3000]', '', text)
        return cleaned

    def _is_numeric(self, value: str) -> bool:
        """
        Check if a string value represents a number.
        """
        try:
            # Clean the value
            clean_val = str(value).replace(
                ',', '').replace(' ', '').replace('　', '')
            float(clean_val)
            return True
        except (ValueError, TypeError):
            return False

    def _get_column_mapping_from_header(self, header_row: pd.Series, sheet_name: str) -> Dict[str, int]:
        """
        Create column mapping from header row.
        Enhanced to map hierarchical item columns (1-5) in addition to standard columns.
        """
        col_mapping = {}

        logger.info(f"Analyzing header row for sheet '{sheet_name}'")

        for col_idx, cell_value in enumerate(header_row):
            if pd.isna(cell_value):
                # For empty columns 1-5, map as hierarchical item columns
                if 1 <= col_idx <= 5:
                    hierarchical_key = f"hierarchical_item_{col_idx}"
                    col_mapping[hierarchical_key] = col_idx
                    logger.info(
                        f"Mapped empty column {col_idx} to '{hierarchical_key}' (hierarchical item)")
                continue

            cell_str = str(cell_value).strip()
            cell_clean = cell_str.replace(
                '\u3000', '').replace(' ', '').replace('　', '')

            logger.debug(f"Column {col_idx}: '{cell_str}' -> '{cell_clean}'")

            # Check against patterns - FIXED: removed premature break
            column_mapped = False
            for standard_name, patterns in self.column_patterns.items():
                # Skip if this standard name is already mapped
                if standard_name in col_mapping:
                    continue

                for pattern in patterns:
                    if pattern in cell_clean:
                        col_mapping[standard_name] = col_idx
                        logger.info(
                            f"Mapped column {col_idx} ('{cell_str}') to '{standard_name}'")
                        column_mapped = True
                        break

                if column_mapped:
                    break

            # Special handling for quantity column with Unicode spaces
            if not column_mapped and "数量" not in col_mapping:
                if ('数' in cell_str and '量' in cell_str) or '数\u3000量' in cell_str:
                    col_mapping["数量"] = col_idx
                    logger.info(
                        f"Mapped column {col_idx} ('{cell_str}') to '数量' (special case)")

        logger.info(f"Final column mapping for '{sheet_name}': {col_mapping}")
        return col_mapping

    def _process_single_row_with_spanning(self, row: pd.Series, col_mapping: Dict[str, int],
                                          sheet_name: str, sheet_idx: int, row_idx: int,
                                          existing_items: List) -> Union[TenderItem, str, None]:
        """
        Enhanced row spanning logic to handle:
        - Row 1: Has item name but no quantity (or quantity = 0)
        - Row 2: Has quantity and unit but no item name -> Combine into single item
        - Row 2: Has item name AND quantity/unit -> Concatenate names and combine
        """
        # First check if row is completely empty
        if self._is_completely_empty_row(row):
            return "skipped"

        # Extract fields from row
        raw_fields = {}
        quantity = 0.0

        for col_name, col_idx in col_mapping.items():
            if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                cell_value = str(row.iloc[col_idx]).strip()
                if cell_value and cell_value not in ["", "None", "nan", "0"]:
                    if col_name == "数量":
                        quantity = self._extract_quantity(cell_value)
                    else:
                        raw_fields[col_name] = cell_value

        # Check what type of data this row contains
        has_item_fields = self._has_item_identifying_fields(raw_fields)
        has_quantity_data = quantity > 0 or "単位" in raw_fields

        logger.debug(
            f"Excel Row {row_idx}: has_item_fields={has_item_fields}, has_quantity_data={has_quantity_data}, quantity={quantity}")

        # Case 1: Row has item fields but no quantity data (incomplete item - needs spanning)
        if has_item_fields and not has_quantity_data:
            logger.debug(
                f"Excel Row {row_idx}: Creating incomplete item (name only) - expecting quantity in next row")
            item_key = self._create_item_key_from_fields(raw_fields)
            if not item_key:
                return "skipped"

            return TenderItem(
                item_key=item_key,
                raw_fields=raw_fields,
                quantity=0.0,  # Will be updated when quantity row is found
                source="Excel"
            )

        # Case 2: Row has quantity data but no item fields (completion row for spanning)
        elif has_quantity_data and not has_item_fields:
            logger.debug(
                f"Excel Row {row_idx}: Found completion row with quantity data only")
            # Check if this is a pure quantity-only row (row spanning scenario)
            if self._is_quantity_only_row(raw_fields, quantity):
                return self._merge_quantity_with_previous_item(existing_items, quantity)
            else:
                return self._complete_previous_item_with_quantity_data(existing_items, raw_fields, quantity)

        # Case 3: Row has both item fields and quantity data
        elif has_item_fields and has_quantity_data:
            # First check if previous item exists and needs completion (quantity = 0)
            if existing_items and existing_items[-1].quantity == 0:
                logger.debug(
                    f"Excel Row {row_idx}: Found item with name+quantity, combining with previous incomplete item")
                return self._combine_items_with_name_concatenation(existing_items, raw_fields, quantity)
            else:
                logger.debug(
                    f"Excel Row {row_idx}: Creating complete standalone item")
                item_key = self._create_item_key_from_fields(raw_fields)
                if not item_key:
                    return "skipped"

                return TenderItem(
                    item_key=item_key,
                    raw_fields=raw_fields,
                    quantity=quantity,
                    source="Excel"
                )

        # Case 4: Row has neither meaningful item fields nor quantity data
        else:
            logger.debug(f"Excel Row {row_idx}: Skipping - no meaningful data")
            return "skipped"

    def _combine_items_with_name_concatenation(self, existing_items: List[TenderItem],
                                               raw_fields: Dict[str, str], quantity: float) -> str:
        """
        Combine the previous incomplete item with current item by concatenating their names.
        """
        if not existing_items:
            logger.warning(
                "Excel: Found item for concatenation but no previous item to combine with")
            return "skipped"

        last_item = existing_items[-1]

        # Check if the last item needs completion (has quantity = 0)
        if last_item.quantity > 0:
            logger.warning(
                f"Excel: Previous item '{last_item.item_key}' already has quantity {last_item.quantity}")
            return "skipped"

        # Get current item name from this row
        current_item_key = self._create_item_key_from_fields(raw_fields)
        if not current_item_key:
            logger.warning(
                "Excel: Cannot create item key from current row for concatenation")
            return "skipped"

        # Concatenate item names: previous_name + current_name
        old_item_key = last_item.item_key
        concatenated_key = f"{old_item_key} + {current_item_key}"

        # Update the last item
        last_item.item_key = concatenated_key
        last_item.quantity = quantity

        # Merge all fields from both rows, prioritizing new row for duplicates
        for field_name, field_value in raw_fields.items():
            if field_name not in last_item.raw_fields or not last_item.raw_fields[field_name]:
                last_item.raw_fields[field_name] = field_value
                logger.debug(
                    f"Excel: Added field '{field_name}' = '{field_value}' to combined item")
            else:
                # For existing fields, combine them if different
                existing_value = last_item.raw_fields[field_name]
                if existing_value != field_value:
                    last_item.raw_fields[field_name] = f"{existing_value} + {field_value}"
                    logger.debug(
                        f"Excel: Combined field '{field_name}' = '{existing_value} + {field_value}'")

        logger.info(
            f"Excel row spanning with name concatenation: '{old_item_key}' + '{current_item_key}' = '{concatenated_key}' with quantity {quantity}")

        return "merged"

    def _has_item_identifying_fields(self, raw_fields: Dict[str, str]) -> bool:
        """
        Check if the row contains fields that identify an item (name, classification, specification).
        Enhanced to look across hierarchical columns for item names.
        """
        # If custom item name column is set, prioritize it
        if self.custom_item_name_column:
            # Check if the custom column is directly present in raw_fields
            if self.custom_item_name_column in raw_fields and raw_fields[self.custom_item_name_column] and raw_fields[self.custom_item_name_column].strip():
                return True

        # Look for item names in hierarchical structure (columns 1-5 contain item names)
        # This handles the Excel structure where item names appear in different columns based on hierarchy level
        hierarchical_item_fields = [
            "工事区分・工種・種別・細別",  # Main field
            "hierarchical_item_1",  # Level 1 items
            "hierarchical_item_2",  # Level 2 items
            "hierarchical_item_3",  # Level 3 items
            "hierarchical_item_4",  # Level 4 items
            "hierarchical_item_5",  # Level 5 items
        ]

        for field in hierarchical_item_fields:
            if field in raw_fields and raw_fields[field] and raw_fields[field].strip():
                return True

        # Also check specification and remarks as fallback
        other_identifying_fields = ["規格", "摘要"]
        for field in other_identifying_fields:
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
                "Excel: Found quantity completion row but no previous item to complete")
            return "skipped"

        last_item = existing_items[-1]

        # Check if the last item needs completion (has quantity = 0)
        if last_item.quantity > 0:
            logger.warning(
                f"Excel: Previous item '{last_item.item_key}' already has quantity {last_item.quantity}")
            return "skipped"

        # Update the last item with quantity and any additional fields
        old_quantity = last_item.quantity
        last_item.quantity = quantity

        # Merge additional fields (like unit) from the completion row
        for field_name, field_value in raw_fields.items():
            if field_name not in last_item.raw_fields or not last_item.raw_fields[field_name]:
                last_item.raw_fields[field_name] = field_value
                logger.debug(
                    f"Excel: Added field '{field_name}' = '{field_value}' to item '{last_item.item_key}'")

        logger.info(
            f"Excel row spanning completed: '{last_item.item_key}' quantity {old_quantity} -> {quantity}")
        if "単位" in raw_fields:
            logger.info(
                f"Excel row spanning: Added unit '{raw_fields['単位']}' to '{last_item.item_key}'")

        return "merged"

    def _is_completely_empty_row(self, row: pd.Series) -> bool:
        """
        Check if all cells in the row are empty or contain only whitespace/NaN.
        """
        for value in row:
            if not pd.isna(value):
                str_value = str(value).strip()
                if str_value and str_value not in ["", "None", "nan", "NaN", "0"]:
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
                f"Excel: Detected quantity-only row: quantity={quantity}, fields={meaningful_fields}")

        return is_quantity_only

    def _merge_quantity_with_previous_item(self, existing_items: List[TenderItem], quantity: float) -> str:
        """
        Enhanced quantity merging with better error handling and logging.
        """
        if not existing_items:
            logger.warning(
                "Excel: Quantity-only row found but no previous item to merge with")
            return "skipped"

        # Get the last item for merging
        last_item = existing_items[-1]
        old_quantity = last_item.quantity
        new_quantity = old_quantity + quantity

        # Update the last item's quantity
        last_item.quantity = new_quantity

        logger.info(
            f"Excel row spanning merge: '{last_item.item_key}' quantity {old_quantity} + {quantity} = {new_quantity}")

        return "merged"

    def _create_item_key_from_fields(self, raw_fields: Dict[str, str]) -> str:
        """
        Create item key prioritizing actual item names from hierarchical columns over remarks.
        Fixed: Extract actual item names, not reference codes from remarks.
        """
        # Priority 1: If custom item name column is set, prioritize it
        if self.custom_item_name_column and self.custom_item_name_column in raw_fields:
            if raw_fields[self.custom_item_name_column] and raw_fields[self.custom_item_name_column].strip():
                return raw_fields[self.custom_item_name_column].strip()

        # Priority 2: Look for actual item names in hierarchical columns (1-5)
        # These contain the real item names like "オープンカット、土砂", not reference codes
        hierarchical_fields = [
            "hierarchical_item_5",  # Most specific level first
            "hierarchical_item_4",
            "hierarchical_item_3",
            "hierarchical_item_2",
            "hierarchical_item_1",
            "工事区分・工種・種別・細別"  # Main field
        ]

        for field in hierarchical_fields:
            if field in raw_fields and raw_fields[field] and raw_fields[field].strip():
                item_name = raw_fields[field].strip()
                # Filter out units and obvious non-item values
                if item_name not in ["式", "m3", "m2", "m", "個", "本", "箇所", "1", "0"] and not item_name.isdigit():
                    return item_name

        # Priority 3: Fallback to specification or other fields (but not remarks which contain reference codes)
        fallback_fields = ["規格"]
        for field in fallback_fields:
            if field in raw_fields and raw_fields[field] and raw_fields[field].strip():
                return raw_fields[field].strip()

        # Priority 4: Last resort - use remarks (but this gives reference codes like "明1号")
        if "摘要" in raw_fields and raw_fields["摘要"] and raw_fields["摘要"].strip():
            return raw_fields["摘要"].strip()

        # Fallback to default fields if no custom column or custom column is empty
        key_fields = [
            "工事区分・工種・種別・細別",
            "規格",
            "摘要"
        ]

        # Use the first available field as the key
        for field in key_fields:
            if field in raw_fields and raw_fields[field] and raw_fields[field].strip():
                return raw_fields[field].strip()

        # If no primary fields available, create a simple key from available data
        for field_name, field_value in raw_fields.items():
            if field_value and field_value.strip() and field_name not in ["単位", "数量", "単価", "金額"]:
                return field_value.strip()

        # Return empty string if no meaningful key can be created
        return ""

    def _process_single_row(self, row: pd.Series, col_mapping: Dict[str, int],
                            sheet_name: str, sheet_idx: int, row_idx: int) -> Optional[TenderItem]:
        """
        Legacy method - kept for backward compatibility but now uses new spanning logic.
        """
        # Use new spanning logic with empty list (no previous items to merge with)
        result = self._process_single_row_with_spanning(
            row, col_mapping, sheet_name, sheet_idx, row_idx, [])

        if isinstance(result, TenderItem):
            return result
        else:
            return None

    def _find_header_row(self, df: pd.DataFrame) -> Optional[int]:
        """
        Find the row that contains column headers (legacy method for compatibility)
        """
        header_idx, _ = self._find_header_row_iteratively(df, "unknown")
        return header_idx

    def _find_column_mapping(self, header_row: pd.Series) -> Dict[str, int]:
        """
        Map column names to indices based on patterns (legacy method for compatibility)
        """
        return self._get_column_mapping_from_header(header_row, "unknown")

    def _extract_quantity(self, cell_value) -> float:
        """
        Extract numeric quantity from cell value
        """
        if pd.isna(cell_value):
            return 0.0

        # Convert to string and clean
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

    def _is_valid_data_row(self, row: pd.Series, col_mapping: Dict[str, int]) -> bool:
        """
        Check if row contains valid data
        """
        # Check if at least one mapped column has meaningful data
        meaningful_cells = 0
        for col_name, col_idx in col_mapping.items():
            if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                cell_value = str(row.iloc[col_idx]).strip()
                if cell_value and cell_value not in ["", "None", "nan", "0"]:
                    meaningful_cells += 1

        return meaningful_cells >= 1

    def _process_sheet(self, df: pd.DataFrame, sheet_name: str) -> List[TenderItem]:
        """
        Process a single Excel sheet (legacy method for compatibility)
        """
        # This method is kept for backward compatibility but now uses the new iterative approach
        items = []

        # Find header row
        header_row_idx = self._find_header_row(df)
        if header_row_idx is None:
            # Try alternative approach for sheets without clear headers
            # Look for rows with mixed text and numbers (data rows)
            for i in range(min(20, len(df))):
                row = df.iloc[i]
                non_null_values = [str(val) for val in row if pd.notna(val)]
                if len(non_null_values) >= 3:  # At least 3 columns with data
                    # Check if it looks like a data row (text + numbers)
                    has_text = any(not str(val).replace('.', '').replace(
                        ',', '').isdigit() for val in non_null_values)
                    has_numbers = any(str(val).replace('.', '').replace(
                        ',', '').isdigit() for val in non_null_values)
                    if has_text and has_numbers:
                        # Use previous row as header if possible
                        header_row_idx = max(0, i - 1)
                        break

            if header_row_idx is None:
                print(f"No header found in sheet: {sheet_name}")
                return items

        # Get column mapping
        header_row = df.iloc[header_row_idx]
        col_mapping = self._find_column_mapping(header_row)

        # If no mapping found, try generic approach
        if not col_mapping:
            print(
                f"No standard columns found in sheet: {sheet_name}, trying generic approach")
            # Try to find columns with common patterns
            for col_idx, cell_value in enumerate(header_row):
                if pd.isna(cell_value):
                    continue
                cell_str = str(cell_value).strip().replace(
                    '\u3000', '').replace(' ', '').replace('　', '')

                # Generic patterns for Japanese construction
                if any(pattern in cell_str for pattern in ["名称", "品名", "項目"]):
                    col_mapping["名称"] = col_idx
                elif any(pattern in cell_str for pattern in ["数量", "数"]):
                    col_mapping["数量"] = col_idx
                elif any(pattern in cell_str for pattern in ["単位"]):
                    col_mapping["単位"] = col_idx
                elif any(pattern in cell_str for pattern in ["単価", "価格"]):
                    col_mapping["単価"] = col_idx
                elif any(pattern in cell_str for pattern in ["金額", "合計"]):
                    col_mapping["金額"] = col_idx

        if not col_mapping:
            print(f"No recognizable columns found in sheet: {sheet_name}")
            return items

        print(f"Sheet '{sheet_name}': Found columns {col_mapping}")

        # Process data rows
        for row_idx in range(header_row_idx + 1, len(df)):
            row = df.iloc[row_idx]

            if not self._is_valid_data_row(row, col_mapping):
                continue

            # Extract fields
            raw_fields = {}
            for col_name, col_idx in col_mapping.items():
                if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                    raw_fields[col_name] = str(row.iloc[col_idx]).strip()

            # Skip if no meaningful fields
            if not raw_fields:
                continue

            # Extract quantity
            quantity = 0.0
            if "数量" in col_mapping:
                quantity = self._extract_quantity(row.iloc[col_mapping["数量"]])

            # Create item key from simple fields
            item_key = self._create_item_key_from_fields(raw_fields)

            items.append(TenderItem(
                item_key=item_key,
                raw_fields=raw_fields,
                quantity=quantity,
                source="Excel"
            ))

        return items

    def extract_items(self, excel_path: str) -> List[TenderItem]:
        """
        Extract items from Excel file and convert to TenderItem objects.
        """
        items = []
        excel_file = None

        try:
            # Get all sheet names - use context manager to ensure proper cleanup
            excel_file = pd.ExcelFile(excel_path)
            print(
                f"Processing Excel file with sheets: {excel_file.sheet_names}")

            # Process each sheet
            for sheet_name in excel_file.sheet_names:
                try:
                    df = pd.read_excel(
                        excel_file, sheet_name=sheet_name, header=None)
                    print(
                        f"Processing sheet '{sheet_name}' with shape {df.shape}")

                    sheet_items = self._process_sheet(df, sheet_name)
                    items.extend(sheet_items)

                    print(
                        f"Extracted {len(sheet_items)} items from sheet '{sheet_name}'")

                except Exception as e:
                    print(f"Error processing sheet '{sheet_name}': {e}")
                    continue

        except Exception as e:
            print(f"Error reading Excel file: {e}")
        finally:
            # Ensure proper cleanup of Excel file handle
            if excel_file is not None:
                try:
                    excel_file.close()
                except Exception as e:
                    print(f"Error closing Excel file: {e}")

        print(f"Total items extracted from Excel: {len(items)}")
        return items
