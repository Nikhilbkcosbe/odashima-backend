import pandas as pd
import re
import logging
from typing import List, Dict, Optional, Union, Tuple
from io import BytesIO
from ..schemas.tender import TenderItem, SubtableItem
from .excel_table_extractor_service import ExcelTableExtractorService

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
        # Optional custom column name for item extraction
        self.custom_item_name_column = None
        # Initialize the corrected table extractor service
        self.table_extractor_service = ExcelTableExtractorService()

    def extract_items_from_buffer(self, excel_buffer: BytesIO) -> List[TenderItem]:
        """
        Extract items from Excel buffer iteratively, sheet by sheet.
        """
        return self.extract_items_from_buffer_with_sheet(excel_buffer, None)

    def extract_items_from_buffer_with_sheet(self, excel_buffer: BytesIO, sheet_name: Optional[str] = None) -> List[TenderItem]:
        """
        Extract items from Excel buffer iteratively, sheet by sheet.
        Focuses on the specified sheet or all sheets if none specified.

        Args:
            excel_buffer: BytesIO buffer containing Excel data
            sheet_name: Specific sheet name to extract from (optional)
        """
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

    def extract_main_table_from_buffer(self, excel_buffer: BytesIO, sheet_name: str) -> List[TenderItem]:
        """
        Extract main table from Excel buffer using the corrected table extraction logic.
        This method uses the new ExcelTableExtractorCorrected logic for main table extraction.

        Args:
            excel_buffer: BytesIO buffer containing Excel data
            sheet_name: Name of the sheet to extract from

        Returns:
            List of TenderItem objects extracted from the main table
        """
        logger.info(
            f"Extracting main table from sheet '{sheet_name}' using corrected logic")

        try:
            # Use the corrected table extractor service
            main_table_items = self.table_extractor_service.extract_main_table_from_buffer(
                excel_buffer, sheet_name
            )

            logger.info(
                f"Extracted {len(main_table_items)} items from main table using corrected logic")
            return main_table_items

        except Exception as e:
            logger.error(
                f"Error extracting main table with corrected logic: {e}")
            # Fallback to original method if corrected logic fails
            logger.info("Falling back to original extraction method")
            return self.extract_items_from_buffer_with_sheet(excel_buffer, sheet_name)

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

            # If we find multiple matching patterns or item header, it's likely a header
            if matches >= 2 or has_item_header:
                logger.info(
                    f"Header found at row {i + 1} with {matches} patterns, item_header={has_item_header}")
                return i, row

            # Special handling for combined header patterns
            if ("費目/工種/種別/細別/規格" in row_str or
                    ("費目" in row_str and "工種" in row_str and "種別" in row_str)):
                logger.info(f"Combined header pattern found at row {i + 1}")
                return i, row

        return None, None

    def _should_process_subtable(self, header_row: pd.Series, sheet_name: str, subtable_num: int) -> bool:
        """
        Check if a subtable should be processed.
        """
        return True  # Process all subtables

    def _find_subtable_end(self, df: pd.DataFrame, start_row: int, max_row: int) -> int:
        """
        Enhanced subtable end detection: Find the end row of a subtable by looking for the next header or empty section.

        ENHANCED LOGIC:
        - Any single completely empty row indicates end of subtable (not requiring 3+ consecutive empty rows)
        - Header rows indicate start of new subtable
        - More sophisticated empty row detection considers data patterns
        """
        current_row = start_row
        consecutive_empty_rows = 0
        last_data_row = start_row - 1  # Track the last row with actual data

        logger.debug(f"Finding subtable end starting from row {start_row + 1}")

        while current_row < max_row:
            row = df.iloc[current_row]

            # Check if this looks like a new header row
            if self._is_potential_header_row(row):
                logger.debug(
                    f"Found potential header at row {current_row + 1}, ending subtable at row {last_data_row + 1}")
                return last_data_row

            # Enhanced empty row detection
            if self._is_subtable_boundary_row(row):
                consecutive_empty_rows += 1
                logger.debug(
                    f"Empty row detected at {current_row + 1}, consecutive count: {consecutive_empty_rows}")

                # ENHANCED: Even 1 empty row can indicate subtable end if followed by significant gap or new data pattern
                if consecutive_empty_rows >= 1:
                    # Look ahead to see if this is a real boundary
                    if self._is_real_subtable_boundary(df, current_row, max_row, consecutive_empty_rows):
                        logger.info(
                            f"Subtable boundary confirmed at row {current_row + 1}, ending subtable at row {last_data_row + 1}")
                        return last_data_row
            else:
                # Found non-empty row, reset counter and update last data row
                if consecutive_empty_rows > 0:
                    logger.debug(
                        f"Non-empty row at {current_row + 1}, resetting empty row counter")
                consecutive_empty_rows = 0
                last_data_row = current_row

            current_row += 1

        # Reached end of sheet
        logger.debug(
            f"Reached end of sheet, subtable ends at row {last_data_row + 1}")
        return last_data_row

    def _is_subtable_boundary_row(self, row: pd.Series) -> bool:
        """
        Conservative check for subtable boundary rows.
        ONLY treats completely empty rows as boundaries to preserve row spanning logic.
        """
        # Only treat completely empty rows as boundaries
        # This ensures row spanning logic is preserved for rows with any meaningful content
        return self._is_completely_empty_row(row)

    def _is_real_subtable_boundary(self, df: pd.DataFrame, empty_row_start: int, max_row: int, consecutive_empty_count: int) -> bool:
        """
        Conservative boundary detection - only treat truly empty rows as subtable boundaries.
        This preserves row spanning logic while still detecting genuine subtable separations.

        Args:
            df: DataFrame to analyze
            empty_row_start: Starting row of empty sequence
            max_row: Maximum row to check
            consecutive_empty_count: Number of consecutive empty rows found

        Returns:
            True if this represents a real subtable boundary
        """
        # Strategy 1: Multiple consecutive empty rows = definitive boundary
        if consecutive_empty_count >= 2:
            logger.debug(
                f"Confirmed boundary: {consecutive_empty_count} consecutive empty rows")
            return True

        # Strategy 2: Single empty row - CONSERVATIVE: Only treat as boundary if followed by clear indicators
        if consecutive_empty_count == 1:
            # Look ahead to see if this is followed by a clear new subtable start
            look_ahead_range = min(5, max_row - empty_row_start - 1)

            for i in range(1, look_ahead_range + 1):
                check_row_idx = empty_row_start + i
                if check_row_idx >= max_row:
                    break

                check_row = df.iloc[check_row_idx]

                # Clear boundary indicators: another empty row or header
                if self._is_subtable_boundary_row(check_row):
                    logger.debug(
                        f"Confirmed boundary: Found additional empty row at {check_row_idx + 1}")
                    return True

                if self._is_potential_header_row(check_row):
                    logger.debug(
                        f"Confirmed boundary: Found header row at {check_row_idx + 1}")
                    return True

                # If we find meaningful data that could be part of row spanning, don't treat as boundary
                if self._has_meaningful_data(check_row):
                    # Check if this looks like it could be part of row spanning (quantity-only or similar structure)
                    if self._could_be_row_spanning_continuation(check_row):
                        logger.debug(
                            f"Boundary rejected: Row {check_row_idx + 1} looks like row spanning continuation")
                        return False
                    # If it's clearly different data pattern, it might be a boundary
                    elif self._has_different_data_pattern(check_row):
                        logger.debug(
                            f"Confirmed boundary: Different data pattern at {check_row_idx + 1}")
                        return True

            # CONSERVATIVE: If no clear indicators, don't treat as boundary to preserve row spanning
            logger.debug(
                f"Boundary rejected: No clear boundary indicators found, preserving for row spanning")
            return False

        return False

    def _could_be_row_spanning_continuation(self, row: pd.Series) -> bool:
        """
        Check if row could be a continuation of row spanning logic.
        This helps preserve row spanning by not treating potential quantity-only rows as boundaries.
        """
        non_empty_cells = 0
        numeric_cells = 0

        # Check first 10 columns for data pattern
        for value in row[:10]:
            if not pd.isna(value):
                str_value = str(value).strip()
                if str_value and str_value not in ["", "None", "nan", "NaN", "0"]:
                    non_empty_cells += 1
                    if self._is_numeric(str_value):
                        numeric_cells += 1

        # Row spanning continuation often has:
        # 1. Few non-empty cells (1-3)
        # 2. At least one numeric value (quantity)
        # 3. Low overall data density

        could_be_spanning = (
            non_empty_cells <= 3 and  # Low data density
            # Has at least one number (potential quantity)
            numeric_cells >= 1 and
            non_empty_cells >= 1      # But not completely empty
        )

        if could_be_spanning:
            logger.debug(
                f"Row could be row spanning continuation: {non_empty_cells} cells, {numeric_cells} numeric")

        return could_be_spanning

    def _has_similar_data_structure(self, row: pd.Series) -> bool:
        """
        Check if row has similar data structure to typical construction tender data.
        This helps determine if data is continuation of same subtable.
        """
        non_empty_cells = 0
        numeric_cells = 0
        text_cells = 0

        for value in row[:10]:  # Check first 10 columns typically used in construction data
            if not pd.isna(value):
                str_value = str(value).strip()
                if str_value and str_value not in ["", "None", "nan", "NaN"]:
                    non_empty_cells += 1
                    if self._is_numeric(str_value):
                        numeric_cells += 1
                    else:
                        text_cells += 1

        # Construction tender rows typically have mixed text and numeric data
        has_mixed_data = text_cells > 0 and numeric_cells > 0
        has_reasonable_density = non_empty_cells >= 2

        return has_mixed_data and has_reasonable_density

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

    def _has_different_data_pattern(self, row: pd.Series) -> bool:
        """
        Check if row has a significantly different data pattern that might indicate a new subtable.
        """
        non_empty_values = [str(val).strip()
                            for val in row if pd.notna(val) and str(val).strip()]

        if len(non_empty_values) == 0:
            return False

        # Check for patterns that indicate new subtable start
        summary_patterns = ["合計", "小計", "計", "総計", "計算", "計画", "予算", "概算"]
        header_patterns = ["工種", "種別", "細別",
                           "数量", "単位", "費目", "名称", "規格", "項目"]

        row_text = " ".join(non_empty_values).replace("　", "").replace(" ", "")

        # If row contains summary patterns, it might be end of previous subtable
        has_summary = any(pattern in row_text for pattern in summary_patterns)

        # If row contains header patterns, it might be start of new subtable
        has_header_pattern = any(
            pattern in row_text for pattern in header_patterns)

        return has_summary or has_header_pattern

    def _has_meaningful_data(self, row: pd.Series) -> bool:
        """
        Check if row contains meaningful data (not just structural elements).
        """
        meaningful_content = 0

        for value in row:
            if not pd.isna(value):
                str_value = str(value).strip()
                if str_value and str_value not in ["", "None", "nan", "NaN", "0"]:
                    # Check if it looks like meaningful content
                    if len(str_value) > 2 or str_value.isdigit() or any(char.isalpha() for char in str_value):
                        meaningful_content += 1

        return meaningful_content >= 2  # At least 2 cells with meaningful content

    def _is_natural_subtable_break(self, row_position: int) -> bool:
        """
        Check if this position represents a natural subtable break based on position patterns.
        Some Excel files have natural breaking points that don't necessarily have multiple empty rows.
        """
        # This is a heuristic - in practice, you might want to adjust this based on your specific Excel format
        # For now, we'll be conservative and only confirm breaks that have clear indicators
        return False

    def _is_completely_empty_row(self, row: pd.Series) -> bool:
        """
        Enhanced check if all cells in the row are empty or contain only whitespace/NaN.
        Now includes more comprehensive empty value detection.
        """
        for value in row:
            if not pd.isna(value):
                str_value = str(value).strip()
                # Enhanced empty value detection
                if str_value and str_value not in ["", "None", "nan", "NaN", "0", "0.0", "-", "―", "ー"]:
                    # Check if it's just whitespace or special characters
                    clean_value = str_value.replace("　", "").replace(
                        " ", "").replace("\t", "").replace("\n", "")
                    if clean_value:  # If there's still content after removing whitespace
                        return False

        return True

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
        Enhanced row spanning logic to handle the following scenarios:
        1. Row with name only (no quantity data) → defer for spanning
        2. Row with quantity and we have pending name → combine
        3. Complete row or standalone → process normally

        Example scenarios:
        Row 1: name="鋼材", qty="" → pending
        Row 2: name="加工", qty="10", unit="kg" → result: "鋼材 加工", qty=10, unit="kg"
        OR
        Row 1: name="基礎工事", qty="" → pending
        Row 2: name="", qty="5", unit="m3" → result: "基礎工事", qty=5, unit="m3"
        """
        # Extract the 3 essential fields
        item_name = self._extract_item_name_simple(row, col_mapping)
        quantity = self._extract_quantity_simple(row, col_mapping)
        unit = self._extract_unit_simple(row, col_mapping)

        # Extract remarks/摘要 data
        remarks = self._extract_remarks_simple(row, col_mapping)

        logger.info(
            f"Excel Row {row_idx}: RAW DATA: " +
            " ".join([f"[{i}]='{row.iloc[i]}'" for i in range(
                min(8, len(row))) if pd.notna(row.iloc[i]) and str(row.iloc[i]).strip()])
        )
        logger.info(
            f"Excel Row {row_idx}: EXTRACTED: item='{item_name}', quantity={quantity}, unit='{unit}'")

        # Check if we have a previous incomplete item (pending name)
        has_previous_incomplete = existing_items and existing_items[-1].quantity == 0

        # Skip table headers and structural elements, BUT NOT if we have a previous incomplete item
        if self._is_table_header_or_structural_with_context(row, item_name, quantity, has_previous_incomplete):
            logger.info(
                f"Excel Row {row_idx}: SKIP - Table header/structural element")
            return "skipped"

        # CASE 1: Row with name only (no quantity data) → defer for spanning
        if item_name and quantity == 0:
            logger.info(
                f"Excel Row {row_idx}: CASE 1 - Name only (pending for spanning): '{item_name}'")
            raw_fields = {"工事区分・工種・種別・細別": item_name}
            if unit:
                raw_fields["単位"] = unit
            if remarks:
                raw_fields["摘要"] = remarks
            return TenderItem(
                item_key=item_name,
                raw_fields=raw_fields,
                quantity=0.0,
                unit=unit,
                source="Excel",
                page_number=None
            )

        # CASE 2: Row with quantity and we have pending name → combine
        elif quantity > 0 and has_previous_incomplete:
            if item_name:
                # CASE 2A: Current row also has name → concatenate: "previous_name current_name"
                logger.info(
                    f"Excel Row {row_idx}: CASE 2A - Concatenating names: pending + current")
                return self._concatenate_with_previous_item(existing_items, item_name, quantity, unit, remarks)
            else:
                # CASE 2B: Current row has no name → use pending name only
                logger.info(
                    f"Excel Row {row_idx}: CASE 2B - Using pending name only")
                return self._complete_previous_item_simple(existing_items, quantity, unit, remarks)

        # CASE 3: Complete row or standalone → process normally
        elif item_name and quantity > 0:
            logger.info(
                f"Excel Row {row_idx}: CASE 3 - Complete standalone item: '{item_name}' {quantity}")
            raw_fields = {"工事区分・工種・種別・細別": item_name}
            if unit:
                raw_fields["単位"] = unit
            if remarks:
                raw_fields["摘要"] = remarks
            return TenderItem(
                item_key=item_name,
                raw_fields=raw_fields,
                quantity=quantity,
                unit=unit,
                source="Excel",
                page_number=None
            )

        # CASE 4: Row has quantity but no item name and no previous incomplete item
        elif not item_name and quantity > 0 and not has_previous_incomplete:
            logger.info(
                f"Excel Row {row_idx}: CASE 4 - Quantity without context, skipping")
            return "skipped"

        # CASE 5: No useful data
        else:
            logger.info(
                f"Excel Row {row_idx}: CASE 5 - No useful data, skipping")
            return "skipped"

    def _concatenate_with_previous_item(self, existing_items: List[TenderItem],
                                        current_item_name: str, quantity: float, unit: str = "", remarks: str = "") -> str:
        """
        Concatenate current item name with previous incomplete item name.
        Result: "previous_name current_name"
        """
        if not existing_items:
            return "skipped"

        last_item = existing_items[-1]

        # Concatenate names: "previous_name current_name"
        combined_name = f"{last_item.item_key} {current_item_name}".strip()
        last_item.item_key = combined_name
        last_item.quantity = quantity
        last_item.raw_fields["工事区分・工種・種別・細別"] = combined_name

        # Update unit if provided
        if unit and unit.strip():
            last_item.unit = unit.strip()

        # Update remarks if provided
        if remarks and remarks.strip():
            last_item.raw_fields["摘要"] = remarks.strip()

        logger.info(
            f"Excel: Concatenated names -> '{combined_name}' with quantity {quantity}")
        return "merged"

    def _is_table_header_or_structural_with_context(self, row: pd.Series, item_name: str, quantity: float, has_previous_incomplete: bool) -> bool:
        """
        ENHANCED: Check if row is a table header or structural element to skip.
        CONTEXT-AWARE: Don't filter out potential completion rows when there's a previous incomplete item!
        CRITICAL FIX: Never filter rows that have valid quantities - they are real data rows!
        """
        # CRITICAL: Never filter rows with valid quantities (> 0) - they contain actual data
        if quantity > 0:
            logger.debug(f"Not filtering row with valid quantity: {quantity}")
            return False

        # Skip rows that look like table headers (but only if they have no quantity)
        if item_name:
            header_patterns = [
                "費目", "工種", "種別", "細別", "規格", "数量", "単位", "単価", "金額", "摘要",
                "名称", "品名", "項目", "明細"
            ]
            if any(pattern in item_name for pattern in header_patterns):
                return True

        # CRITICAL: If we have a previous incomplete item, don't filter potential completion rows
        if has_previous_incomplete and not item_name and quantity > 0:
            logger.debug(
                f"Not filtering potential completion row: quantity={quantity} for previous incomplete item")
            return False

        # Only filter out table numbers/structural elements (no item name, small integer quantities)
        if not item_name and 0 < quantity <= 10 and quantity == int(quantity):
            # Count meaningful cells in the row
            meaningful_cells = 0
            for cell in row:
                if pd.notna(cell) and str(cell).strip() not in ["", "0", "0.0"]:
                    meaningful_cells += 1

            # Only filter if VERY minimal cells (≤2) to avoid filtering completion rows
            # Completion rows typically have quantity + price, so at least 2 meaningful cells
            if meaningful_cells <= 2:
                return True

        return False

    def _extract_item_name_simple(self, row: pd.Series, col_mapping: Dict[str, int]) -> str:
        """
        SIMPLIFIED: Extract item name focusing ONLY on the main item column and hierarchical columns.
        Looks for: 費目・工種・種別・細別・規格 and variations
        """
        # Priority 1: Main item column
        main_item_columns = [
            "工事区分・工種・種別・細別",
            "費目・工種・種別・細別・規格",
            "費目",
            "工種",
            "種別",
            "細別",
            "規格"
        ]

        for col_name in main_item_columns:
            if col_name in col_mapping:
                col_idx = col_mapping[col_name]
                if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                    value = str(row.iloc[col_idx]).strip()
                    if (value and value not in ["", "None", "nan", "0"] and
                            not all(c in " 　\t\n\r" for c in value)):
                        return value

        # Priority 2: Check hierarchical columns where item names are often stored
        for col_name, col_idx in col_mapping.items():
            if "hierarchical_item" in col_name:
                if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                    value = str(row.iloc[col_idx]).strip()
                    if (value and value not in ["", "None", "nan", "0"] and
                        not all(c in " 　\t\n\r" for c in value) and
                            not self._is_likely_unit(value)):
                        return value

        return ""

    def _extract_unit_simple(self, row: pd.Series, col_mapping: Dict[str, int]) -> str:
        """
        SIMPLIFIED: Extract unit focusing on the main unit column and hierarchical columns.
        Looks for: 単位 and variations with spaces
        """
        # Priority 1: Main unit columns
        unit_columns = [
            "単位",        # Standard without space
            "単　位",      # With full-width space
            "単 位",       # With regular space
            "Unit",
            "units"
        ]

        for col_name in unit_columns:
            if col_name in col_mapping:
                col_idx = col_mapping[col_name]
                if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                    value = str(row.iloc[col_idx]).strip()
                    if (value and value not in ["", "None", "nan", "0"] and
                            not all(c in " 　\t\n\r" for c in value)):
                        return value

        # Priority 2: Check hierarchical columns where units might be stored
        for col_name, col_idx in col_mapping.items():
            if "hierarchical_item" in col_name:
                if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                    value = str(row.iloc[col_idx]).strip()
                    if (value and value not in ["", "None", "nan", "0"] and
                        not all(c in " 　\t\n\r" for c in value) and
                            self._is_likely_unit(value)):
                        return value

        return ""

    def _extract_remarks_simple(self, row: pd.Series, col_mapping: Dict[str, int]) -> str:
        """
        ENHANCED: Extract remarks/摘要 data from the row.
        Only extracts content when there is a pattern of Japanese counter character followed by a number.
        Otherwise treats the cell as empty or whitespace.
        """
        # Priority 1: Main remarks columns
        remarks_columns = [
            "摘要",        # Standard without space
            "摘　要",      # With full-width space
            "摘 要",       # With regular space
            "備考",        # Alternative name
            "notes",
            "remarks"
        ]

        for col_name in remarks_columns:
            if col_name in col_mapping:
                col_idx = col_mapping[col_name]
                if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                    value = str(row.iloc[col_idx]).strip()
                    if (value and value not in ["", "None", "nan", "0"] and
                            not all(c in " 　\t\n\r" for c in value)):
                        # ENHANCED: Only extract if it contains construction reference code
                        if self._contains_construction_reference_code(value):
                            logger.debug(
                                f"Extracted remarks with construction reference code: '{value}'")
                            return value
                        else:
                            # Treat as empty if no construction reference code found
                            logger.debug(
                                f"Filtered out remarks without construction reference code: '{value}'")
                            return ""

        return ""

    def _is_table_header_or_structural(self, row: pd.Series, item_name: str, quantity: float) -> bool:
        """
        SIMPLIFIED: Check if row is a table header or structural element to skip.
        CRITICAL: Don't filter out potential completion rows!
        """
        # Skip rows that look like table headers
        if item_name:
            header_patterns = [
                "費目", "工種", "種別", "細別", "規格", "数量", "単位", "単価", "金額", "摘要",
                "名称", "品名", "項目", "明細"
            ]
            if any(pattern in item_name for pattern in header_patterns):
                return True

        # CRITICAL: Don't filter out potential completion rows for hierarchical items
        # Only filter out when quantity is small AND there are very few meaningful cells
        if not item_name and 0 < quantity <= 10 and quantity == int(quantity):
            # Count meaningful cells in the row
            meaningful_cells = 0
            for cell in row:
                if pd.notna(cell) and str(cell).strip() not in ["", "0", "0.0"]:
                    meaningful_cells += 1

            # Only filter if VERY minimal cells (≤2) to avoid filtering completion rows
            # Completion rows typically have quantity + price, so at least 2 meaningful cells
            if meaningful_cells <= 2:
                return True

        return False

    def _combine_with_previous_item_simple(self, existing_items: List[TenderItem],
                                           current_item_name: str, quantity: float, unit: str = "", remarks: str = "") -> str:
        """
        SIMPLIFIED: Combine current item name with previous incomplete item.
        """
        if not existing_items:
            return "skipped"

        last_item = existing_items[-1]

        # Combine names and set quantity
        combined_name = f"{last_item.item_key} {current_item_name}".strip()
        last_item.item_key = combined_name
        last_item.quantity = quantity
        last_item.raw_fields["工事区分・工種・種別・細別"] = combined_name

        # Update unit if provided
        if unit and unit.strip():
            last_item.unit = unit.strip()

        # Update remarks if provided
        if remarks and remarks.strip():
            last_item.raw_fields["摘要"] = remarks.strip()

        logger.info(
            f"Excel: Combined items -> '{combined_name}' with quantity {quantity}")
        return "merged"

    def _complete_previous_item_simple(self, existing_items: List[TenderItem], quantity: float, unit: str = "", remarks: str = "") -> str:
        """
        SIMPLIFIED: Complete previous incomplete item with quantity and unit.
        """
        if not existing_items:
            return "skipped"

        last_item = existing_items[-1]
        last_item.quantity = quantity

        # Update unit if provided
        if unit and unit.strip():
            last_item.unit = unit.strip()

        # Update remarks if provided
        if remarks and remarks.strip():
            last_item.raw_fields["摘要"] = remarks.strip()

        if unit and unit.strip():
            logger.info(
                f"Excel: Completed item '{last_item.item_key}' with quantity {quantity} and unit '{unit}'")
        else:
            logger.info(
                f"Excel: Completed item '{last_item.item_key}' with quantity {quantity}")

        return "merged"

    # REMOVED: Old complex row data completion method - replaced with simple approach

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

        # Prioritize work classification column for accurate extraction
        # Check work classification first as it has the most accurate data
        if "工事区分・工種・種別・細別" in raw_fields and raw_fields["工事区分・工種・種別・細別"] and raw_fields["工事区分・工種・種別・細別"].strip():
            return True

        # Only check limited hierarchical fields to reduce redundant item creation
        hierarchical_item_fields = [
            "hierarchical_item_5",  # Most specific level
            "hierarchical_item_4",  # Second level
            "hierarchical_item_3",  # Third level only
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
        ENHANCED: Only completes when 単位 (unit) field is present to ensure valid row spanning.
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

        # CRITICAL: Must have 単位 (unit) field to be a valid completion row
        # This prevents table numbers from being treated as quantity completion
        if "単位" not in raw_fields or not raw_fields["単位"].strip():
            logger.debug(
                f"Excel: Rejecting completion row - no unit field present (quantity={quantity})")
            return "skipped"

        # Update the last item with quantity and any additional fields
        old_quantity = last_item.quantity
        last_item.quantity = quantity

        # Update unit if available in completion row
        if "単位" in raw_fields and raw_fields["単位"] and raw_fields["単位"].strip():
            last_item.unit = raw_fields["単位"].strip()
            logger.debug(
                f"Excel: Updated unit to '{last_item.unit}' for item '{last_item.item_key}'")

        # Merge additional fields (like unit) from the completion row
        for field_name, field_value in raw_fields.items():
            if field_name not in last_item.raw_fields or not last_item.raw_fields[field_name]:
                last_item.raw_fields[field_name] = field_value
                logger.debug(
                    f"Excel: Added field '{field_name}' = '{field_value}' to item '{last_item.item_key}'")

        logger.info(
            f"Excel row spanning completed: '{last_item.item_key}' quantity {old_quantity} -> {quantity}")
        logger.info(
            f"Excel row spanning: Added unit '{raw_fields['単位']}' to '{last_item.item_key}'")

        return "merged"

    def _is_quantity_only_row(self, raw_fields: Dict[str, str], quantity: float) -> bool:
        """
        Enhanced check for quantity-only rows (indicating row spanning).
        Only considers a row as quantity-only if it has 単位 (unit) field present.
        This prevents table numbers from being treated as quantities.
        """
        # Must have quantity but no other meaningful fields
        if quantity <= 0:
            return False

        # CRITICAL: Must have 単位 (unit) field to be considered a valid quantity row
        # This prevents table numbers (1, 2, 4, 5) from being treated as quantities
        if "単位" not in raw_fields or not raw_fields["単位"].strip():
            logger.debug(
                f"Excel: Rejecting quantity-only row - no unit field present (quantity={quantity})")
            return False

        # Check if we have any meaningful non-quantity fields (excluding unit)
        meaningful_fields = 0
        for field_name, field_value in raw_fields.items():
            if field_name != "単位" and field_value and field_value.strip():
                meaningful_fields += 1

        # If we have quantity + unit but no other fields, it's a valid spanned row
        is_quantity_only = meaningful_fields == 0

        if is_quantity_only:
            logger.debug(
                f"Excel: Detected valid quantity-only row with unit: quantity={quantity}, unit='{raw_fields['単位']}'")
        else:
            logger.debug(
                f"Excel: Rejecting quantity-only row - has other meaningful fields: {meaningful_fields}")

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

    def _create_item_key_from_fields(self, raw_fields: Dict[str, str], exclude_reference_codes: bool = False) -> str:
        """
        Create item key prioritizing actual item names from hierarchical columns over remarks.

        Args:
            raw_fields: Dictionary of field values from the row
            exclude_reference_codes: If True, excludes 摘要 (remarks) data to avoid reference codes
        """
        # Priority 1: If custom item name column is set, prioritize it
        if self.custom_item_name_column and self.custom_item_name_column in raw_fields:
            if raw_fields[self.custom_item_name_column] and raw_fields[self.custom_item_name_column].strip():
                return raw_fields[self.custom_item_name_column].strip()

        # Priority 2: Work classification column has the most accurate extraction
        # Prioritize work classification over hierarchical item fields
        work_classification_field = "工事区分・工種・種別・細別"
        if work_classification_field in raw_fields and raw_fields[work_classification_field] and raw_fields[work_classification_field].strip():
            item_name = raw_fields[work_classification_field].strip()
            # Filter out units, specifications, and obvious non-item values
            if (not self._is_likely_unit(item_name) and
                not self._is_likely_specification_notes(item_name) and
                    not item_name.isdigit()):
                return item_name

        # Priority 3: Only use hierarchical columns if work classification is empty
        # Use fewer hierarchical fields to reduce redundant items
        hierarchical_fields = [
            "hierarchical_item_5",  # Most specific level only
            "hierarchical_item_4",
            "hierarchical_item_3"   # Only keep top 3 levels to reduce noise
        ]

        for field in hierarchical_fields:
            if field in raw_fields and raw_fields[field] and raw_fields[field].strip():
                item_name = raw_fields[field].strip()
                # Filter out units, specifications, and obvious non-item values
                if (not self._is_likely_unit(item_name) and
                    not self._is_likely_specification_notes(item_name) and
                        not item_name.isdigit()):
                    return item_name

        # Priority 4: Fallback to specification or other fields (but not remarks which contain reference codes)
        fallback_fields = ["規格"]
        for field in fallback_fields:
            if field in raw_fields and raw_fields[field] and raw_fields[field].strip():
                return raw_fields[field].strip()

        # Priority 5: Last resort - use remarks (but only if not excluding reference codes)
        if not exclude_reference_codes and "摘要" in raw_fields and raw_fields["摘要"] and raw_fields["摘要"].strip():
            return raw_fields["摘要"].strip()

        # Final fallback: Work classification has already been checked at Priority 2
        # So just use specification if available
        key_fields = ["規格"]

        # Add 摘要 only if not excluding reference codes
        if not exclude_reference_codes:
            key_fields.append("摘要")

        # Use the first available field as the key
        for field in key_fields:
            if field in raw_fields and raw_fields[field] and raw_fields[field].strip():
                return raw_fields[field].strip()

        # If no primary fields available, create a simple key from available data (excluding quantity/price fields)
        for field_name, field_value in raw_fields.items():
            if field_value and field_value.strip() and field_name not in ["単位", "数量", "単価", "金額", "摘要" if exclude_reference_codes else ""]:
                return field_value.strip()

        # Return empty string if no meaningful key can be created
        return ""

    def _is_reference_code(self, text: str) -> bool:
        """
        Check if text looks like a reference code (e.g., 明54号, 23号内訳書) rather than an actual item name.
        """
        if not text or not text.strip():
            return False

        text = text.strip()

        # Common patterns for reference codes in Japanese construction documents
        reference_patterns = [
            r'^明\d+号$',        # 明54号, 明5号
            r'^\d+号内訳書$',     # 23号内訳書
            r'^\d+号$',          # Just number + 号
            r'^[A-Z]+\d+$',      # Letter + number codes
            r'^別紙\d+$',        # 別紙1, 別紙2
            r'^図\d+$',          # 図1, 図2
            r'^表\d+$',          # 表1, 表2
        ]

        import re
        for pattern in reference_patterns:
            if re.match(pattern, text):
                return True

        return False

    def _is_likely_unit(self, text: str) -> bool:
        """
        Check if text is likely a unit rather than an item name.
        This helps prevent units from being treated as item names in row spanning scenarios.
        """
        if not text or not text.strip():
            return False

        text = text.strip()

        # Common Japanese construction units
        common_units = [
            # Basic units
            "式", "m3", "m2", "m", "cm", "mm", "km",
            "個", "本", "箇所", "台", "基", "機", "枚", "組", "セット",
            "kg", "t", "g", "L", "ℓ",

            # Construction-specific units
            "孔",      # holes
            "穴",      # holes (alternative)
            "口",      # openings/mouths
            "回",      # times/rounds
            "日",      # days
            "時間",    # hours
            "分",      # minutes
            "年",      # years
            "月",      # months
            "週",      # weeks
            "人",      # people
            "名",      # people (counter)
            "社",      # companies
            "件",      # cases/items
            "箇",      # pieces (alternative to 個)
            "か所",    # locations (alternative to 箇所)
            "ヶ所",    # locations (another alternative)
            "ケ所",    # locations (katakana version)

            # Specific units found in problem cases
            "構造物",   # structures (found in 左官工法 case)
            "部材",     # members/materials (found in 補強部材 cases)
            "掛m2",     # hanging square meters (found in 単管足場 case)
            "掛㎡",     # hanging square meters (alternative)

            # Measurement units
            "㎡", "㎥", "㎝", "㎜", "㎞",
            "ha", "a",  # area units
            "坪", "畳",  # Japanese area units

            # Common single digit indicators
            "1", "0", "-"
        ]

        # Exact match check
        if text in common_units:
            return True

        # Pattern-based checks
        import re

        # Mathematical expressions that look like units
        if re.match(r'^[\d.]+[a-zA-Z]+$', text):  # Like "10m", "2.5kg"
            return True

        # Single character units (many Japanese units are single characters)
        if len(text) == 1 and re.match(r'[個本台基機枚組孔穴口回日人名社件箇]', text):
            return True

        # Units with numeric prefixes
        if re.match(r'^\d+[式個本台基機枚組孔穴口回日人名社件箇]$', text):
            return True

        return False

    def _is_likely_specification_notes(self, text: str) -> bool:
        """
        Check if text looks like specification/notes rather than an item name.
        This prevents notes like "A：1名､B：1名" from being concatenated to item names.
        """
        if not text or not text.strip():
            return False

        text = text.strip()

        import re

        # Pattern 1: Contains colons and commas with numbers (specification format)
        # Examples: "A：1名､B：1名", "A:1人,B:2人", "タイプA：5個､タイプB：3個"
        if re.search(r'[A-Za-z][:：]\d+[名人個本台]', text):
            return True

        # Pattern 2: Multiple comma-separated items with quantities
        # Examples: "1名､2名", "A型､B型", "10個､20個"
        if ',' in text or '､' in text:
            parts = re.split(r'[,､]', text)
            if len(parts) >= 2:
                # Check if parts contain quantity-like patterns
                quantity_parts = 0
                for part in parts:
                    part = part.strip()
                    if re.search(r'\d+[名人個本台基機枚組]', part):
                        quantity_parts += 1
                # If most parts have quantities, it's likely specifications
                if quantity_parts >= len(parts) / 2:
                    return True

        # Pattern 3: Contains parentheses with specifications
        # Examples: "(A：1名)", "（詳細：10個）"
        if re.search(r'[（(][^)）]*[:：]\d+[名人個本台][)）]', text):
            return True

        # Pattern 4: Contains specific specification keywords
        specification_keywords = [
            "タイプ", "型", "種類", "仕様", "規格", "詳細", "内訳", "明細",
            "A:", "B:", "C:", "A：", "B：", "C：",
            "1名", "2名", "3名", "4名", "5名",  # Common person counts
            "1人", "2人", "3人", "4人", "5人"   # Alternative person counts
        ]

        for keyword in specification_keywords:
            if keyword in text:
                return True

        # Pattern 5: Looks like enumeration (A, B, C with details)
        if re.search(r'^[A-Z][：:][^A-Z]*[､,][A-Z][：:]', text):
            return True

        return False

    def _is_detailed_specification(self, text: str) -> bool:
        """
        Check if text looks like detailed specifications that should be treated as completion data
        rather than a new item name. This is specifically for the row spanning issue.

        Examples:
        - "1部材当り平均質量G≦20kg" (weight specification)
        - "1構造物当り修復延べ体積:0.17m3,材料種類:ﾎﾟﾘﾏｰｾﾒﾝﾄﾓﾙﾀﾙ,鉄筋ｹﾚﾝ･鉄筋防錆処理:有り"
        - "安全ﾈｯﾄ:有り"
        - "塗装種別:有機ｼﾞﾝｸﾘｯﾁﾍﾟｲﾝﾄ(1層) ｽﾌﾟﾚｰ,塗装箇所:桁等,塗装回数:1回"
        """
        if not text or not text.strip():
            return False

        text = text.strip()

        import re

        # Pattern 1: Technical specifications with measurements and conditions
        # Examples: "1部材当り平均質量G≦20kg", "1構造物当り修復延べ体積:0.17m3"
        if re.search(r'\d+[部材構造物][当り]+.*[:：].*[\d.]+[a-zA-Z0-9]+', text):
            return True

        # Pattern 2: Contains technical symbols and measurements
        # Examples: "G≦20kg", "φ25mm", "H=1500", "R=2mm"
        if re.search(r'[G≦≧≤≥φΦH=L=W=R=][\d.]+[a-zA-Z]+', text):
            return True

        # Pattern 3: Complex specifications with multiple technical terms separated by commas
        # Examples: "材料種類:ﾎﾟﾘﾏｰｾﾒﾝﾄﾓﾙﾀﾙ,鉄筋ｹﾚﾝ･鉄筋防錆処理:有り"
        if ',' in text and ':' in text:
            parts = text.split(',')
            technical_parts = 0
            for part in parts:
                if ':' in part or '：' in part:
                    technical_parts += 1
            # If most parts contain technical details, it's a specification
            if technical_parts >= len(parts) / 2:
                return True

        # Pattern 4: Safety or condition specifications
        # Examples: "安全ﾈｯﾄ:有り", "防錆処理:有り", "ｹﾚﾝ:無し"
        if re.search(r'[安全防錆処理ｹﾚﾝﾈｯﾄ].*[:：][有無]り?', text):
            return True

        # Pattern 5: Painting/coating specifications (for 下塗ﾄﾗｽ部 etc.)
        # Examples: "塗装種別:有機ｼﾞﾝｸﾘｯﾁﾍﾟｲﾝﾄ(1層) ｽﾌﾟﾚｰ,塗装箇所:桁等,塗装回数:1回"
        if re.search(r'塗装[種別箇所回数][:：]', text):
            return True

        # Pattern 6: Technical specifications with parentheses and layers
        # Examples: "(1層)", "(2層)", "弱溶剤形変性ｴﾎﾟｷｼ樹脂塗料(2層)"
        if re.search(r'[(（]\d+層[)）]', text) or '樹脂塗料' in text or 'ｽﾌﾟﾚｰ' in text:
            return True

        # Pattern 7: Very long descriptive specifications (over 20 characters with technical terms)
        if len(text) > 20 and any(term in text for term in [
            "当り", "平均", "質量", "体積", "材料", "種類", "処理", "仕様", "規格",
            "ﾎﾟﾘﾏｰ", "ｾﾒﾝﾄ", "ﾓﾙﾀﾙ", "鉄筋", "ｹﾚﾝ", "防錆", "ﾈｯﾄ", "塗装", "塗料",
            "弱溶剤", "ふっ素", "淡彩", "有機", "変性", "ｴﾎﾟｷｼ"
        ]):
            return True

        # Pattern 8: Contains specific measurement formats
        # Examples: "L100×100×10×160(SS400)", "M22×55(S10T)"
        if re.search(r'[LM]\d+×\d+', text) or re.search(r'\([A-Z0-9]+\)$', text):
            return True

        return False

    def _is_likely_table_number_or_structural_element(self, row: pd.Series, col_mapping: Dict[str, int],
                                                      item_name: str, unit: str, quantity: float) -> bool:
        """
        Check if this row is likely a table number or structural element that should be ignored.

        Args:
            row: The pandas Series representing the row
            col_mapping: Column mapping dictionary
            item_name: Extracted item name (if any)
            unit: Extracted unit (if any)
            quantity: Extracted quantity

        Returns:
            True if this looks like a table number/structural element to ignore
        """
        # Criteria 1: Small integer quantity (1-10) with no item name or unit
        if (quantity <= 10 and quantity == int(quantity) and
                not item_name and not unit):
            logger.debug(
                f"Detected possible table number: quantity={quantity}, no item/unit")
            return True

        # Criteria 2: Check if the row only contains the quantity and nothing else meaningful
        meaningful_cells = 0
        for col_idx, cell_value in enumerate(row):
            if pd.notna(cell_value):
                str_value = str(cell_value).strip()
                if str_value and str_value not in ['', '0', '0.0']:
                    # Don't count the quantity cell itself
                    if not (str_value == str(quantity) or str_value == str(int(quantity))):
                        meaningful_cells += 1

        # If there are very few meaningful cells (≤1), it's likely structural
        if meaningful_cells <= 1:
            logger.debug(
                f"Detected structural element: quantity={quantity}, meaningful_cells={meaningful_cells}")
            return True

        # Criteria 3: Check for common table number patterns
        # Table numbers are often in the first few columns
        first_few_values = []
        for i in range(min(5, len(row))):
            if pd.notna(row.iloc[i]):
                first_few_values.append(str(row.iloc[i]).strip())

        # If the row starts with just a number and has minimal other data
        if (len(first_few_values) <= 2 and
            len(first_few_values) > 0 and
            first_few_values[0].isdigit() and
                int(first_few_values[0]) <= 20):
            logger.debug(
                f"Detected table number pattern: first_values={first_few_values}")
            return True

        return False

    def _is_complete_item_name(self, text: str) -> bool:
        """
        Check if text appears to be a complete item name that should NOT be concatenated.
        Complete items contain specifications, measurements, or are self-contained descriptions.
        """
        if not text or not text.strip():
            return False

        text = text.strip()

        import re

        # Pattern 1: Contains measurements or specifications with equals sign
        # Examples: "L=58.9km", "H=2.5m", "W=1000mm"
        if re.search(r'[LHWlhw]=[\d.]+[a-zA-Z]+', text):
            return True

        # Pattern 2: Contains specific measurements
        # Examples: "58.9km", "2.5m", "1000mm" (when part of item name)
        if re.search(r'\d+\.?\d*[kmcm]+', text):
            return True

        # Pattern 3: Contains material specifications
        # Examples: "φ25mm", "Φ300", "直径25mm"
        if re.search(r'[φΦ直径]\d+', text):
            return True

        # Pattern 4: Contains complete descriptive phrases
        # Examples: "運搬費" (transport cost), "材料費" (material cost), "工事費" (construction cost)
        if any(keyword in text for keyword in ["費", "工事", "材料", "運搬", "設置", "撤去", "組立", "解体"]):
            return True

        # Pattern 5: Contains specific construction item types that are typically complete
        # Examples: "ガードレール", "フェンス", "標識", "舗装"
        complete_item_types = [
            "ガードレール", "ガード", "フェンス", "標識", "舗装", "コンクリート",
            "アスファルト", "鉄筋", "型枠", "足場", "支保工", "土留", "排水",
            "配管", "電線", "照明", "信号", "看板"
        ]

        for item_type in complete_item_types:
            if item_type in text:
                return True

        # Pattern 6: Long descriptive names (usually complete)
        # Items longer than 8 characters are usually complete descriptions
        if len(text) > 8:
            return True

        return False

    def _extract_item_name(self, row: pd.Series, col_mapping: Dict[str, int]) -> str:
        """
        Extract item name from core columns: 費目/工種/種別/細別/規格
        ENHANCED: Also checks hierarchical columns where item names might be stored
        """
        # Core item name columns in priority order
        item_name_columns = [
            "工事区分・工種・種別・細別",  # Main work classification
            "費目",                      # Cost item
            "工種",                      # Work type
            "種別",                      # Category
            "細別",                      # Subcategory
            "規格"                       # Specification
        ]

        logger.debug(
            f"Extracting item name from row. Col mapping keys: {list(col_mapping.keys())}")

        # First try core columns
        for col_name in item_name_columns:
            if col_name in col_mapping:
                col_idx = col_mapping[col_name]
                if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                    value = str(row.iloc[col_idx]).strip()
                    # Enhanced check for meaningful content - exclude various types of spaces and empty patterns
                    if (value and value not in ["", "None", "nan", "0"] and
                            not all(c in " 　\t\n\r" for c in value)):  # Exclude full-width spaces and other whitespace
                        logger.debug(
                            f"Found item name in column '{col_name}': '{value}'")
                        return value

        # If not found in core columns, try hierarchical columns
        # (Many Japanese Excel files store item names in these columns)
        hierarchical_columns = ["hierarchical_item_1", "hierarchical_item_2",
                                "hierarchical_item_3", "hierarchical_item_4", "hierarchical_item_5"]

        for col_name in hierarchical_columns:
            if col_name in col_mapping:
                col_idx = col_mapping[col_name]
                if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                    value = str(row.iloc[col_idx]).strip()
                    # Enhanced check for meaningful content and ensure it's not a unit
                    if (value and value not in ["", "None", "nan", "0"] and
                        not all(c in " 　\t\n\r" for c in value) and
                            not self._is_likely_unit(value)):
                        logger.debug(
                            f"Found item name in hierarchical column '{col_name}': '{value}'")
                        return value

        logger.debug("No item name found, returning empty string")
        return ""

    def _extract_unit(self, row: pd.Series, col_mapping: Dict[str, int]) -> str:
        """
        Extract unit from unit column - handles Japanese Excel patterns where unit might be in different cells
        """
        unit_columns = ["単位", "Unit", "units"]

        logger.debug(
            f"Extracting unit from row. Col mapping keys: {list(col_mapping.keys())}")

        # First try the mapped unit column
        for col_name in unit_columns:
            if col_name in col_mapping:
                col_idx = col_mapping[col_name]
                if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                    value = str(row.iloc[col_idx]).strip()
                    # Include spaces
                    if value and value not in ["", "None", "nan", "0", " ", "　"]:
                        logger.debug(
                            f"Found unit in column '{col_name}': '{value}'")
                        return value

        # Japanese Excel pattern: unit might be in hierarchical columns (especially column 3)
        # Look in hierarchical columns for units like "式", "m3", etc.
        for col_name, col_idx in col_mapping.items():
            if "hierarchical_item" in col_name:
                if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                    value = str(row.iloc[col_idx]).strip()
                    if value and self._is_likely_unit(value):
                        logger.debug(
                            f"Found unit in hierarchical column '{col_name}': '{value}'")
                        return value

        logger.debug("No unit found, returning empty string")
        return ""

    def _extract_quantity_simple(self, row: pd.Series, col_mapping: Dict[str, int]) -> float:
        """
        Extract quantity from quantity column - handles Japanese formatting with spaces
        """
        # All possible quantity column variations including full-width spaces
        quantity_columns = [
            "数量",        # Standard without space
            "数　量",      # With full-width space
            "数 量",       # With regular space
            "Quantity",
            "Qty",
            "Count"
        ]

        logger.debug(
            f"Extracting quantity from row. Col mapping keys: {list(col_mapping.keys())}")

        for col_name in quantity_columns:
            if col_name in col_mapping:
                col_idx = col_mapping[col_name]
                if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                    cell_value = row.iloc[col_idx]
                    quantity = self._extract_quantity(cell_value)
                    logger.debug(
                        f"Found quantity column '{col_name}' at index {col_idx}, value: '{cell_value}' -> {quantity}")
                    return quantity

        # If no exact match found, try to find it by searching for 数 and 量 characters
        for col_name, col_idx in col_mapping.items():
            if "数" in col_name and "量" in col_name:  # Look for any column containing 数 and 量
                if col_idx < len(row) and not pd.isna(row.iloc[col_idx]):
                    cell_value = row.iloc[col_idx]
                    quantity = self._extract_quantity(cell_value)
                    logger.debug(
                        f"Found quantity column via character search '{col_name}' at index {col_idx}, value: '{cell_value}' -> {quantity}")
                    return quantity

        logger.debug("No quantity column found, returning 0.0")
        return 0.0

    def _combine_with_previous_item(self, existing_items: List[TenderItem],
                                    current_item_name: str, unit: str, quantity: float) -> str:
        """
        Combine current row data with previous incomplete item
        """
        if not existing_items:
            logger.warning("Excel: No previous item to combine with")
            return "skipped"

        last_item = existing_items[-1]

        # Combine item names: previous + current
        combined_name = f"{last_item.item_key} {current_item_name}".strip()

        # Update the last item
        last_item.item_key = combined_name
        last_item.quantity = quantity
        last_item.raw_fields["工事区分・工種・種別・細別"] = combined_name
        last_item.raw_fields["単位"] = unit

        logger.info(
            f"Excel: Combined items: '{last_item.item_key}' with {quantity} {unit}")

        return "merged"

    def _complete_with_unit_quantity(self, existing_items: List[TenderItem],
                                     unit: str, quantity: float) -> str:
        """
        Complete previous incomplete item with unit and quantity only
        """
        if not existing_items:
            logger.warning("Excel: No previous item to complete")
            return "skipped"

        last_item = existing_items[-1]

        # Update with unit and quantity
        last_item.quantity = quantity
        last_item.raw_fields["単位"] = unit

        logger.info(
            f"Excel: Completed item '{last_item.item_key}' with {quantity} {unit}")

        return "merged"

    def _complete_with_combined_name_unit_quantity(self, existing_items: List[TenderItem],
                                                   detailed_name: str, unit: str, quantity: float) -> str:
        """
        Complete previous incomplete item by combining names and adding unit and quantity
        """
        if not existing_items:
            logger.warning(
                "Excel: No previous item to complete with combined name")
            return "skipped"

        last_item = existing_items[-1]

        # Combine the base item name with the detailed specification
        original_name = last_item.item_key
        combined_name = f"{original_name} {detailed_name}".strip()

        # Update the item with combined name, unit and quantity
        last_item.item_key = combined_name
        last_item.quantity = quantity
        last_item.raw_fields["工事区分・工種・種別・細別"] = combined_name
        last_item.raw_fields["単位"] = unit

        logger.info(
            f"Excel: Completed item with combined name '{combined_name}' ({quantity} {unit})")

        return "merged"

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
                unit=raw_fields.get("単位"),
                source="Excel",
                page_number=None
            ))

        return items

    def extract_items(self, excel_path: str) -> List[TenderItem]:
        """
        Extract items from Excel file and convert to TenderItem objects.
        UPDATED: Now uses the new subtable-based processing with row spanning logic.
        """
        items = []
        excel_file = None

        try:
            # Get all sheet names - use context manager to ensure proper cleanup
            excel_file = pd.ExcelFile(excel_path)
            print(
                f"Processing Excel file with sheets: {excel_file.sheet_names}")

            # Process each sheet using NEW subtable-based approach with row spanning
            for sheet_idx, sheet_name in enumerate(excel_file.sheet_names):
                try:
                    print(
                        f"Processing sheet '{sheet_name}' with shape {pd.read_excel(excel_file, sheet_name=sheet_name, header=None).shape}")

                    # Use NEW processing method that includes row spanning logic
                    sheet_items = self._process_single_sheet(
                        excel_file, sheet_name, sheet_idx)
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

    def _contains_construction_reference_code(self, text: str) -> bool:
        """
        Check if text contains construction document reference codes.
        Used in construction documents to reference specifications (e.g., 単1号, 内2号, 外3号).

        Args:
            text: The text to check for reference codes

        Returns:
            True if text contains construction reference codes, False otherwise
        """
        if not text or not text.strip():
            return False

        import re

        # Construction document reference codes pattern
        # Matches patterns like: 単1号, 内2号, 外3号, 別4号, 参5号, 資6号, 機7号, 材8号, 計9号, 上10号
        spec_reference_pattern = r'[単内外別参資機材計上][\d０-９]+号'

        if re.search(spec_reference_pattern, text):
            logger.debug(
                f"Found construction reference code in remarks: '{text}'")
            return True

        logger.debug(
            f"No construction reference code found in remarks: '{text}' - treating as empty")
        return False

    def extract_subtables_from_buffer(self, excel_buffer: BytesIO) -> List[SubtableItem]:
        """
        Extract subtables from Excel buffer by scanning all non-main sheets for reference numbers.

        Args:
            excel_buffer: BytesIO buffer containing Excel data

        Returns:
            List of SubtableItem objects extracted from Excel sheets
        """
        all_subtable_items = []
        excel_file = None

        logger.info("Starting Excel subtable extraction from buffer")

        try:
            # Read from buffer
            excel_file = pd.ExcelFile(excel_buffer)
            sheet_names = excel_file.sheet_names

            logger.info(
                f"Excel file has {len(sheet_names)} sheets: {sheet_names}")

            # STEP 1: Discover reference patterns from main sheet
            logger.info(
                "Step 1: Discovering reference patterns from main Excel sheet...")
            discovered_patterns = self._discover_reference_patterns_from_main_sheet(
                excel_file, sheet_names)

            if not discovered_patterns:
                logger.warning(
                    "No reference patterns discovered from main sheet, using fallback patterns")
                # Fallback to known patterns
                discovered_patterns = ['単', '道', '内']

            logger.info(
                f"Discovered reference patterns: {discovered_patterns}")

            # STEP 2: Identify main sheet and subtable sheets
            main_sheet_name = self._identify_main_sheet(sheet_names)
            subtable_sheets = [
                sheet for sheet in sheet_names if sheet != main_sheet_name]

            logger.info(f"Main sheet: {main_sheet_name}")
            logger.info(
                f"Subtable sheets ({len(subtable_sheets)}): {subtable_sheets}")

            # STEP 3: Extract subtables from non-main sheets
            for sheet_name in subtable_sheets:
                logger.info(f"Processing subtable sheet: {sheet_name}")

                try:
                    sheet_subtables = self._extract_subtables_from_sheet(
                        excel_file, sheet_name, discovered_patterns
                    )

                    logger.info(
                        f"Extracted {len(sheet_subtables)} subtable items from sheet '{sheet_name}'")
                    all_subtable_items.extend(sheet_subtables)

                except Exception as e:
                    logger.error(
                        f"Error processing subtable sheet '{sheet_name}': {e}")
                    continue

            logger.info(
                f"Total subtable items extracted from Excel: {len(all_subtable_items)}")

        except Exception as e:
            logger.error(f"Error reading Excel buffer for subtables: {e}")
            raise
        finally:
            # Ensure proper cleanup
            if excel_file is not None:
                try:
                    excel_file.close()
                except Exception as e:
                    logger.error(f"Error closing Excel file: {e}")

        return all_subtable_items

    def _discover_reference_patterns_from_main_sheet(self, excel_file: pd.ExcelFile, sheet_names: List[str]) -> List[str]:
        """
        Dynamically discover all reference patterns from the main Excel sheet.

        Args:
            excel_file: Opened Excel file object
            sheet_names: List of all sheet names

        Returns:
            List of discovered reference prefixes (e.g., ['単', '内', '道'])
        """
        discovered_patterns = set()

        # Identify main sheet (typically the first sheet or one with aggregate data)
        main_sheet_name = self._identify_main_sheet(sheet_names)

        logger.info(
            f"Scanning main sheet '{main_sheet_name}' for reference patterns")

        try:
            # Read main sheet
            df = pd.read_excel(
                excel_file, sheet_name=main_sheet_name, header=None)

            # Look for reference patterns in all cells
            for row_idx in range(min(50, len(df))):  # Check first 50 rows
                for col_idx in range(min(20, len(df.columns))):  # Check first 20 columns
                    cell_value = df.iloc[row_idx, col_idx]
                    if pd.notna(cell_value):
                        cell_str = str(cell_value).strip()

                        # Skip very long text to avoid false positives
                        if len(cell_str) > 100:
                            continue

                        # Find all reference patterns in the cell
                        patterns = self._extract_all_reference_patterns_from_text(
                            cell_str)
                        discovered_patterns.update(patterns)

        except Exception as e:
            logger.error(f"Error scanning main sheet for patterns: {e}")

        # ENHANCED: Also scan subtable sheets for additional patterns
        logger.info(
            f"Scanning subtable sheets for additional reference patterns")
        for sheet_name in sheet_names:
            if sheet_name != main_sheet_name:
                try:
                    subtable_df = pd.read_excel(
                        excel_file, sheet_name=sheet_name, header=None)

                    # Scan first 100 rows for reference patterns (subtables are usually near the top)
                    for row_idx in range(min(100, len(subtable_df))):
                        for col_idx in range(min(20, len(subtable_df.columns))):
                            cell_value = subtable_df.iloc[row_idx, col_idx]
                            if pd.notna(cell_value):
                                cell_str = str(cell_value).strip()

                                # Skip very long text to avoid false positives
                                if len(cell_str) > 100:
                                    continue

                                # Find all reference patterns in the cell
                                patterns = self._extract_all_reference_patterns_from_text(
                                    cell_str)
                                discovered_patterns.update(patterns)
                except Exception as e:
                    logger.warning(f"Error scanning sheet '{sheet_name}': {e}")

        # Convert to sorted list for consistent ordering
        pattern_list = sorted(list(discovered_patterns))

        logger.info(
            f"Discovered {len(pattern_list)} unique reference patterns: {pattern_list}")

        return pattern_list

    def _identify_main_sheet(self, sheet_names: List[str]) -> str:
        """
        Identify the main table sheet (usually contains aggregate data).

        Args:
            sheet_names: List of all sheet names

        Returns:
            Name of the main sheet
        """
        # Look for keywords that indicate main/aggregate sheets
        main_indicators = ['標準', '本工事', '内訳書', '合計', '集計', 'メイン', 'main', '総計']

        for sheet_name in sheet_names:
            if any(indicator in sheet_name for indicator in main_indicators):
                return sheet_name

        # If no obvious main sheet, use the first sheet
        return sheet_names[0] if sheet_names else ""

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
                if self._is_valid_reference_prefix_excel(prefix) and number.isdigit():
                    patterns.append(prefix)

        return patterns

    def _is_valid_reference_prefix_excel(self, prefix: str) -> bool:
        """
        Check if a prefix is a valid reference prefix for Excel parsing.

        Args:
            prefix: Potential reference prefix

        Returns:
            True if it's a valid reference prefix
        """
        # Single character prefixes are most common
        if len(prefix) != 1:
            return False

        # Common Japanese reference prefixes
        common_prefixes = {
            '単', '内', '道', '諸', '雑', '材', '工', '機', '労', '共', '設',
            '備', '運', '管', '設', '電', '水', '土', '建', '構', '橋', '舗', '代'
        }

        # Allow common prefixes or single Japanese characters
        return prefix in common_prefixes or '\u3040' <= prefix <= '\u309f' or '\u30a0' <= prefix <= '\u30ff' or '\u4e00' <= prefix <= '\u9faf'

    def _extract_subtables_from_sheet(self, excel_file: pd.ExcelFile, sheet_name: str, discovered_patterns: List[str]) -> List[SubtableItem]:
        """
        Extract subtables from a single Excel sheet.
        ENHANCED: Continuously scans for multiple subtables per sheet.

        Args:
            excel_file: Opened Excel file object
            sheet_name: Name of the sheet to process
            discovered_patterns: List of reference prefixes discovered from main sheet

        Returns:
            List of SubtableItem objects from this sheet
        """
        subtable_items = []

        try:
            # Read the sheet
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            logger.info(f"Sheet '{sheet_name}' has shape {df.shape}")

            if df.empty:
                logger.warning(f"Sheet '{sheet_name}' is empty")
                return subtable_items

            # ENHANCED: Continuously scan for multiple subtables throughout the sheet
            current_row = 0
            subtable_count = 0

            while current_row < len(df):
                # Look for the next reference number starting from current_row
                reference_number = None
                reference_row = None

                # Scan ahead to find the next reference number (larger scan range)
                for scan_row in range(current_row, min(current_row + 50, len(df))):
                    row = df.iloc[scan_row]
                    found_reference = self._extract_reference_number_from_excel_row(
                        row, discovered_patterns)

                    if found_reference:
                        reference_number = found_reference
                        reference_row = scan_row
                        logger.info(
                            f"Found reference number '{reference_number}' at row {scan_row + 1}")
                        break

                if not reference_number:
                    # No more reference numbers found, we're done with this sheet
                    break

                # Look for subtable header AFTER the reference number
                header_row_idx = None
                col_mapping = {}

                # Search for header in the next few rows after the reference
                for header_scan_row in range(reference_row + 1, min(reference_row + 10, len(df))):
                    row = df.iloc[header_scan_row]

                    if self._is_excel_subtable_header_row(row):
                        header_row_idx = header_scan_row
                        col_mapping = self._get_excel_subtable_column_mapping(
                            row)
                        logger.info(
                            f"Found subtable header at row {header_scan_row + 1} for reference '{reference_number}'")
                        break

                if header_row_idx is not None and col_mapping:
                    # Process this subtable's data rows
                    data_start_row = header_row_idx + 1

                    # Find the end of this subtable by looking for the next reference number or end of sheet
                    subtable_end_row = len(df)  # Default to end of sheet

                    # Look ahead to find where this subtable ends
                    for end_scan_row in range(data_start_row, len(df)):
                        row = df.iloc[end_scan_row]

                        # Check if we've found a new reference number (indicating start of next subtable)
                        next_reference = self._extract_reference_number_from_excel_row(
                            row, discovered_patterns)
                        if next_reference and next_reference != reference_number:
                            subtable_end_row = end_scan_row
                            logger.info(
                                f"Subtable for '{reference_number}' ends at row {end_scan_row + 1} (next reference: '{next_reference}')")
                            break

                        # Check for natural subtable end indicators
                        if self._is_excel_subtable_end_row(row):
                            subtable_end_row = end_scan_row
                            logger.info(
                                f"Subtable for '{reference_number}' ends at row {end_scan_row + 1} (end indicator)")
                            break

                    # Process this subtable's data rows with the specific reference number
                    sheet_subtables = self._process_excel_subtable_data_rows_bounded(
                        df, data_start_row, subtable_end_row, col_mapping, reference_number, sheet_name, discovered_patterns
                    )

                    subtable_items.extend(sheet_subtables)
                    subtable_count += 1

                    logger.info(
                        f"Extracted {len(sheet_subtables)} items from subtable {subtable_count} ('{reference_number}') in sheet '{sheet_name}'")

                    # Continue searching from the end of this subtable
                    current_row = subtable_end_row
                else:
                    # No valid header found for this reference, skip ahead
                    logger.warning(
                        f"No valid subtable header found for reference '{reference_number}' in sheet '{sheet_name}'")
                    current_row = reference_row + 1

            logger.info(
                f"Extracted {subtable_count} subtables with {len(subtable_items)} total items from sheet '{sheet_name}'")

        except Exception as e:
            logger.error(
                f"Error extracting subtables from sheet '{sheet_name}': {e}")

        return subtable_items

    def _extract_reference_number_from_excel_row(self, row: pd.Series, discovered_patterns: List[str]) -> Optional[str]:
        """
        Extract reference numbers from an Excel row using dynamically discovered patterns.
        Enhanced with flexible character matching to handle full-width/half-width differences.

        Args:
            row: Excel row data
            discovered_patterns: List of reference prefixes found in main sheet

        Returns:
            Reference number string (e.g., "単 1号") or None
        """
        if not discovered_patterns:
            return None

        for cell_value in row:
            if pd.notna(cell_value):
                cell_str = str(cell_value).strip()
                normalized_cell = self._normalize_text_for_matching(cell_str)

                # Check each discovered pattern with flexible matching
                for pattern in discovered_patterns:
                    normalized_pattern = self._normalize_text_for_matching(
                        pattern)

                    # Look for pattern + number + 号 (flexible matching)
                    # First check if the pattern characters are present
                    if self._contains_characters(cell_str, pattern):
                        # Use regex on normalized text to extract number
                        escaped_pattern = re.escape(normalized_pattern)
                        reference_pattern = f'({escaped_pattern})(\\d+)号?'
                        match = re.search(reference_pattern, normalized_cell)

                        if match:
                            number = match.group(2)
                            return f"{pattern} {number}号"

                        # Alternative: look for just pattern + number (for short strings)
                        reference_pattern_alt = f'({escaped_pattern})(\\d+)'
                        match_alt = re.search(
                            reference_pattern_alt, normalized_cell)

                        if match_alt and len(normalized_cell) <= 10:
                            number = match_alt.group(2)
                            return f"{pattern} {number}号"

        return None

    def _normalize_text_for_matching(self, text: str) -> str:
        """
        Normalize text for flexible matching by handling full-width/half-width differences.

        Args:
            text: Text to normalize

        Returns:
            Normalized text for comparison
        """
        if not text:
            return ""

        # Convert to string and strip
        normalized = str(text).strip()

        # Remove all types of spaces and separators
        normalized = normalized.replace('\u3000', '')  # Full-width space
        # Full-width space (alternative)
        normalized = normalized.replace('　', '')
        normalized = normalized.replace(' ', '')       # Regular space
        normalized = normalized.replace('\t', '')      # Tab
        normalized = normalized.replace('/', '')       # Slash
        normalized = normalized.replace('・', '')      # Middle dot
        normalized = normalized.replace('-', '')       # Hyphen
        normalized = normalized.replace('(', '')       # Parentheses
        normalized = normalized.replace(')', '')
        normalized = normalized.replace('（', '')      # Full-width parentheses
        normalized = normalized.replace('）', '')

        # Convert full-width numbers and letters to half-width
        full_to_half = {
            '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
            '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
            'Ａ': 'A', 'Ｂ': 'B', 'Ｃ': 'C', 'Ｄ': 'D', 'Ｅ': 'E',
            'Ｆ': 'F', 'Ｇ': 'G', 'Ｈ': 'H', 'Ｉ': 'I', 'Ｊ': 'J',
            'Ｋ': 'K', 'Ｌ': 'L', 'Ｍ': 'M', 'Ｎ': 'N', 'Ｏ': 'O',
            'Ｐ': 'P', 'Ｑ': 'Q', 'Ｒ': 'R', 'Ｓ': 'S', 'Ｔ': 'T',
            'Ｕ': 'U', 'Ｖ': 'V', 'Ｗ': 'W', 'Ｘ': 'X', 'Ｙ': 'Y', 'Ｚ': 'Z'
        }

        for full, half in full_to_half.items():
            normalized = normalized.replace(full, half)

        return normalized

    def _contains_characters(self, text: str, target_chars: str) -> bool:
        """
        Check if text contains all characters from target_chars (flexible matching).

        Args:
            text: Text to search in
            target_chars: Characters to look for

        Returns:
            True if all target characters are found
        """
        normalized_text = self._normalize_text_for_matching(text)
        normalized_target = self._normalize_text_for_matching(target_chars)

        # Check if all characters from target are present in text
        for char in normalized_target:
            if char not in normalized_text:
                return False
        return True

    def _is_excel_subtable_header_row(self, row: pd.Series) -> bool:
        """
        Check if an Excel row is a subtable header based on column patterns.
        Enhanced with flexible character matching (ignores full-width/half-width differences).

        Args:
            row: Excel row data

        Returns:
            True if it's a subtable header row
        """
        if row is None or len(row) == 0:
            return False

        # Core required patterns for subtables (check for character presence)
        required_patterns = ["名称", "単位", "数量"]
        found_patterns = 0

        for cell_value in row:
            if pd.notna(cell_value):
                cell_str = str(cell_value).strip()

                # Check each required pattern using flexible character matching
                for pattern in required_patterns:
                    if self._contains_characters(cell_str, pattern):
                        found_patterns += 1
                        break

        # Consider it a header if we find at least 2 of the 3 required patterns
        is_header = found_patterns >= 2

        if is_header:
            logger.info(
                f"Excel subtable header detected with {found_patterns} patterns: {[str(cell) for cell in row if pd.notna(cell)]}")

        return is_header

    def _get_excel_subtable_column_mapping(self, header_row: pd.Series) -> Dict[str, int]:
        """
        Map Excel subtable column names to indices based on header row.
        Enhanced to handle Excel formatting with full-width spaces, separators, and complex patterns.

        Args:
            header_row: Header row data

        Returns:
            Dictionary mapping column names to indices
        """
        col_indices = {}

        # Subtable-specific column patterns based on Excel structure
        # Enhanced to handle complex Excel formatting
        subtable_column_patterns = {
            "名称・規格": ["名称", "規格", "名称・規格", "名\u3000称", "規\u3000格", "名称/規格", "名\u3000称/規\u3000格", "注\u3000釈"],
            "条件": ["条件", "条\u3000件"],
            "単位": ["単位", "単\u3000位"],
            "数量": ["数量", "数\u3000量"],
            "単価": ["単価", "単\u3000価"],
            "金額": ["金額", "金\u3000額"],
            "摘要": ["摘要", "備考", "摘\u3000要", "備\u3000考"]
        }

        for col_name, patterns in subtable_column_patterns.items():
            for i, cell_value in enumerate(header_row):
                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip()

                    # Use flexible character matching for each pattern
                    for pattern in patterns:
                        if self._contains_characters(cell_str, pattern):
                            col_indices[col_name] = i
                            logger.debug(
                                f"Mapped column '{col_name}' to index {i} (matched '{pattern}' in '{cell_str}')")
                            break

                    if col_name in col_indices:
                        break

        logger.info(f"Excel subtable column mapping: {col_indices}")
        return col_indices

    def _process_excel_subtable_data_rows_bounded(self, df: pd.DataFrame, data_start_row: int, data_end_row: int,
                                                  col_mapping: Dict[str, int], reference_number: Optional[str],
                                                  sheet_name: str, discovered_patterns: Optional[List[str]] = None) -> List[SubtableItem]:
        """
        Process Excel subtable data rows within specific boundaries with row spanning logic.

        Args:
            df: DataFrame containing the sheet data
            data_start_row: Starting row index for data
            data_end_row: Ending row index for data (exclusive)
            col_mapping: Column name to index mapping
            reference_number: Reference number for this subtable
            sheet_name: Name of the sheet being processed
            discovered_patterns: List of reference prefixes discovered from main sheet

        Returns:
            List of SubtableItem objects
        """
        subtable_items = []
        i = data_start_row

        logger.info(
            f"Processing subtable data rows {data_start_row + 1} to {data_end_row} for reference '{reference_number}'")

        while i < data_end_row:
            row = df.iloc[i]

            # Check if this is the end of the subtable (redundant check, but safe)
            if self._is_excel_subtable_end_row(row):
                logger.info(f"Found subtable end indicator at row {i + 1}")
                break

            # Try to process this row as a potential subtable item
            item_data = self._extract_excel_subtable_item_with_spanning(
                df, i, col_mapping, reference_number, sheet_name, discovered_patterns
            )

            if item_data:
                # Ensure we don't process beyond our boundary
                next_index = min(item_data['next_index'], data_end_row)
                subtable_items.append(item_data['item'])
                i = next_index
            else:
                i += 1

        logger.info(
            f"Extracted {len(subtable_items)} items from bounded subtable for reference '{reference_number}'")
        return subtable_items

    def _process_excel_subtable_data_rows(self, df: pd.DataFrame, data_start_row: int,
                                          col_mapping: Dict[str, int], reference_number: Optional[str],
                                          sheet_name: str, discovered_patterns: Optional[List[str]] = None) -> List[SubtableItem]:
        """
        Process Excel subtable data rows with row spanning logic.

        Args:
            df: DataFrame containing the sheet data
            data_start_row: Starting row index for data
            col_mapping: Column name to index mapping
            reference_number: Reference number for this subtable
            sheet_name: Name of the sheet being processed
            discovered_patterns: List of reference prefixes discovered from main sheet

        Returns:
            List of SubtableItem objects
        """
        subtable_items = []
        i = data_start_row

        while i < len(df):
            row = df.iloc[i]

            # Check if this is the end of the subtable
            if self._is_excel_subtable_end_row(row):
                break

            # Try to process this row as a potential subtable item
            item_data = self._extract_excel_subtable_item_with_spanning(
                df, i, col_mapping, reference_number, sheet_name, discovered_patterns
            )

            if item_data:
                subtable_items.append(item_data['item'])
                i = item_data['next_index']  # Skip processed rows
            else:
                i += 1

        return subtable_items

    def _is_excel_subtable_end_row(self, row: pd.Series) -> bool:
        """
        Check if an Excel row indicates the end of a subtable.
        ENHANCED: More conservative detection to avoid stopping too early.

        Args:
            row: Excel row data

        Returns:
            True if this row indicates subtable end
        """
        if row is None or len(row) == 0:
            return True

        # Check for typical end indicators - be more conservative
        for cell_value in row:
            if pd.notna(cell_value):
                cell_str = str(cell_value).strip()

                # Only treat as end if it's a clear overall total, not just a single subtable total
                # "計" alone is often just a subtotal, not the end of all subtables
                definitive_end_patterns = ["合計", "総計", "全計", "最終計"]

                # Check for definitive end patterns
                for pattern in definitive_end_patterns:
                    if pattern in cell_str and len(cell_str) <= 15:
                        return True

                # For "計" and "小計", be more restrictive - only end if it's clearly a final total
                # Look for patterns like "工事費計", "総工事費計", etc.
                if ("計" in cell_str or "小計" in cell_str) and len(cell_str) <= 10:
                    # If it's just "計" by itself, it's likely a subtable total - continue scanning
                    if cell_str == "計":
                        continue
                    # If it contains other text with "計", it might be a section total
                    if any(word in cell_str for word in ["工事", "費", "額", "総", "全"]):
                        return True

        return False

    def _extract_excel_subtable_item_with_spanning(self, df: pd.DataFrame, start_idx: int,
                                                   col_mapping: Dict[str, int], reference_number: Optional[str],
                                                   sheet_name: str, discovered_patterns: Optional[List[str]] = None) -> Optional[Dict]:
        """
        Extract Excel subtable item data with row spanning logic.

        Args:
            df: DataFrame containing the sheet data
            start_idx: Starting row index
            col_mapping: Column name to index mapping
            reference_number: Reference number for this subtable
            sheet_name: Name of the sheet being processed
            discovered_patterns: List of reference prefixes discovered from main sheet

        Returns:
            Dictionary with 'item' and 'next_index' or None if no valid item found
        """
        if start_idx >= len(df):
            return None

        # Look ahead up to 3 rows to gather complete item information
        item_name = None
        quantity = 0.0
        unit = None
        raw_fields = {}
        rows_processed = 1

        # ENHANCED: More aggressive row spanning to capture all related data
        for offset in range(min(5, len(df) - start_idx)):  # Extended search range
            row = df.iloc[start_idx + offset]

            # Stop spanning when we hit a completely empty row – treat it as a hard boundary
            if self._is_completely_empty_row(row):
                logger.debug(
                    f"Encountered completely empty row {start_idx + offset + 1} during spanning – terminating span here")
                break

            # ENHANCED FIX: Only span if row has meaningful data for subtables
            # Must have at least one of: item name, quantity, or unit
            has_meaningful_data = False
            for col_name, col_idx in col_mapping.items():
                if col_name in ["名称・規格", "数量", "単位"] and col_idx < len(row):
                    if pd.notna(row.iloc[col_idx]):
                        cell_value = str(row.iloc[col_idx]).strip()
                        if cell_value and cell_value not in [' ', '　', '']:
                            has_meaningful_data = True
                            break

            # Also check for meaningful data in any cell (for shifted columns)
            if not has_meaningful_data:
                for cell in row:
                    if pd.notna(cell):
                        cell_str = str(cell).strip()
                        if (cell_str and cell_str not in [' ', '　', ''] and
                            (len(cell_str) > 2 or self._is_likely_unit(cell_str) or
                             self._extract_quantity(cell_str) > 0)):
                            has_meaningful_data = True
                            break

            if not has_meaningful_data:
                logger.debug(
                    f"Skipping row {start_idx + offset + 1} - no meaningful subtable data")
                continue

            # First, try to extract data from expected columns
            for col_name, col_idx in col_mapping.items():
                cell_value = None
                found_value = False

                # Try the expected column first
                if col_idx < len(row) and pd.notna(row.iloc[col_idx]):
                    cell_value = str(row.iloc[col_idx]).strip()
                    if cell_value and cell_value not in [' ', '　', '']:
                        found_value = True

                # Process the found value
                if found_value and cell_value:
                    if col_name == "名称・規格" and not item_name:
                        item_name = cell_value
                        raw_fields[col_name] = cell_value
                    elif col_name == "数量" and col_name not in raw_fields:
                        quantity = self._extract_quantity(cell_value)
                        raw_fields[col_name] = cell_value
                    elif col_name == "単位" and not unit:
                        unit = cell_value
                        raw_fields[col_name] = cell_value
                    elif col_name not in raw_fields:
                        raw_fields[col_name] = cell_value

            # ENHANCED: Scan entire row for units and quantities if not found yet
            if not unit or quantity == 0.0:
                for col_idx, cell in enumerate(row):
                    if pd.notna(cell):
                        cell_str = str(cell).strip()
                        if cell_str and cell_str not in [' ', '　', '']:
                            # Look for units
                            if not unit and self._is_likely_unit(cell_str):
                                unit = cell_str
                                raw_fields["単位"] = cell_str

                            # Look for quantities
                            if quantity == 0.0:
                                try:
                                    test_qty = self._extract_quantity(cell_str)
                                    if test_qty > 0:  # Only positive quantities
                                        quantity = test_qty
                                        raw_fields["数量"] = cell_str
                                except:
                                    pass

            # ENHANCED: Stop if we have all essential data or find a new item name
            if item_name and unit and quantity > 0:
                break  # We have complete data

            # Stop if we find a new item name on a later row
            if offset > 0 and col_mapping.get("名称・規格", -1) < len(row):
                name_cell = row.iloc[col_mapping["名称・規格"]
                                     ] if col_mapping["名称・規格"] < len(row) else None
                if (pd.notna(name_cell) and str(name_cell).strip() and
                    item_name and str(name_cell).strip() != item_name and
                        not self._is_reference_number_only(str(name_cell).strip(), discovered_patterns or [])):
                    break

            rows_processed = offset + 1

        # Apply subtable filtering rules (same as PDF)
        # Must have item name
        if not item_name:
            return None

        # ENHANCED: Skip rows that contain only reference numbers (like "単 1号")
        if discovered_patterns and self._is_reference_number_only(item_name, discovered_patterns):
            logger.debug(f"Skipping reference number row: '{item_name}'")
            return None

        # ENHANCED: More aggressive filtering to capture all meaningful items
        # Accept items with any of:
        # 1. Item name with positive quantity (primary target)
        # 2. Item name with unit (even if quantity is 0)
        # 3. Item name that looks like a construction item (even without quantity/unit)
        # 4. Standalone quantity/unit data (for completion)

        has_positive_quantity = quantity > 0
        has_name_and_unit = item_name and unit
        has_meaningful_name = item_name and self._is_meaningful_construction_item(
            item_name)
        has_quantity_or_unit_data = quantity > 0 or unit

        # Primary acceptance criteria
        if has_positive_quantity or has_name_and_unit or (has_meaningful_name and has_quantity_or_unit_data):
            pass  # Accept these items
        # Accept longer item names that might be meaningful
        elif item_name and len(item_name) > 2:
            pass
        else:
            return None

        # Enhanced filtering - reject only clearly non-item rows
        if item_name:
            # Reject clear structural elements
            reject_patterns = ["合計", "小計", "計算",
                               "＊＊＊", "----", "====", "※", "注釈", "備考"]
            if any(pattern in item_name for pattern in reject_patterns) and not (quantity > 0 or unit):
                return None

            # Reject pure reference numbers or very short non-descriptive text
            if len(item_name) <= 2 and not (quantity > 0 or unit):
                return None

        # Create item key (just the item name, reference number stored separately)
        item_key = item_name

        if not item_key:
            return None

        # Create SubtableItem
        subtable_item = SubtableItem(
            item_key=item_key,
            raw_fields=raw_fields,
            quantity=quantity,
            unit=unit,
            source="Excel",
            page_number=None,  # Excel doesn't have page numbers
            reference_number=reference_number,
            sheet_name=sheet_name  # Add sheet name for Excel subtables
        )

        # Calculate next index more accurately to prevent double-processing
        # If we spanned multiple rows and gathered data from later rows,
        # we need to skip those to prevent processing them again
        if rows_processed > 1 and (unit or quantity > 0):
            # We likely gathered data from subsequent rows, so skip them
            next_index = start_idx + rows_processed
        else:
            # Only processed the current row
            next_index = start_idx + 1

        return {
            'item': subtable_item,
            'next_index': next_index
        }

    def _is_meaningful_construction_item(self, text: str) -> bool:
        """
        Check if text represents a meaningful construction item.

        Args:
            text: Text to check

        Returns:
            True if it looks like a construction item
        """
        if not text or len(text) < 3:
            return False

        # Construction-related keywords
        construction_keywords = [
            "作業", "工事", "設置", "撤去", "施工", "材料", "部材", "金具", "管",
            "柵", "欄", "板", "孔", "削", "溶接", "接続", "明細", "当り", "費",
            "コンクリート", "ｺﾝｸﾘｰﾄ", "アンカー", "ｱﾝｶｰ", "カプセル", "ｶﾌﾟｾﾙ",
            "チューブ", "ﾁｭｰﾌﾞ", "フレキシブル", "ﾌﾚｷｼﾌﾞﾙ", "排水", "防護", "高欄",
            "橋名", "支持", "樹脂", "現場"
        ]

        # Check if text contains construction keywords
        for keyword in construction_keywords:
            if keyword in text:
                return True

        # Check for typical construction item patterns
        if any(pattern in text for pattern in ["当り", "1m", "1個", "1孔", "1枚", "1組"]):
            return True

        return False

    def _is_reference_number_only(self, text: str, discovered_patterns: List[str]) -> bool:
        """
        Check if a text string is only a reference number (like "単 1号") that should be skipped.
        Enhanced with flexible character matching to handle full-width/half-width differences.

        Args:
            text: Text to check
            discovered_patterns: List of discovered reference patterns

        Returns:
            True if the text is only a reference number
        """
        if not text or not discovered_patterns:
            return False

        text_clean = text.strip()
        normalized_text = self._normalize_text_for_matching(text_clean)

        # Check each pattern with flexible matching
        for pattern in discovered_patterns:
            if self._contains_characters(text_clean, pattern):
                normalized_pattern = self._normalize_text_for_matching(pattern)
                escaped_pattern = re.escape(normalized_pattern)

                # Check if the entire text is just a reference number (flexible)
                reference_pattern = f'^{escaped_pattern}\\d+号?$'
                if re.search(reference_pattern, normalized_text):
                    return True

        return False

    # =====================================================
    # 2024 – Lightweight subtable extractor (simple version)
    # =====================================================
    def extract_subtables_from_buffer_simple(self, excel_buffer: BytesIO, main_sheet_name: str) -> List[SubtableItem]:
        """Simpler, opinionated Excel-subtable extraction.

        Parameters
        ----------
        excel_buffer : BytesIO
            Raw Excel bytes.
        main_sheet_name : str
            Name of the sheet that contains the *main* table.  All other
            sheets are scanned for subtables.

        Returns
        -------
        List[SubtableItem]
            Every row of every sub-table, already merged via row-spanning and
            linked to its reference number.
        """
        import unicodedata
        import re

        # -------------------------------------------------
        # Helper – normalise text (remove spaces / width)
        # -------------------------------------------------
        def _norm(txt: str) -> str:
            if not isinstance(txt, str):
                txt = str(txt) if txt is not None else ""
            txt = unicodedata.normalize("NFKC", txt)  # full->half width
            return txt.replace(" ", "").replace("　", "").strip()

        # -------------------------------------------------
        # Helper – detect header row (≥2 of the 4 targets)
        # -------------------------------------------------
        def _is_header_row(row: pd.Series) -> bool:
            targets = ["名称", "数量", "単位", "備考"]
            hit = 0
            for val in row:
                if pd.notna(val):
                    cell = _norm(val)
                    if any(t in cell for t in targets):
                        hit += 1
            return hit >= 2

        # -------------------------------------------------
        # Helper – map columns we care about
        # -------------------------------------------------
        def _map_header(row: pd.Series) -> dict:
            mapping = {}
            for idx, val in enumerate(row):
                if pd.isna(val):
                    continue
                cell = _norm(val)

                # Enhanced matching: focus on core characters only
                # Name/specification column
                if (("名称" in cell) or ("規格" in cell) or ("名" in cell and "称" in cell) or
                        ("規" in cell and "格" in cell)) and "name" not in mapping:
                    mapping["name"] = idx

                # Quantity column - flexible matching
                elif (("数量" in cell) or ("数" in cell and "量" in cell)) and "qty" not in mapping:
                    mapping["qty"] = idx

                # Unit column - enhanced flexible matching
                elif (("単位" in cell) or ("単" in cell and "位" in cell) or
                      # Also match individual characters if column is short
                      (len(cell) <= 3 and ("単" in cell or "位" in cell))) and "unit" not in mapping:
                    mapping["unit"] = idx

                # Remarks column - enhanced flexible matching for 摘要
                elif (("摘要" in cell) or ("備考" in cell) or
                      ("摘" in cell and "要" in cell) or ("備" in cell and "考" in cell) or
                      # Match individual characters if column is short
                      (len(cell) <= 3 and ("摘" in cell or "要" in cell or "備" in cell or "考" in cell))) and "remarks" not in mapping:
                    mapping["remarks"] = idx

            return mapping

        # -------------------------------------------------
        # Helper – reference-row detection
        # -------------------------------------------------
        ref_regex = re.compile(r"[一-龯A-Za-z]?\s*\d+\s*号")

        def _extract_reference(row: pd.Series) -> str | None:
            for val in row:
                if pd.notna(val):
                    m = ref_regex.search(_norm(val))
                    if m:
                        # Insert a space between prefix and number for uniformity
                        ref = m.group(0)
                        ref = re.sub(r"\s+", "", ref)  # remove existing spaces
                        ref = ref[:-1]  # drop the trailing '号'
                        # ensure prefix + number + 号 pattern
                        prefix, number = ref[0], ref[1:]
                        return f"{prefix} {number}号"
            return None

        # -------------------------------------------------
        # Open the workbook
        # -------------------------------------------------
        xl = pd.ExcelFile(excel_buffer)
        sheets = [s for s in xl.sheet_names if s != main_sheet_name]
        summary: list[SubtableItem] = []

        for sheet in sheets:
            df = pd.read_excel(xl, sheet_name=sheet, header=None)
            if df.empty:
                continue

            i = 0
            while i < len(df):
                # 1. look for a reference row
                ref = _extract_reference(df.iloc[i])
                if not ref:
                    i += 1
                    continue

                # 2. scan ahead (≤10 rows) for header
                header_idx = None
                for j in range(i + 1, min(i + 11, len(df))):
                    if _is_header_row(df.iloc[j]):
                        header_idx = j
                        break
                if header_idx is None:
                    i += 1
                    continue  # cannot parse without header

                col_map = _map_header(df.iloc[header_idx])
                if len(col_map) < 3:  # need at least name + qty + unit
                    i = header_idx + 1
                    continue

                # 3. iterate data rows until 3 blank rows
                blank = 0
                pending_name = ""
                pending_unit = ""
                last_name = ""
                last_unit = ""
                start_data = header_idx + 1

                # Helper to judge blank rows (ignores ditto marks and spaces)
                def _row_is_blank(r: pd.Series) -> bool:
                    for v in r:
                        if pd.notna(v):
                            c = _norm(v)
                            if c and c not in {"-", "―", "－", "—", "ｰ"}:
                                return False
                    return True

                for k in range(start_data, len(df)):
                    row = df.iloc[k]

                    if _row_is_blank(row):
                        blank += 1
                        if blank >= 3:
                            i = k  # resume search after this block
                            break
                        continue
                    blank = 0

                    # extract cells safely
                    def _cell(idx):
                        if idx not in col_map.values() or idx >= len(row):
                            return ""
                        val = row.iloc[idx]
                        if pd.isna(val):
                            return ""
                        txt = str(val).strip()
                        # Treat common ditto marks / placeholders as empty
                        if txt in {"-", "―", "－", "—", "ｰ"}:
                            return ""
                        return txt

                    # ENHANCED: Extract with adjacent column checking for unit
                    def _cell_with_adjacent_check(column_key, col_map, row):
                        """Extract cell value, checking adjacent columns if main column is empty (for unit column)"""
                        idx = col_map.get(column_key, -1)
                        if idx == -1:
                            return ""

                        # Try the mapped column first
                        main_val = _cell(idx)
                        if main_val:
                            return main_val

                        # For unit column, check adjacent columns if main is empty
                        if column_key == "unit":
                            # ENHANCED: More aggressive adjacent column checking, prioritizing rightward shifts
                            # Common in Excel files where data is shifted due to merged cells or complex headers
                            # Get name column to avoid confusion
                            name_col_idx = col_map.get("name", -1)
                            # Prioritize right shift, then left
                            for offset in [1, 2, 3, -1, -2, -3]:
                                adj_idx = idx + offset
                                # CRITICAL: Skip the name column to avoid extracting name as unit
                                if adj_idx == name_col_idx:
                                    continue
                                if 0 <= adj_idx < len(row):
                                    adj_val = row.iloc[adj_idx]
                                    if pd.notna(adj_val):
                                        adj_str = str(adj_val).strip()
                                        if (adj_str and adj_str not in {"-", "―", "－", "—", "ｰ"} and
                                            # Enhanced unit detection - more comprehensive patterns including Japanese construction units
                                            (len(adj_str) <= 10 and  # Allow longer units like "1部材当り"
                                             (_is_likely_unit_pattern(adj_str) or  # Use dedicated unit detection
                                              adj_str in ["ｔ", "t", "kg", "m", "m2", "m3", "個", "本", "枚", "組", "回", "式", "人", "日", "ｍ", "ｋｇ", "㎡", "㎥", "孔", "部材", "構造物"] or
                                              any(char in adj_str for char in ["ｔ", "m", "kg", "個", "本", "枚", "組", "回", "式", "人", "日", "ｍ", "ｋｇ", "孔", "部", "材", "構", "造", "物"]) or
                                              any(pattern in adj_str for pattern in ["当り", "あたり", "当", "m2", "m3", "㎡", "㎥", "mm", "cm", "km", "孔", "部材", "構造物", "断面", "系", "工区"])))):
                                            return adj_str

                        return ""

                    # Helper function for enhanced unit detection
                    def _is_likely_unit_pattern(text):
                        """Check if text looks like a unit with more comprehensive patterns"""
                        if not text or len(text) > 10:  # Increased length for Japanese construction units
                            return False

                        # Common Japanese/construction unit patterns
                        unit_patterns = [
                            # Weight/mass units
                            "ｔ", "t", "kg", "ｋｇ", "g", "ton", "トン",
                            # Length/area/volume units
                            "m", "ｍ", "mm", "cm", "km", "m2", "m3", "㎡", "㎥",
                            # Count units
                            "個", "本", "枚", "組", "回", "式", "人", "日", "時間",
                            # Volume units
                            "L", "ℓ", "リットル", "cc", "ml",
                            # Equipment/structure units
                            "台", "基", "箇所", "ヶ所", "か所", "set", "セット",
                            # Japanese construction-specific units
                            "孔", "部材", "構造物", "体", "棟", "間", "スパン", "径", "巻", "束",
                            "系", "系統", "区間", "工区", "断面", "面", "側", "面積", "延長",
                            "幅", "厚", "深度", "高さ", "長さ", "工", "材", "品", "点",
                            "一式", "丁", "対", "連", "段", "層", "区", "列", "班",
                            # Per-unit expressions
                            "m当り", "当り", "あたり", "当", "毎", "per"
                        ]

                        # Check exact matches
                        if text in unit_patterns:
                            return True

                        # Check if it contains unit-like patterns
                        for pattern in unit_patterns:
                            if pattern in text:
                                return True

                        # Check for numeric + unit patterns (like "1m当り", "1部材当り")
                        import re
                        if re.match(r'^\d+[a-zA-Zｍｔｇ個本枚組回式人日時間当り孔部材構造物]+$', text):
                            return True

                        return False

                    name = _cell(col_map.get("name", -1))
                    qty_raw = _cell(col_map.get("qty", -1))
                    unit = _cell_with_adjacent_check("unit", col_map, row)
                    remarks_val = _cell(col_map.get("remarks", -1))

                    # ENHANCED: Check adjacent columns for potential name continuations
                    # This handles cases where continuation text is in a different column
                    if not name and qty_raw and pending_name:
                        # Look for name continuation in adjacent columns (up to 3 columns from name column)
                        name_col_idx = col_map.get("name", -1)
                        if name_col_idx >= 0:
                            # Check 1-3 columns to the right
                            for adj_offset in range(1, 4):
                                adj_col_idx = name_col_idx + adj_offset
                                if adj_col_idx < len(row):
                                    adj_val = row.iloc[adj_col_idx]
                                    if pd.notna(adj_val):
                                        adj_str = str(adj_val).strip()
                                        if (adj_str and adj_str not in {"-", "―", "－", "—", "ｰ"} and
                                            # Check if it looks like a name continuation (contains parentheses or descriptive text)
                                            (("(" in adj_str) or ("（" in adj_str) or (")" in adj_str) or ("）" in adj_str) or
                                             (len(adj_str) > 3))):
                                            name = adj_str
                                            break

                    qty = 0.0
                    try:
                        qty = self._extract_quantity(
                            qty_raw) if qty_raw else 0.0
                    except Exception:
                        qty = 0.0

                    # Track latest non-empty values
                    if name:
                        last_name = name
                    if unit:
                        last_unit = unit

                    # --- Enhanced row-spanning merge logic ---
                    # Case 1: Row with name but NO quantity data → defer for spanning
                    if name and qty_raw == "":
                        pending_name = name
                        pending_unit = unit or pending_unit
                        continue

                    # Case 2: Row with quantity and we have pending name → combine
                    elif qty_raw and pending_name:
                        if name:
                            # Concatenate names: pending + current
                            final_name = f"{pending_name} {name}"
                        else:
                            # Use pending name only
                            final_name = pending_name
                        final_unit = pending_unit or unit or last_unit
                        pending_name = ""
                        pending_unit = ""

                    # Case 3: Complete row or standalone item
                    else:
                        final_name = name or last_name
                        final_unit = unit or last_unit

                    if final_name:
                        raw = {"名称・規格など": final_name, "名称・規格": final_name}
                        if qty_raw:
                            raw["数量"] = qty_raw
                        if final_unit:
                            raw["単位"] = final_unit
                        if remarks_val:
                            raw["備考"] = remarks_val

                        summary.append(
                            SubtableItem(
                                item_key=final_name,
                                raw_fields=raw,
                                quantity=qty,
                                unit=final_unit,
                                source="Excel",
                                page_number=None,
                                reference_number=ref,
                                sheet_name=sheet,
                            )
                        )

                else:
                    # reached end of sheet without hitting 3 consecutive blanks
                    i = len(df)
                    continue

        return summary

    # ------------------------------------------------------------------
    # PUBLIC ALIAS – use the lightweight extractor by default going forward
    # ------------------------------------------------------------------
    extract_subtables_from_buffer = extract_subtables_from_buffer_simple
