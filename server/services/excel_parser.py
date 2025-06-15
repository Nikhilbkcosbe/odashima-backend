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
        self.column_patterns = {
            "工事区分・工種・種別・細別": ["工事区分・工種・種別・細別", "工事区分", "工種", "種別", "細別", "費目"],
            "規格": ["規格", "規 格", "名称・規格", "名称", "項目", "品名"],
            "単位": ["単位", "単 位"],
            "数量": ["数量", "数 量"],
            "単価": ["単価", "単 価"],
            "金額": ["金額", "金 額"],
            "数量・金額増減": ["数量・金額増減", "増減", "変更"],
            "摘要": ["摘要", "備考", "摘 要"]
        }

    def extract_items_from_buffer(self, excel_buffer: BytesIO) -> List[TenderItem]:
        """
        Extract items from Excel buffer iteratively, sheet by sheet.
        """
        all_items = []
        excel_file = None

        logger.info("Starting Excel extraction from buffer")

        try:
            # Read from buffer
            excel_file = pd.ExcelFile(excel_buffer)
            total_sheets = len(excel_file.sheet_names)
            logger.info(
                f"Excel file has {total_sheets} sheets to process: {excel_file.sheet_names}")

            # Process each sheet iteratively
            for sheet_idx, sheet_name in enumerate(excel_file.sheet_names):
                logger.info(
                    f"Processing sheet {sheet_idx + 1}/{total_sheets}: '{sheet_name}'")

                try:
                    sheet_items = self._process_single_sheet(
                        excel_file, sheet_name, sheet_idx)

                    logger.info(
                        f"Extracted {len(sheet_items)} items from sheet '{sheet_name}'")

                    # Join items from this sheet to the total collection
                    all_items.extend(sheet_items)

                except Exception as e:
                    logger.error(f"Error processing sheet '{sheet_name}': {e}")
                    continue

            logger.info(f"Total items extracted from Excel: {len(all_items)}")

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
        """
        items = []

        try:
            # Read sheet data
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            logger.info(f"Sheet '{sheet_name}' has shape {df.shape}")

            if df.empty:
                logger.warning(f"Sheet '{sheet_name}' is empty")
                return items

            # Find header row iteratively
            header_row_idx, header_row = self._find_header_row_iteratively(
                df, sheet_name)

            if header_row_idx is None:
                logger.warning(f"No header row found in sheet '{sheet_name}'")
                return items

            logger.info(
                f"Found header at row {header_row_idx + 1} in sheet '{sheet_name}'")

            # Get column mapping
            col_mapping = self._get_column_mapping_from_header(
                header_row, sheet_name)

            if not col_mapping:
                logger.warning(
                    f"No recognizable columns found in sheet '{sheet_name}'")
                return items

            logger.info(
                f"Column mapping for sheet '{sheet_name}': {col_mapping}")

            # Process data rows iteratively with row spanning logic
            data_rows_start = header_row_idx + 1
            total_rows = len(df) - data_rows_start

            logger.info(
                f"Processing {total_rows} data rows in sheet '{sheet_name}'")

            for row_idx in range(data_rows_start, len(df)):
                try:
                    result = self._process_single_row_with_spanning(
                        df.iloc[row_idx], col_mapping, sheet_name, sheet_idx, row_idx, items
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
                        f"Error processing row {row_idx + 1} in sheet '{sheet_name}': {e}")
                    continue

        except Exception as e:
            logger.error(f"Error processing sheet '{sheet_name}': {e}")

        return items

    def _find_header_row_iteratively(self, df: pd.DataFrame, sheet_name: str) -> Tuple[Optional[int], Optional[pd.Series]]:
        """
        Find the header row by checking each row iteratively.
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

            # If we find multiple matching patterns, it's likely a header
            if matches >= 2:
                logger.info(
                    f"Header found at row {i + 1} with {matches} matching patterns")
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
        """
        col_mapping = {}

        logger.info(f"Analyzing header row for sheet '{sheet_name}'")

        for col_idx, cell_value in enumerate(header_row):
            if pd.isna(cell_value):
                continue

            cell_str = str(cell_value).strip()
            cell_clean = cell_str.replace(
                '\u3000', '').replace(' ', '').replace('　', '')

            logger.debug(f"Column {col_idx}: '{cell_str}' -> '{cell_clean}'")

            # Check against patterns
            for standard_name, patterns in self.column_patterns.items():
                for pattern in patterns:
                    if pattern in cell_clean:
                        col_mapping[standard_name] = col_idx
                        logger.info(
                            f"Mapped column {col_idx} ('{cell_str}') to '{standard_name}'")
                        break
                if standard_name in col_mapping:
                    break

            # Special handling for quantity column with Unicode spaces
            if standard_name not in col_mapping or standard_name != "数量":
                if ('数' in cell_str and '量' in cell_str) or '数\u3000量' in cell_str:
                    col_mapping["数量"] = col_idx
                    logger.info(
                        f"Mapped column {col_idx} ('{cell_str}') to '数量' (special case)")

        return col_mapping

    def _process_single_row_with_spanning(self, row: pd.Series, col_mapping: Dict[str, int],
                                          sheet_name: str, sheet_idx: int, row_idx: int,
                                          existing_items: List) -> Union[TenderItem, str, None]:
        """
        Process a single data row with row spanning logic.
        Returns: TenderItem, "merged", "skipped", or None
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

        # Check for quantity-only row (row spanning case)
        if self._is_quantity_only_row(raw_fields, quantity):
            return self._merge_quantity_with_previous_item(existing_items, quantity)

        # Skip if no meaningful fields extracted (empty row)
        if not raw_fields:
            return "skipped"

        # Create simple item key
        item_key = self._create_item_key_from_fields(raw_fields)

        # Skip if we couldn't create a meaningful key
        if not item_key:
            return "skipped"

        return TenderItem(
            item_key=item_key,
            raw_fields=raw_fields,
            quantity=quantity,
            source="Excel"
        )

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
        Check if this row only contains quantity data (indicating row spanning).
        """
        # Must have quantity but no other meaningful fields
        if quantity <= 0:
            return False

        # Check if we have any meaningful non-quantity fields
        meaningful_fields = 0
        for field_name, field_value in raw_fields.items():
            if field_value and field_value.strip():
                meaningful_fields += 1

        # If we have quantity but no other fields, it's likely a spanned row
        return meaningful_fields == 0

    def _merge_quantity_with_previous_item(self, existing_items: List[TenderItem], quantity: float) -> str:
        """
        Merge quantity with the most recent item (row spanning logic).
        """
        if not existing_items:
            logger.warning(
                "Quantity-only row found but no previous item to merge with")
            return "skipped"

        # Add quantity to the last item
        last_item = existing_items[-1]
        old_quantity = last_item.quantity
        new_quantity = old_quantity + quantity

        # Update the last item's quantity
        last_item.quantity = new_quantity

        logger.info(
            f"Merged quantity in Excel: {last_item.item_key} - {old_quantity} + {quantity} = {new_quantity}")

        return "merged"

    def _create_item_key_from_fields(self, raw_fields: Dict[str, str]) -> str:
        """
        Create a simple item key from available fields - each row treated independently.
        No hierarchical concatenation, just use the main identifying field.
        """
        # Priority order for creating item key (use first available field)
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
