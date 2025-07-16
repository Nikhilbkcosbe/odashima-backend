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

    def extract_tables_with_range(self, pdf_path: str, start_page: Optional[int] = None, end_page: Optional[int] = None) -> List[TenderItem]:
        """
        Extract tables from PDF iteratively with specified page range.
        This is the main entry point for parsing the main table.
        """
        all_items = []
        logger.info(f"Starting PDF extraction from: {pdf_path}")
        logger.info(
            f"Page range: {start_page or 'start'} to {end_page or 'end'}")

        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                logger.info(f"PDF has {total_pages} pages total")

                actual_start = (
                    start_page - 1) if start_page is not None else 0
                actual_end = (
                    end_page - 1) if end_page is not None else total_pages - 1

                actual_start = max(0, actual_start)
                actual_end = min(total_pages - 1, actual_end)

                if actual_start > actual_end:
                    logger.warning(
                        f"Invalid page range: start={actual_start+1}, end={actual_end+1}")
                    return all_items

                logger.info(
                    f"Processing pages {actual_start + 1} to {actual_end + 1}")

                for page_num in range(actual_start, actual_end + 1):
                    page = pdf.pages[page_num]
                    logger.info(
                        f"Processing page {page_num + 1}/{total_pages}")
                    page_items = self._extract_tables_from_page(page, page_num)
                    all_items.extend(page_items)

        except Exception as e:
            logger.error(
                f"Error processing PDF for main table: {e}", exc_info=True)
            raise
        return all_items

    def _extract_tables_from_page(self, page, page_num: int) -> List[TenderItem]:
        """Extract all tables from a single page."""
        page_items = []
        try:
            tables = page.extract_tables()
            logger.info(f"Found {len(tables)} tables on page {page_num + 1}")
            for table_num, table in enumerate(tables):
                page_items.extend(self._process_single_table(
                    table, page_num, table_num))
        except Exception as e:
            logger.error(
                f"Error processing page {page_num + 1}: {e}", exc_info=True)
        return page_items

    def _process_single_table(self, table: List[List], page_num: int, table_num: int) -> List[TenderItem]:
        """Process a single table and extract all valid items from it."""
        items = []
        if not table or len(table) < 2:
            return items

        header_row, header_idx = self._find_header_row(table)
        if header_row is None:
            return items

        col_indices = self._get_column_mapping(header_row)
        if not col_indices:
            return items

        data_rows = table[header_idx + 1:]
        for row_idx, row in enumerate(data_rows):
            try:
                result = self._process_single_row_with_spanning(
                    row, col_indices, page_num, table_num, header_idx + 1 + row_idx, items)
                if isinstance(result, TenderItem):
                    items.append(result)
            except Exception as e:
                logger.error(
                    f"Error processing row {row_idx + 1} in table {table_num + 1}: {e}", exc_info=True)
        return items

    def _process_single_row_with_spanning(self, row: List, col_indices: Dict[str, int],
                                          page_num: int, table_num: int, row_num: int,
                                          existing_items: List) -> Union[TenderItem, str, None]:
        """Handles row spanning for the main table."""
        if self._is_completely_empty_row(row):
            return "skipped"

        raw_fields, quantity, unit = self._extract_fields_from_row(
            row, col_indices)

        has_item_fields = self._has_item_identifying_fields(raw_fields)
        has_quantity_data = quantity > 0 or "単位" in raw_fields

        if has_item_fields and not has_quantity_data:
            item_key = self._create_item_key_from_fields(raw_fields)
            if not item_key:
                return "skipped"
            return TenderItem(item_key=item_key, raw_fields=raw_fields, quantity=0.0, unit=unit, source="PDF", page_number=page_num + 1)
        elif has_quantity_data and not has_item_fields:
            return self._complete_previous_item_with_quantity_data(existing_items, raw_fields, quantity)
        elif has_item_fields and has_quantity_data:
            item_key = self._create_item_key_from_fields(raw_fields)
            if not item_key:
                return "skipped"
            return TenderItem(item_key=item_key, raw_fields=raw_fields, quantity=quantity, unit=unit, source="PDF", page_number=page_num + 1)
        else:
            return "skipped"

    def _extract_fields_from_row(self, row: List, col_indices: Dict[str, int]) -> Tuple[Dict[str, str], float, Optional[str]]:
        """Extracts all relevant fields from a single row."""
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
                    raw_fields[col_name] = cell_value
        return raw_fields, quantity, unit

    def _has_item_identifying_fields(self, raw_fields: Dict[str, str]) -> bool:
        """Checks if the row contains fields that identify an item."""
        identifying_fields = ["工事区分・工種・種別・細別", "規格", "摘要"]
        return any(field in raw_fields and raw_fields[field] for field in identifying_fields)

    def _complete_previous_item_with_quantity_data(self, existing_items: List[TenderItem],
                                                   raw_fields: Dict[str, str], quantity: float) -> str:
        """Completes the previous incomplete item with quantity and unit data."""
        if not existing_items or existing_items[-1].quantity > 0:
            return "skipped"
        last_item = existing_items[-1]
        last_item.quantity = quantity
        if "単位" in raw_fields:
            last_item.unit = raw_fields["単位"]
        for k, v in raw_fields.items():
            if k not in last_item.raw_fields:
                last_item.raw_fields[k] = v
        return "merged"

    def _is_completely_empty_row(self, row: List) -> bool:
        """Checks if all cells in the row are empty or contain only whitespace."""
        return not any(cell and str(cell).strip() for cell in row)

    def _create_item_key_from_fields(self, raw_fields: Dict[str, str]) -> str:
        """Creates a concatenated item key."""
        key_fields = ["工事区分・工種・種別・細別", "摘要"]
        base_key = next(
            (raw_fields[f] for f in key_fields if f in raw_fields and raw_fields[f]), None)
        if not base_key:
            base_key = next((v for k, v in raw_fields.items() if v and k not in [
                            "単位", "数量", "単価", "金額", "規格"]), "")

        if "規格" in raw_fields and raw_fields["規格"]:
            return f"{base_key} + {raw_fields['規格']}"
        return base_key

    def _find_header_row(self, table: List[List]) -> Tuple[Optional[List], int]:
        """Finds the header row in the table."""
        for i, row in enumerate(table[:10]):
            if row and any(any(indicator in str(cell) for indicator in ["名称", "工種", "数量", "単位"]) for cell in row):
                return row, i
        return (table[0], 0) if table else (None, -1)

    def _get_column_mapping(self, header_row: List) -> Dict[str, int]:
        """Maps column names to indices based on header row."""
        col_indices = {}
        for col_name, patterns in self.column_patterns.items():
            for i, cell in enumerate(header_row):
                if cell and any(p in str(cell) for p in patterns):
                    col_indices[col_name] = i
                    break
        return col_indices

    def _extract_quantity(self, cell_value) -> float:
        """Extracts numeric quantity from a cell value."""
        if not cell_value:
            return 0.0
        value_str = str(cell_value).replace(",", "")
        number_match = re.search(r'[\d.]+', value_str)
        if number_match:
            try:
                return float(number_match.group())
            except ValueError:
                pass
        return 0.0

    def extract_subtables_with_range(self, pdf_path: str, start_page: Optional[int] = None, end_page: Optional[int] = None, reference_numbers: Optional[List[str]] = None) -> List[SubtableItem]:
        """
        Main entry point for extracting subtables from a PDF.
        Now scans all pages in the specified range and extracts every subtable matching the pattern (reference number row, header row, data rows).
        """
        all_subtable_items = []
        logger.info(f"Starting PDF subtable extraction from: {pdf_path}")
        logger.info(
            f"Page range: {start_page or 'start'} to {end_page or 'end'}")

        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                actual_start = (
                    start_page - 1) if start_page is not None else 0
                actual_end = (
                    end_page - 1) if end_page is not None else total_pages - 1
                actual_start = max(0, actual_start)
                actual_end = min(total_pages - 1, actual_end)

                if actual_start > actual_end:
                    return []

                logger.info(
                    f"Processing pages {actual_start + 1} to {actual_end + 1} for subtables (pattern-based)")
                for page_num in range(actual_start, actual_end + 1):
                    page = pdf.pages[page_num]
                    page_subtables = self._extract_all_subtables_from_page_pattern(
                        page, page_num)
                    all_subtable_items.extend(page_subtables)

        except Exception as e:
            logger.error(
                f"Error processing PDF for subtables: {e}", exc_info=True)
            raise

        return all_subtable_items

    def _extract_all_subtables_from_page_pattern(self, page, page_num: int) -> List[SubtableItem]:
        """
        Extracts all subtables from a single page by scanning for reference number rows, header rows, and data rows.
        """
        page_subtable_items = []
        try:
            tables = page.extract_tables()
            for table_num, table in enumerate(tables):
                i = 0
                while i < len(table):
                    row = table[i]
                    # Look for reference number row
                    ref = None
                    for cell in row:
                        if cell and isinstance(cell, str):
                            import re
                            m = re.search(r'([一-龯第])\s*(\d+)号', cell)
                            if m:
                                ref = f"{m.group(1)} {m.group(2)}号"
                                break
                            m2 = re.search(r'([一-龯第])(\d+)号', cell.replace(' ', ''))
                            if m2:
                                ref = f"{m2.group(1)}{m2.group(2)}号"
                                break
                    if not ref:
                        i += 1
                        continue
                    # Look for header row in the next few rows
                    header_row_idx, col_mapping = None, None
                    for j in range(i + 1, min(i + 6, len(table))):
                        if self._is_subtable_header_row(table[j]):
                            header_row_idx = j
                            col_mapping = self._get_subtable_column_mapping(table[j])
                            break
                    if header_row_idx is None or not col_mapping:
                        i += 1
                        continue
                    # Data rows start after header
                    data_start = header_row_idx + 1
                    # Data ends at next reference number or subtable end
                    data_end = len(table)
                    for k in range(data_start, len(table)):
                        # If we find another reference number row, stop
                        found_next_ref = False
                        for cell in table[k]:
                            if cell and isinstance(cell, str):
                                import re
                                if re.search(r'([一-龯第])\s*(\d+)号', cell) or re.search(r'([一-龯第])(\d+)号', cell.replace(' ', '')):
                                    data_end = k
                                    found_next_ref = True
                                    break
                        if found_next_ref:
                            break
                    data_rows = table[data_start:data_end]
                    # Extract items from data rows
                    page_subtable_items.extend(self._process_subtable_data_rows_with_spanning(
                        data_rows, col_mapping, page_num, ref, data_start))
                    i = data_end
        except Exception as e:
            logger.error(
                f"Error processing page {page_num + 1} for pattern-based subtables: {e}", exc_info=True)
        return page_subtable_items

    def _find_remarks_column_index(self, table: List[List]) -> Optional[int]:
        """Finds the index of the '摘要' (remarks) column."""
        for row in table[:5]:  # Check top 5 rows for header
            if row:
                for i, cell in enumerate(row):
                    if cell and ('摘要' in str(cell) or '備考' in str(cell)):
                        return i
        return None

    def _extract_complete_reference_numbers_from_text(self, text: str) -> List[str]:
        """Extracts complete reference numbers like '単3号' from text."""
        # This regex is simplified for clarity; it finds a Japanese char, digits, and '号'
        return re.findall(r'([一-龯第])(\d+)号', text)

    def _extract_subtables_from_page(self, page, page_num: int, reference_numbers: List[str]) -> List[SubtableItem]:
        """Extracts all subtables from a single page."""
        page_subtable_items = []
        try:
            tables = page.extract_tables()
            for table_num, table in enumerate(tables):
                page_subtable_items.extend(self._process_single_table_for_subtables(
                    table, page_num, table_num, reference_numbers))
        except Exception as e:
            logger.error(
                f"Error processing page {page_num + 1} for subtables: {e}", exc_info=True)
        return page_subtable_items

    def _process_single_table_for_subtables(self, table: List[List], page_num: int, table_num: int, reference_numbers: List[str]) -> List[SubtableItem]:
        """Finds and processes all subtables within a single larger table."""
        subtable_items = []
        i = 0
        while i < len(table):
            row = table[i]
            found_reference = self._find_reference_in_row(
                row, reference_numbers)
            if not found_reference:
                i += 1
                continue

            header_row_idx, col_mapping = self._find_subtable_header(
                table, i + 1)
            if header_row_idx is None:
                i += 1
                continue

            data_start = header_row_idx + 1
            data_end = self._find_subtable_data_end(
                table, data_start, reference_numbers)
            data_rows = table[data_start:data_end]

            subtable_items.extend(self._process_subtable_data_rows_with_spanning(
                data_rows, col_mapping, page_num, found_reference, data_start))
            i = data_end
        return subtable_items

    def _find_reference_in_row(self, row: List, reference_numbers: List[str]) -> Optional[str]:
        """Finds a reference number in a given row, normalizing spaces for robust matching."""
        for cell in row:
            if cell:
                cell_str = str(cell).replace(' ', '').replace('　', '').strip()
                for ref in reference_numbers:
                    ref_norm = ref.replace(' ', '').replace('　', '').strip()
                    if ref_norm in cell_str:
                        return ref
        return None

    def _find_subtable_header(self, table: List[List], start_idx: int) -> Tuple[Optional[int], Optional[Dict[str, int]]]:
        """Finds the header row and column mapping for a subtable."""
        for i in range(start_idx, min(start_idx + 5, len(table))):
            row = table[i]
            if self._is_subtable_header_row(row):
                return i, self._get_subtable_column_mapping(row)
        return None, None

    def _is_subtable_header_row(self, row: List) -> bool:
        """Checks if a row is a subtable header."""
        if not row:
            return False
        required_headers = ["名称", "単位", "数量"]
        row_text = ''.join(str(c) for c in row)
        return sum(1 for header in required_headers if header in row_text) >= 2

    def _get_subtable_column_mapping(self, header_row: List) -> Dict[str, int]:
        """Gets column mapping for subtable headers."""
        col_mapping = {}
        patterns = {"名称・規格": ["名称", "規格"], "単位": [
            "単位"], "数量": ["数量"], "摘要": ["摘要", "備考"]}
        for col_name, pats in patterns.items():
            for i, cell in enumerate(header_row):
                if cell and any(p in str(cell) for p in pats):
                    col_mapping[col_name] = i
                    break
        return col_mapping

    def _find_subtable_data_end(self, table: List[List], start_idx: int, reference_numbers: List[str]) -> int:
        """Finds the end of a subtable's data rows."""
        for i in range(start_idx, len(table)):
            row = table[i]
            if not row:
                continue
            if self._find_reference_in_row(row, reference_numbers) or self._is_subtable_end_row(row):
                return i
        return len(table)

    def _is_subtable_end_row(self, row: List) -> bool:
        """Checks if a row indicates the end of a subtable."""
        if not row:
            return True
        cell_text = "".join(str(c) for c in row if c)
        return any(word in cell_text for word in ["合計", "小計", "総計", "計"])

    def _process_subtable_data_rows_with_spanning(self, data_rows: List[List], col_mapping: Dict[str, int],
                                                  page_num: int, reference_number: Optional[str],
                                                  start_row_num: int) -> List[SubtableItem]:
        """Processes subtable data rows with corrected row spanning logic."""
        subtable_items = []
        i = 0
        while i < len(data_rows):
            if self._is_subtable_end_row(data_rows[i]):
                break
            item_result = self._extract_subtable_item_with_spanning(
                data_rows, i, col_mapping, page_num, reference_number)
            if item_result:
                subtable_items.append(item_result['item'])
                i = item_result['next_index']
            else:
                i += 1
        return subtable_items

    def _extract_subtable_item_with_spanning(self, data_rows: List[List], start_idx: int, col_mapping: Dict[str, int],
                                             page_num: int, reference_number: Optional[str]) -> Optional[Dict]:
        """
        Extracts a single logical subtable item, which may span multiple rows.
        Enhanced: Handles cases where first row has only name, and second row has unit (and possibly name/quantity).
        """
        if start_idx >= len(data_rows):
            return None

        # --- First row ---
        row1 = data_rows[start_idx]
        row1_fields, row1_quantity, row1_unit = self._extract_fields_from_row(
            row1, col_mapping)
        row1_name = row1_fields.get("名称・規格")

        # --- Second row (if present) ---
        row2_fields, row2_quantity, row2_unit, row2_name = None, None, None, None
        if start_idx + 1 < len(data_rows):
            row2 = data_rows[start_idx + 1]
            row2_fields, row2_quantity, row2_unit = self._extract_fields_from_row(
                row2, col_mapping)
            row2_name = row2_fields.get("名称・規格")

        # 1. If first row has both name and unit, treat as complete item (quantity=0 if blank)
        if row1_name and row1_unit:
            item_key = self._create_subtable_item_key(
                row1_fields, reference_number)
            subtable_item = SubtableItem(
                item_key=item_key,
                raw_fields=row1_fields,
                quantity=row1_quantity if row1_quantity > 0 else 0.0,
                unit=row1_unit,
                source="PDF",
                page_number=page_num + 1,
                reference_number=reference_number
            )
            return {'item': subtable_item, 'next_index': start_idx + 1}

        # 2. If first row has only name, and second row has unit (and possibly name/quantity)
        if row1_name and not row1_unit and row2_unit:
            # Concatenate names if both present
            combined_name = row1_name
            if row2_name:
                combined_name = f"{row1_name} {row2_name}"
            # Build raw_fields
            combined_fields = dict(row1_fields)
            combined_fields["名称・規格"] = combined_name
            combined_fields["単位"] = row2_unit
            if row2_quantity is not None:
                combined_fields["数量"] = str(
                    row2_quantity) if row2_quantity > 0 else ""
            # Quantity: use row2_quantity, treat blank as 0
            quantity = row2_quantity if row2_quantity and row2_quantity > 0 else 0.0
            item_key = self._create_subtable_item_key(
                combined_fields, reference_number)
            subtable_item = SubtableItem(
                item_key=item_key,
                raw_fields=combined_fields,
                quantity=quantity,
                unit=row2_unit,
                source="PDF",
                page_number=page_num + 1,
                reference_number=reference_number
            )
            return {'item': subtable_item, 'next_index': start_idx + 2}

        # --- Fallback to original logic for other cases ---
        item_name, unit, quantity = None, None, 0.0
        raw_fields = {}
        rows_processed = 0
        for offset in range(min(3, len(data_rows) - start_idx)):
            current_row_idx = start_idx + offset
            row = data_rows[current_row_idx]
            if offset > 0 and self._is_new_item_row(row, col_mapping, item_name):
                break
            current_row_fields, current_row_quantity, current_row_unit = self._extract_fields_from_row(
                row, col_mapping)
            if not item_name and "名称・規格" in current_row_fields:
                item_name = current_row_fields["名称・規格"]
            if not unit and current_row_unit:
                unit = current_row_unit
            if quantity == 0.0 and current_row_quantity > 0.0:
                quantity = current_row_quantity
            for k, v in current_row_fields.items():
                if k not in raw_fields:
                    raw_fields[k] = v
            # If the first row has name and unit, it's a complete item.
            if offset == 0 and item_name and unit:
                rows_processed = 1
                break
            rows_processed = offset + 1
        if not item_name or (any(word in item_name for word in ["合計", "小計", "総計"]) and not unit and quantity == 0.0):
            return None
        item_key = self._create_subtable_item_key(raw_fields, reference_number)
        subtable_item = SubtableItem(item_key=item_key, raw_fields=raw_fields, quantity=quantity,
                                     unit=unit, source="PDF", page_number=page_num + 1, reference_number=reference_number)
        return {'item': subtable_item, 'next_index': start_idx + rows_processed}

    def _is_new_item_row(self, row: List, col_mapping: Dict[str, int], previous_item_name: Optional[str]) -> bool:
        """Checks if a row represents a new item."""
        if not previous_item_name:
            return False
        name_col_idx = col_mapping.get("名称・規格")
        if name_col_idx is not None and name_col_idx < len(row) and row[name_col_idx]:
            current_cell_value = str(row[name_col_idx]).strip()
            if current_cell_value and current_cell_value != previous_item_name:
                return True
        return False

    def _create_subtable_item_key(self, raw_fields: Dict[str, str], reference_number: Optional[str]) -> str:
        """Creates item key for subtable items."""
        return raw_fields.get("名称・規格", "").strip()
