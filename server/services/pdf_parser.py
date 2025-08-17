from subtable_pdf_extractor import SubtablePDFExtractor, extract_subtables_api
import pdfplumber
import os
import re
import logging
from typing import List, Dict, Tuple, Optional, Union
from ..schemas.tender import TenderItem, SubtableItem

# Import the new subtable extractor
import sys
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..'))

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
        """Creates a concatenated item key using space concatenation (consistent with Excel)."""
        key_fields = ["工事区分・工種・種別・細別", "摘要"]
        base_key = next(
            (raw_fields[f] for f in key_fields if f in raw_fields and raw_fields[f]), None)
        if not base_key:
            base_key = next((v for k, v in raw_fields.items() if v and k not in [
                            "単位", "数量", "単価", "金額", "規格"]), "")

        # Use space concatenation instead of + to match Excel behavior
        if "規格" in raw_fields and raw_fields["規格"]:
            return f"{base_key} {raw_fields['規格']}".strip()
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
        NEW: Extract subtables using the API-ready subtable extractor and convert to SubtableItem format.
        This replaces the old subtable extraction logic completely.

        Args:
            pdf_path: Path to the PDF file
            start_page: Starting page number (1-based, None means start from page 1)
            end_page: Ending page number (1-based, None means extract all pages)
            reference_numbers: List of reference numbers to filter (not used in new implementation)

        Returns:
            List of SubtableItem objects
        """
        logger.info("=== USING NEW API-READY PDF SUBTABLE EXTRACTOR ===")
        logger.info(f"PDF file: {pdf_path}")
        logger.info(f"Page range: {start_page} to {end_page}")

        all_subtable_items = []

        try:
            # Determine page range
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)

                # Default to full range if not specified
                actual_start = start_page if start_page is not None else 1
                actual_end = end_page if end_page is not None else total_pages

                # Validate page range
                actual_start = max(1, actual_start)
                actual_end = min(total_pages, actual_end)

                logger.info(
                    f"Processing pages {actual_start} to {actual_end} of {total_pages} total pages")

            # Use the new API-ready subtable extractor
            extractor = SubtablePDFExtractor()
            result = extractor.extract_subtables_from_pdf(
                pdf_path, actual_start, actual_end)

            if "error" in result:
                logger.error(
                    f"New subtable extractor failed: {result['error']}")
                return []

            logger.info(
                f"NEW API extracted {result['total_subtables']} subtables with {result['total_rows']} total rows")

            # Convert the new API response to SubtableItem format
            for subtable in result.get("subtables", []):
                reference_number = subtable.get("reference_number", "")
                page_number = subtable.get("page_number", 0)
                rows = subtable.get("rows", [])

                for row in rows:
                    try:
                        # Extract data from the new format
                        item_name = row.get("名称・規格", "").strip()
                        unit = row.get("単位", "").strip()
                        quantity_str = row.get("数量", "").strip()
                        remarks = row.get("摘要", "").strip()

                        # Parse quantity
                        try:
                            # Handle various number formats and commas
                            quantity = float(quantity_str.replace(
                                ',', '').replace('，', '')) if quantity_str else 0.0
                        except (ValueError, TypeError):
                            quantity = 0.0

                        # Create raw_fields dictionary matching the expected format
                        raw_fields = {
                            "名称・規格": item_name,
                            "単位": unit,
                            "数量": quantity_str,
                            "摘要": remarks,
                            "参照番号": reference_number
                        }

                        # Create SubtableItem only if we have a valid item name
                        if item_name:
                            # Get table title from the subtable
                            table_title = subtable.get("table_title", None)

                            subtable_item = SubtableItem(
                                item_key=item_name,
                                raw_fields=raw_fields,
                                quantity=quantity,
                                unit=unit or None,
                                source="PDF",
                                page_number=page_number,
                                reference_number=reference_number,
                                sheet_name=None,  # PDF doesn't have sheet names
                                table_title=table_title
                            )
                            all_subtable_items.append(subtable_item)

                    except Exception as e:
                        logger.error(
                            f"Error converting subtable row to SubtableItem: {e}")
                        logger.error(f"Row data: {row}")
                        continue

            logger.info(
                f"Successfully converted {len(all_subtable_items)} subtable items using NEW API")

        except Exception as e:
            logger.error(f"Error using new API subtable extractor: {e}")
            logger.error(
                "NEW subtable extraction failed - returning empty list")
            return []

        return all_subtable_items
