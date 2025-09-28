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

        # Kitakami-specific column patterns
        self.kitakami_column_patterns = {
            "費目・工種・種別・細": ["費 目 ・ 工 種 ・ 種 別 ・ 細", "費目・工種・種別・細別・規格"],
            "数量": ["数量", "数 量"],
            "単位": ["単位", "単 位"],
            "明細単価番号": ["明細単価番号", "明 細 単 価 番 号"]
        }

        self.column_patterns = self.default_column_patterns.copy()

    def extract_tables_with_range(self, pdf_path: str, start_page: Optional[int] = None, end_page: Optional[int] = None, project_area: str = "岩手") -> List[TenderItem]:
        """
        Extract tables from PDF iteratively with specified page range.
        This is the main entry point for parsing the main table.
        """
        all_items = []
        logger.info(f"Starting PDF extraction from: {pdf_path}")
        logger.info(
            f"Page range: {start_page or 'start'} to {end_page or 'end'}")
        logger.info(f"Project area: {project_area}")

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
                    page_items = self._extract_tables_from_page(
                        page, page_num, project_area)
                    all_items.extend(page_items)

        except Exception as e:
            logger.error(
                f"Error processing PDF for main table: {e}", exc_info=True)
            raise
        return all_items

    def _extract_tables_from_page(self, page, page_num: int, project_area: str = "岩手") -> List[TenderItem]:
        """Extract all tables from a single page."""
        page_items = []
        try:
            tables = page.extract_tables()
            logger.info(f"Found {len(tables)} tables on page {page_num + 1}")
            for table_num, table in enumerate(tables):
                page_items.extend(self._process_single_table(
                    table, page_num, table_num, project_area))
        except Exception as e:
            logger.error(
                f"Error processing page {page_num + 1}: {e}", exc_info=True)
        return page_items

    def _process_single_table(self, table: List[List], page_num: int, table_num: int, project_area: str = "岩手") -> List[TenderItem]:
        """Process a single table and extract all valid items from it."""
        items = []
        if not table or len(table) < 2:
            return items

        header_row, header_idx = self._find_header_row(table)
        if header_row is None:
            return items

        # Determine effective project area from header if possible
        effective_area = self._detect_project_area_from_header(
            header_row) or project_area

        # Build column mapping (attempt with effective area first, then fallback to other patterns)
        col_indices = self._get_column_mapping(header_row, effective_area)
        if not col_indices:
            # Fallback: try the opposite area's patterns just in case
            fallback_area = "北上市" if effective_area == "岩手" else "岩手"
            col_indices = self._get_column_mapping(header_row, fallback_area)
            if not col_indices:
                return items
            effective_area = fallback_area

        data_rows = table[header_idx + 1:]
        for row_idx, row in enumerate(data_rows):
            try:
                result = self._process_single_row_with_spanning(
                    row, col_indices, page_num, table_num, header_idx + 1 + row_idx, items, effective_area)
                if isinstance(result, TenderItem):
                    items.append(result)
            except Exception as e:
                logger.error(
                    f"Error processing row {row_idx + 1} in table {table_num + 1}: {e}", exc_info=True)
        return items

    def _process_single_row_with_spanning(self, row: List, col_indices: Dict[str, int],
                                          page_num: int, table_num: int, row_num: int,
                                          existing_items: List, project_area: str = "岩手") -> Union[TenderItem, str, None]:
        """Handles row spanning for the main table."""
        if self._is_completely_empty_row(row):
            return "skipped"

        raw_fields, quantity, unit = self._extract_fields_from_row(
            row, col_indices, project_area)

        has_item_fields = self._has_item_identifying_fields(
            raw_fields, project_area)
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

    def _extract_fields_from_row(self, row: List, col_indices: Dict[str, int], project_area: str = "岩手") -> Tuple[Dict[str, str], float, Optional[str]]:
        """Extracts all relevant fields from a single row."""
        raw_fields = {}
        quantity = 0.0
        unit = None

        # For Kitakami projects, ignore rows with "合計" (total) in the item name
        if project_area == "北上市":
            item_name_col = col_indices.get("費目・工種・種別・細", 0)
            if item_name_col < len(row) and row[item_name_col]:
                item_name = str(row[item_name_col]).strip()
                if "合計" in item_name:
                    # Return empty fields for total rows
                    return {}, 0.0, None

        for col_name, col_idx in col_indices.items():
            if col_idx < len(row) and row[col_idx]:
                cell_value = str(row[col_idx]).strip()
                if cell_value:
                    if col_name == "数量":
                        if project_area == "北上市":
                            # For Kitakami, pass row and column index for adjacent column reconstruction
                            quantity = self._extract_kitakami_quantity(
                                cell_value, row, col_idx)
                        else:
                            quantity = self._extract_quantity(
                                cell_value, project_area)
                    elif col_name == "単位":
                        unit = cell_value
                    raw_fields[col_name] = cell_value
        return raw_fields, quantity, unit

    def _has_item_identifying_fields(self, raw_fields: Dict[str, str], project_area: str = "岩手") -> bool:
        """Checks if the row contains fields that identify an item."""
        if project_area == "北上市":
            # Kitakami-specific identifying fields
            identifying_fields = ["費目・工種・種別・細", "明細単価番号"]
        else:
            # Iwate-specific identifying fields
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

    def _get_column_mapping(self, header_row: List, project_area: str = "岩手") -> Dict[str, int]:
        """Maps column names to indices based on header row."""
        col_indices = {}

        # Choose patterns set and also support partial header variants
        pattern_sets = []
        if project_area == "北上市":
            pattern_sets = [
                self.kitakami_column_patterns, self.column_patterns]
        else:
            pattern_sets = [self.column_patterns,
                            self.kitakami_column_patterns]

        for patterns_to_use in pattern_sets:
            tentative = {}
            for col_name, patterns in patterns_to_use.items():
                for i, cell in enumerate(header_row):
                    if cell and any(p in str(cell) for p in patterns):
                        tentative[col_name] = i
                        break
            # Require at least quantity and unit to proceed
            if ("数量" in tentative) and ("単位" in tentative):
                # Also keep any available name/spec/remarks columns
                col_indices = tentative
                break

        return col_indices

    def _detect_project_area_from_header(self, header_row: List) -> Optional[str]:
        """Rudimentary detection of project area based on distinctive headers."""
        try:
            header_text = "|".join([str(c) for c in header_row if c])
            # Kitakami headers often include 明細単価番号 and compact 費 目 ・ 工 種 ・ 種 別 ・ 細
            if any(p in header_text for p in self.kitakami_column_patterns.get("明細単価番号", [])) or \
               any(p in header_text for p in self.kitakami_column_patterns.get("費目・工種・種別・細", [])):
                return "北上市"
        except Exception:
            pass
        return None

    def _extract_quantity(self, cell_value, project_area: str = "岩手") -> float:
        """Extracts numeric quantity from a cell value."""
        if not cell_value:
            return 0.0

        # For Kitakami projects, use special decimal extraction logic
        if project_area == "北上市":
            return self._extract_kitakami_quantity(cell_value)

        # Standard Iwate extraction logic
        value_str = str(cell_value).replace(",", "")
        number_match = re.search(r'[\d.]+', value_str)
        if number_match:
            try:
                return float(number_match.group())
            except ValueError:
                pass
        return 0.0

    def _extract_kitakami_quantity(self, cell_value, row: List = None, qty_idx: int = None) -> float:
        """
        Extract quantity with special Kitakami decimal handling.
        The quantity column is internally divided into normal digits and decimal digits.
        For Kitakami: adjacent columns contain integer part and decimal part (e.g., "1" and "0.5" -> 1.5)
        """
        try:
            if not cell_value:
                return 0.0

            # First, try to get quantity from the main cell
            qty_text = self._normalize_text(str(cell_value))
            quantity = self._extract_number_from_text(qty_text)
            if quantity is not None:
                # Check if this is a standalone integer that might need decimal reconstruction
                if quantity == int(quantity) and row is not None and qty_idx is not None:
                    # Look for decimal part in adjacent columns
                    decimal_part = self._find_adjacent_decimal_part(
                        row, qty_idx)
                    if decimal_part is not None:
                        return float(f"{int(quantity)}.{decimal_part}")
                return quantity

            # Look for decimal patterns in the cell
            decimal_match = re.search(r'(\d+)\.(\d+)', qty_text)
            if decimal_match:
                try:
                    return float(decimal_match.group(0))
                except ValueError:
                    pass

        except Exception as e:
            logger.warning(f"Error extracting Kitakami quantity: {str(e)}")

        return 0.0

    def _find_adjacent_decimal_part(self, row: List, qty_idx: int) -> Optional[str]:
        """
        Find decimal part in adjacent columns for Kitakami quantity reconstruction.
        Looks for patterns like "0.5", "0.06", "0.006" in adjacent cells.
        Also handles cases where decimal part starts with "." like ".06"
        Only checks immediately adjacent columns to avoid false matches from item descriptions.
        """
        try:
            # Only check immediately adjacent columns (left and right)
            for offset in [-1, 1]:
                check_idx = qty_idx + offset
                if 0 <= check_idx < len(row) and row[check_idx]:
                    cell_text = self._normalize_text(str(row[check_idx]))

                    # Skip if the cell contains text that looks like item description
                    if self._is_description_text(cell_text):
                        continue

                    # Look for decimal patterns starting with "0."
                    decimal_match = re.search(r'0\.(\d+)', cell_text)
                    if decimal_match:
                        return decimal_match.group(1)

                    # Look for decimal patterns starting with "."
                    dot_decimal_match = re.search(r'\.(\d+)', cell_text)
                    if dot_decimal_match:
                        return dot_decimal_match.group(1)

                    # Look for patterns like "5", "06", "006" that could be decimal parts
                    if re.match(r'^\d+$', cell_text):
                        # If it's a small number, it might be a decimal part
                        if len(cell_text) <= 3:  # 0.5, 0.06, 0.006
                            return cell_text

        except Exception as e:
            logger.warning(f"Error finding adjacent decimal part: {str(e)}")

        return None

    def _is_description_text(self, text: str) -> bool:
        """
        Check if text looks like item description rather than a quantity.
        Returns True if the text contains description-like patterns.
        """
        if not text:
            return False

        # Check for patterns that indicate this is description text
        description_patterns = [
            r'[A-Za-z]',  # Contains letters
            r'[=]',       # Contains equals sign (like L=12.46m)
            r'[()]',      # Contains parentheses (like (40t))
            r'[kN]',      # Contains units like kN
            r'[m]',       # Contains units like m
            r'[t]',       # Contains units like t
            r'[号]',      # Contains Japanese characters
            r'[明]',      # Contains Japanese characters
        ]

        for pattern in description_patterns:
            if re.search(pattern, text):
                return True

        return False

    def _normalize_text(self, text: str) -> str:
        """Normalize text by removing spaces and handling full-width/half-width."""
        if not text:
            return ""
        # Remove all spaces and normalize
        return re.sub(r'\s+', '', str(text))

    def _extract_number_from_text(self, text: str) -> Optional[float]:
        """Extract number from text."""
        if not text:
            return None

        # Look for decimal numbers
        decimal_match = re.search(r'(\d+\.?\d*)', text)
        if decimal_match:
            try:
                return float(decimal_match.group(1))
            except ValueError:
                pass

        return None

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

                        # Include Kitakami-only code column when present in the subtable row
                        try:
                            code_value = (row.get("明細単価番号", "") or "").strip()
                            if code_value:
                                raw_fields["明細単価番号"] = code_value
                        except Exception:
                            pass

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
