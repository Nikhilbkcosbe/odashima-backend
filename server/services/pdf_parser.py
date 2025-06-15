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
        self._setup_tesseract()

    def _setup_tesseract(self):
        """
        Setup Tesseract environment variables for Lambda.
        """
        # Check if we're running in Lambda
        if os.environ.get('AWS_LAMBDA_FUNCTION_NAME'):
            # Lambda layer paths
            os.environ['LD_LIBRARY_PATH'] = '/opt/lib:' + \
                os.environ.get('LD_LIBRARY_PATH', '')
            os.environ['TESSDATA_PREFIX'] = '/opt/tessdata'

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
        logger.info(f"Page range: {start_page or 'start'} to {end_page or 'end'}")
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                logger.info(f"PDF has {total_pages} pages total")
                
                # Determine actual page range
                actual_start = (start_page - 1) if start_page is not None else 0
                actual_end = (end_page - 1) if end_page is not None else total_pages - 1
                
                # Validate page range
                actual_start = max(0, actual_start)
                actual_end = min(total_pages - 1, actual_end)
                
                if actual_start > actual_end:
                    logger.warning(f"Invalid page range: start={actual_start+1}, end={actual_end+1}")
                    return all_items
                
                pages_to_process = actual_end - actual_start + 1
                logger.info(f"Processing pages {actual_start + 1} to {actual_end + 1} ({pages_to_process} pages)")
                
                # Process specified page range iteratively
                for page_num in range(actual_start, actual_end + 1):
                    page = pdf.pages[page_num]
                    logger.info(f"Processing page {page_num + 1}/{total_pages}")
                    
                    page_items = self._extract_tables_from_page(page, page_num)
                    
                    logger.info(f"Extracted {len(page_items)} items from page {page_num + 1}")
                    
                    # Join items from this page to the total collection
                    all_items.extend(page_items)
                    
                logger.info(f"Total items extracted from PDF (pages {actual_start + 1}-{actual_end + 1}): {len(all_items)}")
                
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
                logger.info(f"Processing table {table_num + 1}/{len(tables)} on page {page_num + 1}")
                
                table_items = self._process_single_table(table, page_num, table_num)
                
                logger.info(f"Extracted {len(table_items)} items from table {table_num + 1}")
                
                # Join table items to page items
                page_items.extend(table_items)
                
        except Exception as e:
            logger.error(f"Error processing page {page_num + 1}: {e}")
        
        return page_items
    
    def _process_single_table(self, table: List[List], page_num: int, table_num: int) -> List[TenderItem]:
        """
        Process a single table and extract all valid items from it.
        """
        items = []
        
        if not table or len(table) < 2:
            logger.warning(f"Table {table_num + 1} on page {page_num + 1} is too small (less than 2 rows)")
            return items
        
        # Find header row
        header_row, header_idx = self._find_header_row(table)
        
        if header_row is None:
            logger.warning(f"No header row found in table {table_num + 1} on page {page_num + 1}")
            return items
        
        logger.info(f"Found header at row {header_idx + 1} in table {table_num + 1}")
        
        # Get column mapping
        col_indices = self._get_column_mapping(header_row)
        
        if not col_indices:
            logger.warning(f"No recognizable columns found in table {table_num + 1} on page {page_num + 1}")
            return items
        
        logger.info(f"Column mapping for table {table_num + 1}: {col_indices}")
        
        # Process data rows iteratively with row spanning logic
        data_rows = table[header_idx + 1:]
        logger.info(f"Processing {len(data_rows)} data rows in table {table_num + 1}")
        
        for row_idx, row in enumerate(data_rows):
            try:
                result = self._process_single_row_with_spanning(
                    row, col_indices, page_num, table_num, header_idx + 1 + row_idx, items
                )
                
                if result == "merged":
                    logger.debug(f"Row {row_idx + 1} merged with previous item (quantity-only row)")
                elif result == "skipped":
                    logger.debug(f"Row {row_idx + 1} skipped (empty row)")
                elif result:
                    items.append(result)
                    
            except Exception as e:
                logger.error(f"Error processing row {row_idx + 1} in table {table_num + 1}: {e}")
                continue
        
        return items
    
    def _process_single_row_with_spanning(self, row: List, col_indices: Dict[str, int], 
                                        page_num: int, table_num: int, row_num: int, 
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
        
        for col_name, col_idx in col_indices.items():
            if col_idx < len(row) and row[col_idx]:
                cell_value = str(row[col_idx]).strip()
                if cell_value:
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
        
        # Create new item with hierarchical key
        item_key = self._create_item_key_from_fields(raw_fields)
        
        # Skip if we couldn't create a meaningful key
        if not item_key:
            return "skipped"
        
        return TenderItem(
            item_key=item_key,
            raw_fields=raw_fields,
            quantity=quantity,
            source="PDF"
        )
    
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
            logger.warning("Quantity-only row found but no previous item to merge with")
            return "skipped"
        
        # Add quantity to the last item
        last_item = existing_items[-1]
        old_quantity = last_item.quantity
        new_quantity = old_quantity + quantity
        
        # Update the last item's quantity
        last_item.quantity = new_quantity
        
        logger.info(f"Merged quantity: {last_item.item_key} - {old_quantity} + {quantity} = {new_quantity}")
        
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
    
    def _find_header_row(self, table: List[List]) -> Tuple[Optional[List], int]:
        """
        Find the header row in the table.
        """
        # Check first 3 rows for headers
        for i, row in enumerate(table[:3]):
            if row and any(cell and "名称" in str(cell) or "工種" in str(cell) or "数量" in str(cell) for cell in row):
                return row, i
        
        # If no clear header found, use first row
        return table[0] if table else None, 0
    
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
        value_str = str(cell_value).replace(",", "").replace(" ", "").replace("　", "")
        
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
        result = self._process_single_row_with_spanning(row, col_indices, page_num, table_num, row_num, [])
        
        if isinstance(result, TenderItem):
            return result
        else:
            return None
