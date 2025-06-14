import pdfplumber
import os
import re
from typing import List, Dict
from ..schemas.tender import TenderItem


class PDFParser:
    def __init__(self):
        # More flexible column patterns for Japanese construction documents
        self.column_patterns = {
            "名称・規格": ["名称・規格", "名称", "規格", "項目", "品名"],
            "単位": ["単位", "単 位"],
            "数量": ["数量", "数 量"],
            "摘要": ["摘要", "備考", "摘 要"],
            "工種": ["工種", "工 種", "費目", "工事区分"],
            "種別": ["種別", "種 別"],
            "細別": ["細別", "細 別"],
            "単価": ["単価", "単 価"],
            "金額": ["金額", "金 額"]
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

    def _is_valid_data_row(self, row: List) -> bool:
        """
        Check if row contains valid data (not empty or header)
        """
        if not row:
            return False
            
        # Check if at least one cell has meaningful content
        meaningful_cells = 0
        for cell in row:
            if cell and str(cell).strip() and str(cell).strip() not in ["", "None", "nan"]:
                meaningful_cells += 1
                
        return meaningful_cells >= 2  # At least 2 meaningful cells

    def extract_tables(self, pdf_path: str) -> List[TenderItem]:
        """
        Extract tables from PDF and convert to TenderItem objects.
        """
        items = []

        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                try:
                    tables = page.extract_tables()

                    for table_num, table in enumerate(tables):
                        if not table or len(table) < 2:
                            continue

                        # Find the header row (might not be the first row)
                        header_row = None
                        header_idx = -1
                        
                        for i, row in enumerate(table[:3]):  # Check first 3 rows for headers
                            if row and any(cell and "名称" in str(cell) for cell in row):
                                header_row = row
                                header_idx = i
                                break
                        
                        if not header_row:
                            # Try to use first row as header anyway
                            header_row = table[0]
                            header_idx = 0

                        # Find column indices
                        col_indices = {}
                        for col_name in ["名称・規格", "単位", "数量", "摘要", "工種", "種別", "細別"]:
                            idx = self._find_column_index(header_row, col_name)
                            if idx != -1:
                                col_indices[col_name] = idx

                        # Process data rows
                        for row_idx, row in enumerate(table[header_idx + 1:], start=header_idx + 1):
                            if not self._is_valid_data_row(row):
                                continue

                            # Extract fields
                            raw_fields = {}
                            for col_name, col_idx in col_indices.items():
                                if col_idx < len(row) and row[col_idx]:
                                    raw_fields[col_name] = str(row[col_idx]).strip()

                            # Skip if no meaningful fields extracted
                            if not raw_fields:
                                continue

                            # Extract quantity
                            quantity = 0.0
                            if "数量" in col_indices and col_indices["数量"] < len(row):
                                quantity = self._extract_quantity(row[col_indices["数量"]])

                            # Create item key from available fields
                            key_parts = []
                            for field in ["工種", "種別", "細別", "名称・規格"]:
                                if field in raw_fields and raw_fields[field]:
                                    key_parts.append(raw_fields[field])
                            
                            item_key = "|".join(key_parts) if key_parts else f"row_{page_num}_{table_num}_{row_idx}"

                            items.append(TenderItem(
                                item_key=item_key,
                                raw_fields=raw_fields,
                                quantity=quantity,
                                source="PDF"
                            ))

                except Exception as e:
                    print(f"Error processing page {page_num}: {e}")
                    continue

        return items
