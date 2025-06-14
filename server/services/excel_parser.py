import pandas as pd
import re
from typing import List, Dict, Optional, Union
from io import BytesIO
from ..schemas.tender import TenderItem


class ExcelParser:
    def __init__(self):
        # Flexible column patterns for Japanese construction documents
        self.column_patterns = {
            "工種": ["工種", "工 種", "費目", "工事区分"],
            "種別": ["種別", "種 別"],
            "細別": ["細別", "細 別"],
            "規格": ["規格", "規 格"],
            "名称": ["名称", "名 称", "項目", "品名"],
            "単位": ["単位", "単 位"],
            "数量": ["数量", "数 量"],
            "単価": ["単価", "単 価"],
            "金額": ["金額", "金 額"],
            "摘要": ["摘要", "備考", "摘 要"]
        }

    def _find_header_row(self, df: pd.DataFrame) -> Optional[int]:
        """
        Find the row that contains column headers
        """
        for i in range(min(10, len(df))):  # Check first 10 rows
            row = df.iloc[i]
            row_str = " ".join([str(val) for val in row if pd.notna(val)])
            # Normalize for search
            row_str = row_str.replace('\u3000', '').replace(
                ' ', '').replace('　', '')

            # Look for Japanese construction document patterns
            # Also look for the specific pattern from the Excel file
            if (any(pattern in row_str for pattern in ["工種", "種別", "細別", "数量", "単位", "費目"]) or
                "費目/工種/種別/細別/規格" in row_str or
                    ("費目" in row_str and "工種" in row_str and "種別" in row_str)):
                return i
        return None

    def _find_column_mapping(self, header_row: pd.Series) -> Dict[str, int]:
        """
        Map column names to indices based on patterns
        """
        col_mapping = {}

        for col_idx, cell_value in enumerate(header_row):
            if pd.isna(cell_value):
                continue

            cell_str = str(cell_value).strip()
            # Handle Unicode whitespace and normalize
            cell_str = cell_str.replace('\u3000', '').replace(
                ' ', '').replace('　', '')

            # Check each pattern
            for standard_name, patterns in self.column_patterns.items():
                for pattern in patterns:
                    if pattern in cell_str:
                        col_mapping[standard_name] = col_idx
                        break
                if standard_name in col_mapping:
                    break

            # Special handling for quantity column with Unicode spaces
            if not any(name == "数量" for name in col_mapping.keys()):
                original_str = str(cell_value).strip()
                if ('数' in original_str and '量' in original_str) or '数\u3000量' in original_str:
                    col_mapping["数量"] = col_idx

        return col_mapping

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
        Process a single Excel sheet
        """
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

            # Create item key from hierarchical fields
            key_parts = []
            for field in ["工種", "種別", "細別", "名称", "規格"]:
                if field in raw_fields and raw_fields[field]:
                    key_parts.append(raw_fields[field])

            # If no hierarchical fields, try to create key from available data
            if not key_parts:
                for field_name, field_value in raw_fields.items():
                    if field_value and field_name not in ["単位", "数量", "単価", "金額", "摘要"]:
                        key_parts.append(field_value)

            item_key = "|".join(
                key_parts) if key_parts else f"{sheet_name}_row_{row_idx}"

            items.append(TenderItem(
                item_key=item_key,
                raw_fields=raw_fields,
                quantity=quantity,
                source="Excel"
            ))

        return items

    def extract_items_from_buffer(self, excel_buffer: BytesIO) -> List[TenderItem]:
        """
        Extract items from Excel buffer (in-memory) to avoid file locking issues.
        """
        items = []
        excel_file = None

        try:
            # Read from buffer
            excel_file = pd.ExcelFile(excel_buffer)
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
            print(f"Error reading Excel buffer: {e}")
        finally:
            # Ensure proper cleanup
            if excel_file is not None:
                try:
                    excel_file.close()
                except Exception as e:
                    print(f"Error closing Excel file: {e}")

        print(f"Total items extracted from Excel: {len(items)}")
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
