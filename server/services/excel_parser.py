import pandas as pd
import logging
from typing import List, Dict, Any, Optional
import os
import re
from dataclasses import dataclass

logger = logging.getLogger(__name__)


@dataclass
class LogicalRow:
    item_name: str
    unit: str
    quantity: str
    unit_price: str
    amount: str
    notes: str
    level: int


class ExcelParser:
    """Service for parsing Excel files and extracting hierarchical data"""

    def __init__(self):
        self.upload_folder = "uploads"
        os.makedirs(self.upload_folder, exist_ok=True)

    def extract_hierarchical_data(self, file_path: str, sheet_name: str, project_area: str = "岩手") -> Optional[List[Dict]]:
        """
        Extract hierarchical data from Excel file

        Args:
            file_path: Path to the Excel file
            sheet_name: Name of the sheet to extract from

        Returns:
            List of hierarchical items or None if extraction fails
        """
        try:
            logger.info(
                f"Extracting hierarchical data from {file_path}, sheet: {sheet_name}")

            # Read the Excel file
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

            # Extract logical rows with spanning
            logical_rows = self._extract_logical_rows_with_spanning(
                df, project_area)

            # Build hierarchy
            hierarchical_data = self._build_hierarchy(logical_rows)

            return hierarchical_data

        except Exception as e:
            logger.error(f"Error extracting hierarchical data: {str(e)}")
            return None

    def _get_cell_value(self, cell_value, preserve_spaces=False):
        """Get cell value with optional space preservation"""
        if pd.isna(cell_value):
            return ""
        value = str(cell_value)
        if preserve_spaces:
            return value
        return value.strip()

    def _get_hierarchy_level(self, item_name: str) -> int:
        """Get hierarchy level based on indentation"""
        if not item_name:
            return 0
        # Count full-width spaces (Japanese indentation)
        return item_name.count('\u3000')

    def _extract_single_logical_row(self, df: pd.DataFrame, start_row: int, project_area: str = "岩手") -> Optional[LogicalRow]:
        """Extract a single logical row with spanning"""
        try:
            # Get the first row of the logical row
            row_data = df.iloc[start_row]

            # Initialize with first row data
            item_name = self._get_cell_value(row_data[1], preserve_spaces=True)
            unit = self._get_cell_value(row_data[3])
            quantity = self._get_cell_value(row_data[4])
            unit_price = self._get_cell_value(row_data[5])
            amount = self._get_cell_value(row_data[6])

            # For Kitakami projects, extract 摘要 (notes) from column 7
            if project_area == "北上市":
                notes = self._get_cell_value(
                    row_data[7]) if len(row_data) > 7 else ""
            else:
                notes = self._get_cell_value(
                    row_data[7]) if len(row_data) > 7 else ""

            # Check for spanning in subsequent rows
            next_row = start_row + 1
            while next_row < len(df):
                next_row_data = df.iloc[next_row]

                # Check if this is a continuation row (empty item_name but has other data)
                next_item = self._get_cell_value(next_row_data[1])
                if next_item.strip() == "":
                    # Merge data from continuation row
                    if not unit and self._get_cell_value(next_row_data[3]):
                        unit = self._get_cell_value(next_row_data[3])
                    if not quantity and self._get_cell_value(next_row_data[4]):
                        quantity = self._get_cell_value(next_row_data[4])
                    if not unit_price and self._get_cell_value(next_row_data[5]):
                        unit_price = self._get_cell_value(next_row_data[5])
                    if not amount and self._get_cell_value(next_row_data[6]):
                        amount = self._get_cell_value(next_row_data[6])
                    if not notes and len(next_row_data) > 7 and self._get_cell_value(next_row_data[7]):
                        notes = self._get_cell_value(next_row_data[7])

                    # Also check for additional item_name content
                    if self._get_cell_value(next_row_data[1], preserve_spaces=True).strip():
                        item_name += " " + \
                            self._get_cell_value(
                                next_row_data[1], preserve_spaces=True)

                    next_row += 1
                else:
                    break

            # Determine hierarchy level
            level = self._get_hierarchy_level(item_name)

            return LogicalRow(
                item_name=item_name.strip(),
                unit=unit,
                quantity=quantity,
                unit_price=unit_price,
                amount=amount,
                notes=notes,
                level=level
            )

        except Exception as e:
            logger.error(
                f"Error extracting logical row at {start_row}: {str(e)}")
            return None

    def _extract_logical_rows_with_spanning(self, df: pd.DataFrame, project_area: str = "岩手") -> List[LogicalRow]:
        """Extract all logical rows with spanning from the dataframe"""
        logical_rows = []
        row_index = 0

        while row_index < len(df):
            row_data = df.iloc[row_index]

            # Skip empty rows
            if row_data.isna().all():
                row_index += 1
                continue

            # Check if this is a table number row (just a number)
            first_cell = self._get_cell_value(row_data[1])
            if first_cell and first_cell.strip().isdigit():
                row_index += 1
                continue

            # Check if this is a header row (contains headers like "項目", "単位", etc.)
            if any(header in str(cell).lower() for cell in row_data if pd.notna(cell)
                   for header in ["項目", "単位", "数量", "単価", "金額"]):
                row_index += 1
                continue

            # Extract logical row
            logical_row = self._extract_single_logical_row(
                df, row_index, project_area)
            if logical_row and logical_row.item_name.strip():
                # Skip header-like rows
                item_name_lower = logical_row.item_name.lower()
                if not any(skip_word in item_name_lower for skip_word in
                           ["費内訳書", "費目", "工種", "種別", "細別", "規格"]):
                    logical_rows.append(logical_row)

            # Move to next row (spanning is handled in _extract_single_logical_row)
            row_index += 1

        return logical_rows

    def _build_hierarchy(self, logical_rows: List[LogicalRow]) -> List[Dict]:
        """Build hierarchical structure from logical rows"""
        hierarchy = []
        stack = []

        for row in logical_rows:
            item = {
                "item_name": row.item_name,
                "unit": row.unit,
                "quantity": row.quantity,
                "unit_price": row.unit_price,
                "amount": row.amount,
                "notes": row.notes,
                "level": row.level,
                "children": []
            }

            # Find the correct parent
            while stack and stack[-1]["level"] >= row.level:
                stack.pop()

            if stack:
                # Add as child of the top item in stack
                stack[-1]["children"].append(item)
            else:
                # Top level item
                hierarchy.append(item)

            stack.append(item)

        return hierarchy

    def get_available_sheets(self, file_path: str) -> List[str]:
        """
        Get available sheet names from Excel file

        Args:
            file_path: Path to the Excel file

        Returns:
            List of sheet names
        """
        try:
            # Read Excel file to get sheet names
            excel_file = pd.ExcelFile(file_path)
            return excel_file.sheet_names
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}")
            return []

    def extract_items_from_buffer_with_sheet(self, buffer, sheet_name: str, project_area: str = "岩手") -> List[Dict]:
        """
        Extract items from Excel buffer for a specific sheet

        Args:
            buffer: Excel file buffer
            sheet_name: Name of the sheet to extract from
            project_area: Project area (岩手 or 北上市)

        Returns:
            List of extracted items
        """
        try:
            # Save buffer to temporary file
            import tempfile
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(buffer)
                tmp_file_path = tmp_file.name

            # Extract hierarchical data
            hierarchical_data = self.extract_hierarchical_data(
                tmp_file_path, sheet_name, project_area)

            # Clean up temporary file
            os.unlink(tmp_file_path)

            if not hierarchical_data:
                return []

            # Convert to flat list of items
            items = []
            for item in hierarchical_data:
                items.extend(self._flatten_hierarchy(item))

            return items

        except Exception as e:
            logger.error(f"Error extracting items from Excel buffer: {str(e)}")
            return []

    def _flatten_hierarchy(self, item: Dict, level: int = 0) -> List[Dict]:
        """Flatten hierarchical data into a list of items"""
        items = []

        # Add current item
        flat_item = {
            "item_key": item.get("item_name", ""),
            "quantity": float(item.get("quantity", 0)) if item.get("quantity") else 0.0,
            "unit": item.get("unit", ""),
            "unit_price": item.get("unit_price", ""),
            "amount": item.get("amount", ""),
            "notes": item.get("notes", ""),
            "level": level,
            "source": "Excel"
        }
        items.append(flat_item)

        # Add children
        for child in item.get("children", []):
            items.extend(self._flatten_hierarchy(child, level + 1))

        return items

    async def save_uploaded_file(self, file, filename: str) -> str:
        """
        Save uploaded file to upload folder

        Args:
            file: Uploaded file object
            filename: Name to save the file as

        Returns:
            Path to saved file
        """
        try:
            file_path = os.path.join(self.upload_folder, filename)
            content = await file.read()
            with open(file_path, 'wb') as f:
                f.write(content)
            return file_path
        except Exception as e:
            logger.error(f"Error saving file: {str(e)}")
            return None
