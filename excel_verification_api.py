import os
import json
import logging
from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from typing import List, Dict, Optional, Union, Tuple, Any
from dataclasses import dataclass, asdict
import pandas as pd
import tempfile

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Create Router
excel_verification_router = APIRouter()

# Configure upload folder
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


@dataclass
class HierarchicalItem:
    """Represents a hierarchical item with parent-child relationships"""
    item_name: str
    unit: str
    quantity: str
    unit_price: str
    amount: str
    notes: str
    level: int  # Hierarchy level (0 = root, 1 = child, 2 = grandchild, etc.)
    children: List['HierarchicalItem']
    raw_fields: Dict[str, str]
    amount_verification: Optional[Dict[str, Any]] = None


@dataclass
class VerificationResult:
    """Represents verification results"""
    total_items: int
    verified_items: int
    mismatched_items: int
    mismatches: List[Dict[str, Any]]
    business_logic_verified: bool
    extraction_successful: bool
    error_message: Optional[str] = None


class HierarchicalExcelExtractor:
    def __init__(self):
        self.column_patterns = {
            "工事区分・工種・種別・細別": ["費 目 ・ 工 種 ・ 種 別 ・ 細 目", "費目・工種・種別・細別・規格", "工事区分・工種・種別・細別", "工事区分", "工種", "種別", "細別", "費目"],
            "規格": ["規格", "規 格", "名称・規格", "名称", "項目", "品名"],
            "単位": ["単位", "単 位"],
            "数量": ["数量", "数　量"],
            "単価": ["単価", "単　価"],
            "金額": ["金額", "金　額"],
            "数量・金額増減": ["数量・金額増減", "増減", "変更"],
            "摘要": ["摘要", "備考", "摘　要"]
        }

    def extract_hierarchical_data(self, file_path: str, sheet_name: str) -> List[HierarchicalItem]:
        """Extract hierarchical data from Excel sheet with row spanning logic"""
        logger.info(f"Extracting hierarchical data from sheet: {sheet_name}")

        # Read Excel file
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        logger.info(f"Excel sheet shape: {df.shape}")

        # Find header row
        header_row_idx = self._find_header_row(df)
        if header_row_idx is None:
            raise ValueError("Header row not found")

        logger.info(f"Header row found at index: {header_row_idx}")

        # Find column positions
        column_positions = self._find_column_positions(df, header_row_idx)
        logger.info(f"Column positions: {column_positions}")

        # Extract logical rows with row spanning
        logical_rows = self._extract_logical_rows_with_spanning(
            df, header_row_idx, column_positions)
        logger.info(f"Extracted {len(logical_rows)} logical rows")

        # Build hierarchical structure
        hierarchical_items = self._build_hierarchy(logical_rows)
        logger.info(
            f"Built hierarchy with {len(hierarchical_items)} root items")

        # Verify amount calculations
        hierarchical_items = self._verify_amount_calculations(
            hierarchical_items)
        logger.info("Amount verification completed")

        return hierarchical_items

    def _find_header_row(self, df: pd.DataFrame) -> Optional[int]:
        """Find the header row containing column names"""
        for idx, row in df.iterrows():
            row_str = " ".join([str(val) for val in row if pd.notna(val)])
            if "費目" in row_str and "工種" in row_str and "種別" in row_str:
                return idx
        return None

    def _find_next_header_row(self, df: pd.DataFrame, start_row: int) -> Optional[int]:
        """Find the next header row starting from start_row"""
        for idx in range(start_row, len(df)):
            row = df.iloc[idx]
            row_str = " ".join([str(val) for val in row if pd.notna(val)])
            if "費目" in row_str and "工種" in row_str and "種別" in row_str:
                return idx
        return None

    def _is_table_number_row(self, row: pd.Series) -> bool:
        """Check if a row contains just a table number"""
        non_empty_values = [str(val).strip()
                            for val in row if pd.notna(val) and str(val).strip()]
        if len(non_empty_values) == 1:
            value = non_empty_values[0]
            try:
                float(value)
                return True
            except ValueError:
                pass
        return False

    def _find_column_positions(self, df: pd.DataFrame, header_row_idx: int) -> Dict[str, int]:
        """Find column positions for different data types"""
        header_row = df.iloc[header_row_idx]
        positions = {}

        for col_idx, cell_value in enumerate(header_row):
            if pd.isna(cell_value):
                continue

            cell_str = str(cell_value).strip()

            # Map column names to positions
            if "費目" in cell_str or "工種" in cell_str or "種別" in cell_str or "細別" in cell_str or "規格" in cell_str:
                positions['item_name'] = col_idx
            elif "単位" in cell_str:
                positions['unit'] = col_idx
            elif "数量" in cell_str:
                positions['quantity'] = col_idx
            elif "単価" in cell_str:
                positions['unit_price'] = col_idx
            elif "金額" in cell_str:
                positions['amount'] = col_idx
            elif "摘要" in cell_str:
                positions['notes'] = col_idx

        # Fallback positions based on observed structure
        if 'item_name' not in positions:
            positions['item_name'] = 1
        if 'unit' not in positions:
            positions['unit'] = 2
        if 'quantity' not in positions:
            positions['quantity'] = 4
        if 'unit_price' not in positions:
            positions['unit_price'] = 5
        if 'amount' not in positions:
            positions['amount'] = 6
        if 'notes' not in positions:
            positions['notes'] = 7

        return positions

    def _extract_logical_rows_with_spanning(self, df: pd.DataFrame, header_row_idx: int, column_positions: Dict[str, int]) -> List[Dict[str, Any]]:
        """Extract logical rows with row spanning logic across multiple tables"""
        logical_rows = []
        current_row_idx = header_row_idx + 1

        while current_row_idx < len(df):
            if self._is_table_number_row(df.iloc[current_row_idx]):
                logger.info(
                    f"Found table number at row {current_row_idx + 1}, looking for next header")
                next_header_idx = self._find_next_header_row(
                    df, current_row_idx + 1)
                if next_header_idx is not None:
                    current_row_idx = next_header_idx + 1
                    logger.info(
                        f"Found next header at row {next_header_idx + 1}, continuing extraction")
                else:
                    logger.info("No more headers found, ending extraction")
                    break
            else:
                logical_row = self._extract_single_logical_row(
                    df, current_row_idx, column_positions)
                if logical_row:
                    logical_rows.append(logical_row)
                    current_row_idx = logical_row['end_row'] + 1
                else:
                    current_row_idx += 1

        return logical_rows

    def _extract_single_logical_row(self, df: pd.DataFrame, start_row: int, column_positions: Dict[str, int]) -> Optional[Dict[str, Any]]:
        """Extract a single logical row with spanning"""
        if start_row >= len(df):
            return None

        first_row = df.iloc[start_row]

        if self._is_empty_row(first_row):
            return None

        # Extract data from first row
        item_name = self._get_cell_value(
            first_row, column_positions.get('item_name', 1), preserve_spaces=True)
        unit = self._get_cell_value(first_row, column_positions.get('unit', 2))
        quantity = self._get_cell_value(
            first_row, column_positions.get('quantity', 4))
        unit_price = self._get_cell_value(
            first_row, column_positions.get('unit_price', 5))
        amount = self._get_cell_value(
            first_row, column_positions.get('amount', 6))
        notes = self._get_cell_value(
            first_row, column_positions.get('notes', 7))

        # Row spanning logic
        if item_name and start_row + 1 < len(df):
            next_row = df.iloc[start_row + 1]
            next_item_name = self._get_cell_value(
                next_row, column_positions.get('item_name', 1), preserve_spaces=True)
            next_quantity = self._get_cell_value(
                next_row, column_positions.get('quantity', 4))
            next_unit = self._get_cell_value(
                next_row, column_positions.get('unit', 2))
            next_unit_price = self._get_cell_value(
                next_row, column_positions.get('unit_price', 5))
            next_amount = self._get_cell_value(
                next_row, column_positions.get('amount', 6))
            next_notes = self._get_cell_value(
                next_row, column_positions.get('notes', 7))

            should_merge = False

            if (not quantity and not unit and not unit_price and not amount) and (next_quantity or next_unit or next_unit_price or next_amount):
                should_merge = True
            elif (quantity or unit or unit_price or amount) and (next_quantity or next_unit or next_unit_price or next_amount):
                should_merge = True
            elif next_item_name and next_item_name.strip():
                should_merge = True

            if should_merge:
                if next_item_name and next_item_name.strip():
                    combined_item_name = item_name + " " + next_item_name.strip()
                else:
                    combined_item_name = item_name

                quantity = next_quantity if next_quantity else quantity
                unit = next_unit if next_unit else unit
                unit_price = next_unit_price if next_unit_price else unit_price
                amount = next_amount if next_amount else amount
                notes = next_notes if next_notes else notes

                return {
                    'item_name': combined_item_name,
                    'unit': unit,
                    'quantity': quantity,
                    'unit_price': unit_price,
                    'amount': amount,
                    'notes': notes,
                    'start_row': start_row,
                    'end_row': start_row + 1,
                    'raw_fields': {
                        '名称・規格': combined_item_name,
                        '単位': unit,
                        '数量': quantity,
                        '単価': unit_price,
                        '金額': amount,
                        '摘要': notes
                    }
                }

        return {
            'item_name': item_name,
            'unit': unit,
            'quantity': quantity,
            'unit_price': unit_price,
            'amount': amount,
            'notes': notes,
            'start_row': start_row,
            'end_row': start_row,
            'raw_fields': {
                '名称・規格': item_name,
                '単位': unit,
                '数量': quantity,
                '単価': unit_price,
                '金額': amount,
                '摘要': notes
            }
        }

    def _get_cell_value(self, row: pd.Series, col_idx: int, preserve_spaces: bool = False) -> str:
        """Get cell value safely"""
        if col_idx < len(row) and pd.notna(row.iloc[col_idx]):
            value = str(row.iloc[col_idx])
            if preserve_spaces:
                return value
            else:
                return value.strip()
        return ""

    def _is_empty_row(self, row: pd.Series) -> bool:
        """Check if row is empty"""
        return all(pd.isna(val) or str(val).strip() == "" for val in row)

    def _build_hierarchy(self, logical_rows: List[Dict[str, Any]]) -> List[HierarchicalItem]:
        """Build hierarchical structure from logical rows across multiple tables"""
        root_items = []
        stack = []

        for row in logical_rows:
            item_name = row['item_name']
            if not item_name:
                continue

            level = self._get_hierarchy_level(item_name)

            hierarchical_item = HierarchicalItem(
                item_name=item_name,
                unit=row['unit'],
                quantity=row['quantity'],
                unit_price=row['unit_price'],
                amount=row['amount'],
                notes=row['notes'],
                level=level,
                children=[],
                raw_fields=row['raw_fields'],
                amount_verification=None
            )

            parent = self._find_parent_across_tables(stack, level)

            if parent is None:
                if level == 0:
                    stack = [hierarchical_item]
                    root_items.append(hierarchical_item)
                else:
                    root_items.append(hierarchical_item)
                    stack = [hierarchical_item]
            else:
                parent.children.append(hierarchical_item)
                self._update_stack_across_tables(
                    stack, hierarchical_item, level)

        return root_items

    def _get_hierarchy_level(self, item_name: str) -> int:
        """Determine hierarchy level based on indentation"""
        if not item_name:
            return 0

        level = 0
        for char in item_name:
            if char == '　':  # Full-width space (U+3000)
                level += 1
            else:
                break

        return level

    def _find_parent_across_tables(self, stack: List[HierarchicalItem], current_level: int) -> Optional[HierarchicalItem]:
        """Find parent item for current level, maintaining relationships across table boundaries"""
        if not stack:
            return None

        for item in reversed(stack):
            if item.level < current_level:
                return item

        return None

    def _update_stack_across_tables(self, stack: List[HierarchicalItem], new_item: HierarchicalItem, level: int):
        """Update stack to maintain hierarchy across table boundaries"""
        while len(stack) <= level:
            stack.append(None)

        stack[level] = new_item

        if level + 1 < len(stack):
            stack[level + 1:] = []

    def to_json(self, hierarchical_items: List[HierarchicalItem]) -> str:
        """Convert hierarchical items to JSON"""
        def item_to_dict(item: HierarchicalItem) -> Dict[str, Any]:
            result = {
                'item_name': item.item_name,
                'unit': item.unit,
                'quantity': item.quantity,
                'unit_price': item.unit_price,
                'amount': item.amount,
                'notes': item.notes,
                'level': item.level,
                'children': [item_to_dict(child) for child in item.children],
                'raw_fields': item.raw_fields
            }

            if hasattr(item, 'amount_verification') and item.amount_verification is not None:
                result['amount_verification'] = item.amount_verification

            return result

        root_items_dict = [item_to_dict(item) for item in hierarchical_items]
        return json.dumps(root_items_dict, ensure_ascii=False, indent=2)

    def _verify_amount_calculations(self, hierarchical_items: List[HierarchicalItem]) -> List[HierarchicalItem]:
        """Verify that root item amounts match the sum according to business rules"""
        for item in hierarchical_items:
            if item.level == 0:  # Root item
                # Get actual amount from the item
                try:
                    actual_amount = float(item.amount.replace(
                        ',', '')) if item.amount else 0.0
                except (ValueError, AttributeError):
                    actual_amount = 0.0

                # Calculate expected amount based on business rules for specific items
                if item.item_name == "純工事費":
                    expected_amount = self._calculate_junkoji_amount(
                        hierarchical_items)
                elif item.item_name == "工事原価":
                    expected_amount = self._calculate_koji_genka_amount(
                        hierarchical_items)
                elif item.item_name == "工事価格":
                    expected_amount = self._calculate_koji_kakaku_amount(
                        hierarchical_items)
                elif item.item_name == "消費税額及び地方消費税額":
                    # No verification required for tax amount
                    expected_amount = actual_amount
                elif item.item_name == "工事費計":
                    expected_amount = self._calculate_koji_kei_amount(
                        hierarchical_items)
                elif item.item_name == "直接工事費":
                    expected_amount = self._calculate_chokkoji_amount(
                        hierarchical_items)
                elif not item.children:
                    # For items without children (and not business logic items), we still need to verify them
                    # They should match their actual amount (no calculation needed)
                    expected_amount = actual_amount
                    is_matched = True
                    difference = 0.0

                    item.amount_verification = {
                        'is_matched': is_matched,
                        'expected_amount': expected_amount,
                        'actual_amount': actual_amount,
                        'difference': difference
                    }
                    logger.info(
                        f"Verification for '{item.item_name}' (no children): Expected=Actual={actual_amount}")
                    continue
                else:
                    # For other items, use standard children sum calculation
                    expected_amount = self._calculate_children_sum(item)

                # Check if amounts match (with small tolerance for floating point precision)
                tolerance = 0.01  # 1 cent tolerance
                is_matched = abs(expected_amount - actual_amount) <= tolerance

                # Add verification fields to the item
                item.amount_verification = {
                    'is_matched': is_matched,
                    'expected_amount': expected_amount,
                    'actual_amount': actual_amount,
                    'difference': actual_amount - expected_amount
                }

                logger.info(f"Amount verification for '{item.item_name}': "
                            f"Expected: {expected_amount:,.0f}, "
                            f"Actual: {actual_amount:,.0f}, "
                            f"Matched: {is_matched}")
            else:
                # For non-root items, set verification to None
                item.amount_verification = None

        return hierarchical_items

    def _calculate_children_sum(self, item: HierarchicalItem) -> float:
        """Calculate the sum of direct children amounts only (not recursive)"""
        total = 0.0

        for child in item.children:
            # Convert child amount to float, handling empty strings and non-numeric values
            try:
                child_amount = float(child.amount.replace(
                    ',', '')) if child.amount else 0.0
            except (ValueError, AttributeError):
                child_amount = 0.0

            # Add only the direct child's amount (not recursive)
            total += child_amount

        return total

    def _calculate_junkoji_amount(self, hierarchical_items: List[HierarchicalItem]) -> float:
        """Calculate 純工事費 amount: 直接工事費 + sum of Level 0 items before 純工事費 (excluding 直接工事費)"""
        total = 0.0

        # Find 直接工事費 amount
        chokkoji_amount = 0.0
        for item in hierarchical_items:
            if item.item_name == "直接工事費":
                try:
                    chokkoji_amount = float(item.amount.replace(
                        ',', '')) if item.amount else 0.0
                except (ValueError, AttributeError):
                    pass
                break

        # Start with 直接工事費 amount
        total = chokkoji_amount

        # Find the position of 純工事費 to determine which items come before it
        junkoji_index = -1
        for i, item in enumerate(hierarchical_items):
            if item.level == 0 and item.item_name == "純工事費":
                junkoji_index = i
                break

        # If 純工事費 is not found, return just 直接工事費 amount
        if junkoji_index == -1:
            return total

        # Add sum of all Level 0 items that come before 純工事費 (excluding 直接工事費 to avoid double counting)
        for i, item in enumerate(hierarchical_items):
            if item.level == 0 and i < junkoji_index and item.item_name != "直接工事費":
                try:
                    amount = float(item.amount.replace(
                        ',', '')) if item.amount else 0.0
                    total += amount
                except (ValueError, AttributeError):
                    pass

        return total

    def _calculate_koji_genka_amount(self, hierarchical_items: List[HierarchicalItem]) -> float:
        """Calculate 工事原価 amount: 純工事費 amount + 現場管理費 amount"""
        junkoji_amount = 0.0
        genkan_amount = 0.0

        # Find 純工事費 amount
        for item in hierarchical_items:
            if item.item_name == "純工事費":
                try:
                    junkoji_amount = float(item.amount.replace(
                        ',', '')) if item.amount else 0.0
                except (ValueError, AttributeError):
                    pass
                break

        # Find 現場管理費 amount in children of 純工事費
        for item in hierarchical_items:
            if item.item_name == "純工事費":
                for child in item.children:
                    if child.item_name == "　現場管理費":  # Note the space prefix
                        try:
                            genkan_amount = float(child.amount.replace(
                                ',', '')) if child.amount else 0.0
                        except (ValueError, AttributeError):
                            pass
                        break
                break

        return junkoji_amount + genkan_amount

    def _calculate_koji_kakaku_amount(self, hierarchical_items: List[HierarchicalItem]) -> float:
        """Calculate 工事価格 amount: 一般管理費等 + 工事原価 amount"""
        genka_amount = 0.0
        ippankan_amount = 0.0

        # Find 工事原価 amount
        for item in hierarchical_items:
            if item.item_name == "工事原価":
                try:
                    genka_amount = float(item.amount.replace(
                        ',', '')) if item.amount else 0.0
                except (ValueError, AttributeError):
                    pass
                break

        # Find 一般管理費等 amount in children of 工事原価
        for item in hierarchical_items:
            if item.item_name == "工事原価":
                for child in item.children:
                    if child.item_name == "　一般管理費等":  # Note the space prefix
                        try:
                            ippankan_amount = float(child.amount.replace(
                                ',', '')) if child.amount else 0.0
                        except (ValueError, AttributeError):
                            pass
                        break
                break

        return ippankan_amount + genka_amount

    def _calculate_koji_kei_amount(self, hierarchical_items: List[HierarchicalItem]) -> float:
        """Calculate 工事費計 amount: 消費税額及び地方消費税額 + 工事価格 amount"""
        kakaku_amount = 0.0
        tax_amount = 0.0

        # Find 工事価格 amount
        for item in hierarchical_items:
            if item.item_name == "工事価格":
                try:
                    kakaku_amount = float(item.amount.replace(
                        ',', '')) if item.amount else 0.0
                except (ValueError, AttributeError):
                    pass
                break

        # Find 消費税額及び地方消費税額 amount
        for item in hierarchical_items:
            if item.item_name == "消費税額及び地方消費税額":
                try:
                    tax_amount = float(item.amount.replace(
                        ',', '')) if item.amount else 0.0
                except (ValueError, AttributeError):
                    pass
                break

        return tax_amount + kakaku_amount

    def _calculate_chokkoji_amount(self, hierarchical_items: List[HierarchicalItem]) -> float:
        """Calculate 直接工事費 amount: Sum of all Level 0 items that come before it"""
        total = 0.0

        # Find the position of 直接工事費 to determine which items come before it
        chokkoji_index = -1

        # Find the position of 直接工事費
        for i, item in enumerate(hierarchical_items):
            if item.level == 0 and item.item_name == "直接工事費":
                chokkoji_index = i
                break

        # If 直接工事費 is not found, return 0
        if chokkoji_index == -1:
            return 0.0

        # Sum all Level 0 items that come before 直接工事費
        for i, item in enumerate(hierarchical_items):
            if item.level == 0 and i < chokkoji_index:  # Only Level 0 items before 直接工事費
                try:
                    amount = float(item.amount.replace(
                        ',', '')) if item.amount else 0.0
                    total += amount
                except (ValueError, AttributeError):
                    pass

        return total


class ComprehensiveVerifier:
    def __init__(self):
        # Items to exclude from standard verification (they have special business logic)
        self.exclude_items = [
            # No items are excluded - we verify everything
        ]

    def verify_business_logic(self, data: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Verify business logic for special items"""
        results = {
            'business_logic_verified': True,
            'business_logic_mismatches': [],
            'business_logic_details': {}
        }

        # Find key items
        junkoji_fee = None
        koji_genka = None
        koji_kakaku = None
        kojihikei = None
        chokkoji_fee = None
        items_before_junkoji = []
        items_before_chokkoji = []

        for item in data:
            if item['item_name'] == '純工事費':
                junkoji_fee = item
            elif item['item_name'] == '工事原価':
                koji_genka = item
            elif item['item_name'] == '工事価格':
                koji_kakaku = item
            elif item['item_name'] == '工事費計':
                kojihikei = item
            elif item['item_name'] == '直接工事費':
                chokkoji_fee = item
            elif item['level'] == 0 and item['item_name'] not in ['純工事費', '直接工事費']:
                items_before_junkoji.append(item)

        # Verify 純工事費
        if junkoji_fee:
            # Find 直接工事費 amount
            chokkoji_amount = 0.0
            for item in data:
                if item['item_name'] == '直接工事費':
                    try:
                        chokkoji_amount = float(
                            item['amount']) if item['amount'] else 0.0
                    except (ValueError, AttributeError):
                        pass
                    break

            # Find the position of 純工事費 to determine which items come before it
            junkoji_index = -1
            for i, item in enumerate(data):
                if item['level'] == 0 and item['item_name'] == '純工事費':
                    junkoji_index = i
                    break

            # Calculate expected amount: 直接工事費 + sum of Level 0 items before 純工事費 (excluding 直接工事費)
            expected_amount = chokkoji_amount
            if junkoji_index != -1:
                for i, item in enumerate(data):
                    if item['level'] == 0 and i < junkoji_index and item['item_name'] != '直接工事費':
                        try:
                            amount = float(
                                item['amount']) if item['amount'] else 0.0
                            expected_amount += amount
                        except (ValueError, AttributeError):
                            pass

            actual_amount = float(junkoji_fee['amount'])
            tolerance = 0.01
            is_matched = abs(actual_amount - expected_amount) <= tolerance

            results['business_logic_details']['純工事費'] = {
                'expected': expected_amount,
                'actual': actual_amount,
                'matched': is_matched
            }

            if not is_matched:
                results['business_logic_verified'] = False
                results['business_logic_mismatches'].append('純工事費')

        # Verify 直接工事費
        if chokkoji_fee:
            # Find the position of 直接工事費
            chokkoji_index = -1
            for i, item in enumerate(data):
                if item['level'] == 0 and item['item_name'] == '直接工事費':
                    chokkoji_index = i
                    break

            # Calculate sum of Level 0 items before 直接工事費
            expected_amount = 0.0
            if chokkoji_index != -1:
                for i, item in enumerate(data):
                    # Only Level 0 items before 直接工事費
                    if item['level'] == 0 and i < chokkoji_index:
                        try:
                            amount = float(
                                item['amount']) if item['amount'] else 0.0
                            expected_amount += amount
                        except (ValueError, AttributeError):
                            pass

            actual_amount = float(chokkoji_fee['amount'])
            tolerance = 0.01
            is_matched = abs(actual_amount - expected_amount) <= tolerance

            results['business_logic_details']['直接工事費'] = {
                'expected': expected_amount,
                'actual': actual_amount,
                'matched': is_matched
            }

            if not is_matched:
                results['business_logic_verified'] = False
                results['business_logic_mismatches'].append('直接工事費')

        # Verify 工事原価
        if koji_genka and junkoji_fee:
            junkoji_amount = float(junkoji_fee['amount'])
            junkoji_children_sum = sum(
                float(child['amount']) for child in junkoji_fee['children'])
            expected_amount = junkoji_amount + junkoji_children_sum
            actual_amount = float(koji_genka['amount'])
            tolerance = 0.01
            is_matched = abs(actual_amount - expected_amount) <= tolerance

            results['business_logic_details']['工事原価'] = {
                'expected': expected_amount,
                'actual': actual_amount,
                'matched': is_matched
            }

            if not is_matched:
                results['business_logic_verified'] = False
                results['business_logic_mismatches'].append('工事原価')

        return results

    def verify_recursive(self, data: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Verify all items recursively for parent-child amount consistency"""
        mismatches = []
        verified_items = []

        def verify_item_recursively(item, parent_path=""):
            current_path = f"{parent_path}/{item['item_name']}" if parent_path else item['item_name']

            # For items without children, we still verify them
            # They should match their actual amount (no calculation needed)
            if not item['children']:
                # These items pass verification automatically since they have no children to sum
                return True, 0, 0

            # For business logic items, use specific business logic calculations
            business_logic_items = ['純工事費', '工事原価', '工事価格', '工事費計', '直接工事費']
            if item['item_name'] in business_logic_items:
                # These items are already verified by HierarchicalExcelExtractor with business logic
                # We don't need to verify them again here since they have special calculation rules
                return True, 0, 0

            # For standard items, calculate children sum
            children_sum = 0
            verified_children = 0
            mismatched_children = 0

            for child in item['children']:
                child_amount = float(child['amount'])
                children_sum += child_amount

                # Recursively verify child
                child_verified, child_verified_count, child_mismatched_count = verify_item_recursively(
                    child, current_path
                )

                if child_verified:
                    verified_children += 1
                else:
                    mismatched_children += 1

            # Check if parent amount matches children sum
            parent_amount = float(item['amount'])
            difference = parent_amount - children_sum
            tolerance = 0.01
            is_matched = abs(difference) <= tolerance

            if is_matched:
                verified_items.append({
                    'path': current_path,
                    'level': item['level'],
                    'amount': parent_amount,
                    'children_sum': children_sum,
                    'difference': difference
                })
            else:
                mismatches.append({
                    'path': current_path,
                    'level': item['level'],
                    'amount': parent_amount,
                    'children_sum': children_sum,
                    'difference': difference,
                    'item_name': item['item_name']
                })

            return is_matched, verified_children, mismatched_children

        # Start recursive verification from root items
        for root_item in data:
            verify_item_recursively(root_item)

        return {
            'total_items_verified': len(verified_items),
            'total_items_mismatched': len(mismatches),
            'mismatches': mismatches,
            'verified_items': verified_items
        }


def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def verify_excel_file(file_path: str, sheet_name: str) -> VerificationResult:
    """
    Main API function to verify Excel file

    Args:
        file_path (str): Path to the Excel file
        sheet_name (str): Name of the sheet to verify

    Returns:
        VerificationResult: Complete verification results
    """
    try:
        # Check if file exists
        if not os.path.exists(file_path):
            return VerificationResult(
                total_items=0,
                verified_items=0,
                mismatched_items=0,
                mismatches=[],
                business_logic_verified=False,
                extraction_successful=False,
                error_message=f"File not found: {file_path}"
            )

        # Extract hierarchical data
        extractor = HierarchicalExcelExtractor()
        hierarchical_items = extractor.extract_hierarchical_data(
            file_path, sheet_name)

        # Convert to JSON format for verification
        json_data = json.loads(extractor.to_json(hierarchical_items))

        # Perform verifications
        verifier = ComprehensiveVerifier()

        # Business logic verification
        business_logic_results = verifier.verify_business_logic(json_data)

        # Recursive verification
        recursive_results = verifier.verify_recursive(json_data)

        # Count all root items (including business logic items)
        total_items = len([item for item in json_data if item['level'] == 0])

        # Also check business logic verification results from HierarchicalExcelExtractor
        business_logic_mismatches = []
        for item in json_data:
            if item.get('amount_verification') and not item['amount_verification']['is_matched']:
                business_logic_mismatches.append({
                    'path': item['item_name'],
                    'level': item['level'],
                    'amount': item['amount_verification']['actual_amount'],
                    'children_sum': item['amount_verification']['expected_amount'],
                    'difference': item['amount_verification']['difference'],
                    'item_name': item['item_name']
                })

        # Combine mismatches from both verification methods and deduplicate
        all_mismatches = recursive_results['mismatches'] + \
            business_logic_mismatches

        # Deduplicate mismatches based on path and level
        seen_paths = set()
        unique_mismatches = []

        for mismatch in all_mismatches:
            path_key = f"{mismatch['path']}_{mismatch['level']}"
            if path_key not in seen_paths:
                seen_paths.add(path_key)
                unique_mismatches.append(mismatch)

        total_mismatched_items = len(unique_mismatches)

        # Calculate verified items: total - mismatched
        total_verified_items = total_items - total_mismatched_items

        return VerificationResult(
            total_items=total_items,
            verified_items=total_verified_items,
            mismatched_items=total_mismatched_items,
            mismatches=unique_mismatches,
            business_logic_verified=business_logic_results['business_logic_verified'],
            extraction_successful=True,
            error_message=None
        )

    except Exception as e:
        logger.error(f"Error during verification: {e}")
        return VerificationResult(
            total_items=0,
            verified_items=0,
            mismatched_items=0,
            mismatches=[],
            business_logic_verified=False,
            extraction_successful=False,
            error_message=str(e)
        )


@excel_verification_router.post("/verify-excel")
async def verify_excel(file: UploadFile = File(...), sheet_name: str = Form(...)):
    """API endpoint to verify Excel file"""
    try:
        # Validate file
        if not file.filename:
            raise HTTPException(status_code=400, detail="No file selected")

        if not allowed_file(file.filename):
            raise HTTPException(
                status_code=400, detail="Invalid file type. Only Excel files (.xlsx, .xls) are allowed")

        if not sheet_name:
            raise HTTPException(
                status_code=400, detail="Sheet name is required")

        # Save file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
            content = await file.read()
            temp_file.write(content)
            temp_file_path = temp_file.name

        try:
            # Verify Excel file
            result = verify_excel_file(temp_file_path, sheet_name)

            # Convert result to dict for JSON response
            result_dict = asdict(result)

            return {
                'success': True,
                'result': result_dict
            }
        finally:
            # Clean up temporary file
            try:
                os.unlink(temp_file_path)
            except:
                pass

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error in verify_excel endpoint: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@excel_verification_router.post("/get-sheets")
async def get_sheets(file: UploadFile = File(...)):
    """API endpoint to get available sheets from Excel file"""
    try:
        # Validate file
        if not file.filename:
            raise HTTPException(status_code=400, detail="No file selected")

        if not allowed_file(file.filename):
            raise HTTPException(
                status_code=400, detail="Invalid file type. Only Excel files (.xlsx, .xls) are allowed")

        # Save file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
            content = await file.read()
            temp_file.write(content)
            temp_file_path = temp_file.name

        try:
            # Get sheet names
            excel_file = pd.ExcelFile(temp_file_path)
            sheet_names = excel_file.sheet_names

            return {
                'success': True,
                'sheets': sheet_names
            }
        except Exception as e:
            raise HTTPException(
                status_code=400, detail=f'Error reading Excel file: {str(e)}')
        finally:
            # Clean up temporary file
            try:
                os.unlink(temp_file_path)
            except:
                pass

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error in get_sheets endpoint: {e}")
        raise HTTPException(status_code=500, detail=str(e))
