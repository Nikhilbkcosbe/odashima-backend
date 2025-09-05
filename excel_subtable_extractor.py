"""
Excel Subtable Extractor
A standalone function for extracting subtables from Excel files based on specific patterns.
Designed to be used as an API service that receives excel file name and sheet name.
"""

import pandas as pd
import re
from typing import List, Dict, Tuple, Optional
import logging
from table_title_extractor import extract_excel_table_title_items, find_excel_table_end

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def normalize_text(text: str) -> str:
    """
    Normalize text by removing spaces and converting full-width characters to half-width
    """
    if not text or pd.isna(text):
        return ""

    # Convert to string if not already
    text = str(text).strip()

    # Convert full-width characters to half-width
    full_to_half = str.maketrans(
        '０１２３４５６７８９ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ'
        'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ　',
        '0123456789abcdefghijklmnopqrstuvwxyz'
        'ABCDEFGHIJKLMNOPQRSTUVWXYZ '
    )
    text = text.translate(full_to_half)

    # Remove all spaces for comparison
    text = re.sub(r'\s+', '', text)

    return text


def find_reference_number_pattern(text: str) -> bool:
    """
    Check if text matches the reference number pattern: kanji + Number + 号
    Examples: 内1号, 内2号, etc.
    """
    if not text:
        return False

    normalized = normalize_text(text)
    # Pattern: one or more kanji characters followed by number(s) followed by 号
    pattern = r'[\u4e00-\u9faf]+\d+号'
    return bool(re.search(pattern, normalized))


def find_column_headers_and_positions(df: pd.DataFrame, start_row: int) -> Tuple[Optional[int], Dict[str, int]]:
    """
    Find the column header row and return exact column positions
    Based on observed Excel structure:
    - Column 2: 名称／規格 (Item Name/Specification)  
    - Column 3: 単位 (Unit)
    - Column 5: 数量 (Quantity)
    - Column 6: 単価 (Unit Price)
    - Column 7: 金額 (Amount)
    - Column 8: 摘要 (Notes)
    """
    target_columns = {
        '名称': ['名称', '名称／規格', '名称/規格', '規格'],
        '単位': ['単位'],
        '数量': ['数量'],
        '単価': ['単価'],
        '摘要': ['摘要']
    }

    # Skip header detection for now and use fixed positions
    # The header detection is not working reliably due to merged cells or formatting issues
    # Let's find the header row but use fixed positions

    # Use fixed positions based on actual observed Excel structure
    # From debug: units are in col 4, not col 3!
    fixed_positions = {
        '名称': 2,    # Column 2: Specific item name/specification
        '単位': 4,    # Column 4: Unit (本, etc.) - CORRECTED!
        '数量': 5,    # Column 5: Quantity
        '単価': 6,    # Column 6: Unit price
        '金額': 7,    # Column 7: Amount
        '摘要': 8     # Column 8: Notes
    }
    logger.info(f"Using fixed column positions (corrected): {fixed_positions}")
    return start_row + 1, fixed_positions  # Skip one row for header


def extract_subtable_data(df: pd.DataFrame, header_row: int, column_positions: Dict[str, int], reference_number: str) -> List[Dict[str, str]]:
    """
    Extract data rows from subtable until reaching '計' marker or next reference number
    """
    data_rows = []
    current_row = header_row + 1

    # Use the provided column positions
    general_item_col = 1  # Column 1: General item category (排水管, etc.)
    item_name_col = column_positions.get(
        '名称', 2)  # Column 2: Specific item name
    unit_col = column_positions.get('単位', 4)  # Column 4: Unit (corrected)
    quantity_col = column_positions.get('数量', 5)
    unit_price_col = column_positions.get('単価', 6)
    amount_col = column_positions.get('金額', 7)
    notes_col = column_positions.get('摘要', 8)

    # Helper to detect trailing table-number-only row which marks the end of a subtable
    def _is_table_number_row(series_row: pd.Series) -> bool:
        try:
            values = [str(v).strip() for v in series_row.tolist()]
            non_empty = [v for v in values if v and v.lower() != 'nan']
            if len(non_empty) != 1:
                return False
            return non_empty[0].isdigit()
        except Exception:
            return False

    while current_row < len(df):
        row_data = df.iloc[current_row].fillna('')

        # End-of-table: trailing row that contains only a single numeric table number
        if _is_table_number_row(row_data):
            logger.debug(
                f"Found trailing table number row at {current_row}; ending subtable '{reference_number}'")
            break

        # Check if we've reached the end marker '計'
        row_text = ' '.join([str(cell)
                            for cell in row_data if str(cell).strip()])
        if '計' in row_text:
            logger.debug(f"Found end marker '計' at row {current_row}")
            break

        # Check if we've reached another reference number (only in typical header positions)
        for col_idx, cell_value in enumerate(row_data):
            if col_idx <= 3 and find_reference_number_pattern(str(cell_value)):
                logger.debug(
                    f"Found next reference number at row {current_row}, stopping extraction")
                return data_rows

        # Extract item names from both general category (col 1) and specific item (col 2)
        general_item = str(row_data.iloc[general_item_col]).strip(
        ) if general_item_col < len(row_data) else ""
        specific_item = str(row_data.iloc[item_name_col]).strip(
        ) if item_name_col < len(row_data) else ""

        # Clean up 'nan' values
        general_item = general_item if general_item != 'nan' else ""
        specific_item = specific_item if specific_item != 'nan' else ""

        # Extract data from specific columns
        unit = str(row_data.iloc[unit_col]).strip(
        ) if unit_col < len(row_data) else ""
        quantity = str(row_data.iloc[quantity_col]).strip(
        ) if quantity_col < len(row_data) else ""
        unit_price = str(row_data.iloc[unit_price_col]).strip(
        ) if unit_price_col < len(row_data) else ""
        amount = str(row_data.iloc[amount_col]).strip(
        ) if amount_col < len(row_data) else ""
        notes = str(row_data.iloc[notes_col]).strip(
        ) if notes_col < len(row_data) else ""

        # Clean up 'nan' values
        def clean_value(val):
            return val if val and val != 'nan' else ""

        unit = clean_value(unit)
        quantity = clean_value(quantity)
        unit_price = clean_value(unit_price)
        amount = clean_value(amount)
        notes = clean_value(notes)

        # Row spanning logic: Check if this row has only general item and next row has specific data
        if (general_item and not specific_item and not quantity and not unit and not amount and current_row + 1 < len(df)):
            logger.debug(
                f"Row spanning triggered for '{reference_number}' at row {current_row}: general_item='{general_item}'")
            next_row_data = df.iloc[current_row + 1].fillna('')
            next_specific_item = str(next_row_data.iloc[item_name_col]).strip(
            ) if item_name_col < len(next_row_data) else ""
            next_unit = str(next_row_data.iloc[unit_col]).strip(
            ) if unit_col < len(next_row_data) else ""
            next_quantity = str(next_row_data.iloc[quantity_col]).strip(
            ) if quantity_col < len(next_row_data) else ""
            next_unit_price = str(next_row_data.iloc[unit_price_col]).strip(
            ) if unit_price_col < len(next_row_data) else ""
            next_amount = str(next_row_data.iloc[amount_col]).strip(
            ) if amount_col < len(next_row_data) else ""

            # Clean up next row values
            next_specific_item = clean_value(next_specific_item)
            next_unit = clean_value(next_unit)
            next_quantity = clean_value(next_quantity)
            next_unit_price = clean_value(next_unit_price)
            next_amount = clean_value(next_amount)

            # Merge if next row has data (with or without specific item name)
            if next_specific_item or next_unit or next_quantity or next_unit_price or next_amount:
                # Combine general and specific item names (if specific item exists)
                if next_specific_item:
                    item_name = f"{general_item} {next_specific_item}".strip()
                else:
                    item_name = general_item  # Use only general item if no specific item
                unit = next_unit
                quantity = next_quantity
                unit_price = next_unit_price
                amount = next_amount
                current_row += 1  # Skip the merged row
                logger.debug(
                    f"Merged rows {current_row-1} and {current_row}: '{general_item}' + '{next_specific_item}'")
            else:
                # Just use the general item name if no mergeable next row
                item_name = general_item
        else:
            # Use specific item if available, otherwise general item
            item_name = specific_item or general_item

        # Filter out header rows and only add rows with meaningful data
        is_header_row = any([
            normalize_text(item_name) in ['名称', '名称／規格', '名称/規格', '規格'],
            normalize_text(unit) == '単位',
            normalize_text(quantity) == '数量',
            normalize_text(unit_price) == '単価',
            # Handle full-width space
            normalize_text(amount) in ['金額', '金\u3000額'],
            '規　格' in item_name,
            '金　額' in str(amount)
        ])

        # Filter out header rows and only add rows with meaningful data
        is_header_row = any([
            normalize_text(item_name) in ['名称', '名称／規格', '名称/規格', '規格'],
            normalize_text(unit) == '単位',
            normalize_text(quantity) == '数量',
            normalize_text(unit_price) == '単価',
            # Handle full-width space
            normalize_text(amount) in ['金額', '金\u3000額'],
            '規　格' in item_name,
            '金　額' in str(amount)
        ])

        if not is_header_row and (item_name or quantity or unit_price or amount):
            extracted_row = {
                'reference_number': reference_number,
                '名称': item_name,
                '単位': unit,
                '数量': quantity,
                '単価': unit_price,
                '金額': amount,
                '摘要': notes
            }
            data_rows.append(extracted_row)
            logger.debug(f"Added data row: {item_name}")

        current_row += 1

    return data_rows


def extract_subtables_from_excel_sheet(excel_file_path: str, sheet_name: str) -> List[Dict]:
    """
    Main API function to extract subtables from a specific Excel sheet

    Args:
        excel_file_path (str): Path to the Excel file
        sheet_name (str): Name of the sheet to process

    Returns:
        List[Dict]: List of extracted subtables with their data
    """
    try:
        # Validate inputs
        if not excel_file_path or not sheet_name:
            raise ValueError(
                "Both excel_file_path and sheet_name are required")

        # Read the Excel sheet
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=None)
        logger.info(
            f"Successfully loaded sheet '{sheet_name}' with {len(df)} rows and {len(df.columns)} columns")

        subtables = []
        reference_counts: Dict[str, int] = {}
        current_row = 0

        while current_row < len(df):
            # Search for reference number pattern
            row_data = df.iloc[current_row].fillna('')

            for col_idx, cell_value in enumerate(row_data):
                # Only look for reference numbers in typical header positions (columns 0-3)
                # This avoids false positives from reference codes in remarks columns
                if col_idx <= 3 and find_reference_number_pattern(str(cell_value)):
                    reference_number = str(cell_value).strip()
                    print(
                        f"DEBUG: Found reference number '{reference_number}' at row {current_row}, col {col_idx}")
                    logger.info(
                        f"Found reference number '{reference_number}' at row {current_row}, col {col_idx}")

                    # Find column headers
                    print(
                        f"DEBUG: About to call find_column_headers_and_positions for {reference_number}")
                    header_row, column_positions = find_column_headers_and_positions(
                        df, current_row + 1)
                    print(
                        f"DEBUG: Header row result: {header_row}, column_positions: {column_positions}")
                    logger.debug(
                        f"Header row result: {header_row}, column_positions: {column_positions}")

                    if header_row is not None:
                        # Extract table title
                        print(
                            f"DEBUG: About to extract title for {reference_number} at row {current_row}")
                        logger.debug(
                            f"Attempting to extract title for {reference_number} at row {current_row}")
                        logger.debug(f"Header row: {header_row}")
                        table_title = extract_excel_table_title_items(
                            df, current_row, header_row)
                        print(
                            f"DEBUG: Title extraction result for {reference_number}: {table_title}")
                        logger.debug(
                            f"Title extraction result for {reference_number}: {table_title}")

                        # Create unique reference number suffix (-2, -3, ...) when the same reference appears again
                        base_ref = reference_number
                        repeat_count = reference_counts.get(base_ref, 0)
                        unique_ref = f"{base_ref}-{repeat_count+1}" if repeat_count >= 1 else base_ref

                        # Extract data rows using unique reference
                        data_rows = extract_subtable_data(
                            df, header_row, column_positions, unique_ref)

                        if data_rows:
                            subtable = {
                                'reference_number': unique_ref,
                                'sheet_name': sheet_name,
                                'start_row': current_row + 1,  # 1-indexed for Excel compatibility
                                'header_row': header_row + 1,  # 1-indexed for Excel compatibility
                                'column_positions': column_positions,
                                'data_rows': data_rows,
                                'total_rows': len(data_rows)
                            }

                            # Add table title if found
                            if table_title:
                                subtable['table_title'] = table_title
                                logger.info(
                                    f"Extracted table title for {reference_number}: {table_title}")

                            subtables.append(subtable)
                            # Update reference occurrence count
                            reference_counts[base_ref] = reference_counts.get(
                                base_ref, 0) + 1
                            logger.info(
                                f"Extracted subtable '{reference_number}' with {len(data_rows)} data rows")
                        else:
                            logger.warning(
                                f"No data rows found for subtable '{reference_number}' - skipping")

                        # Move past this subtable to look for the next one
                        current_row = header_row + len(data_rows) + 3
                        break
                    else:
                        print(
                            f"DEBUG: Header row is None for {reference_number}")
                        logger.warning(
                            f"Header row is None for {reference_number} - skipping")
                        break
            else:
                current_row += 1

        logger.info(
            f"Total subtables extracted from sheet '{sheet_name}': {len(subtables)}")
        return subtables

    except Exception as e:
        logger.error(
            f"Error extracting subtables from sheet '{sheet_name}': {str(e)}")
        raise


def extract_subtables_from_excel(excel_file_path: str, sheet_name: str = None) -> List[Dict]:
    """
    API function to extract subtables from Excel file

    Args:
        excel_file_path (str): Path to the Excel file
        sheet_name (str, optional): Specific sheet name to process. 
                                  If None, processes all sheets except the first one.

    Returns:
        List[Dict]: List of all extracted subtables
    """
    try:
        # Get all sheet names
        xl_file = pd.ExcelFile(excel_file_path)
        all_sheets = xl_file.sheet_names
        logger.info(f"Available sheets in Excel file: {all_sheets}")

        if sheet_name:
            # Process specific sheet
            if sheet_name not in all_sheets:
                raise ValueError(
                    f"Sheet '{sheet_name}' not found in Excel file. Available sheets: {all_sheets}")
            sheets_to_process = [sheet_name]
        else:
            # Process all sheets except the first one (main sheet)
            sheets_to_process = all_sheets[1:]

        all_subtables = []

        for sheet in sheets_to_process:
            logger.info(f"Processing sheet: {sheet}")
            sheet_subtables = extract_subtables_from_excel_sheet(
                excel_file_path, sheet)
            all_subtables.extend(sheet_subtables)

        return all_subtables

    except Exception as e:
        logger.error(
            f"Error processing Excel file '{excel_file_path}': {str(e)}")
        raise


# Example usage and testing
if __name__ == "__main__":
    # Test with the provided Excel file
    excel_file = "【修正】水沢橋　積算書.xlsx"

    try:
        # Extract from second sheet only
        xl_file = pd.ExcelFile(excel_file)
        if len(xl_file.sheet_names) > 1:
            second_sheet = xl_file.sheet_names[1]
            print(f"Testing extraction from sheet: {second_sheet}")

            subtables = extract_subtables_from_excel(excel_file, second_sheet)

            print(f"\n=== EXTRACTION RESULTS ===")
            print(f"Total subtables found: {len(subtables)}")

            for i, subtable in enumerate(subtables, 1):
                print(f"\nSubtable {i}: {subtable['reference_number']}")
                print(f"  Total rows: {subtable['total_rows']}")

                # Show first 2 data rows as example
                for j, row in enumerate(subtable['data_rows'][:2]):
                    print(f"  Row {j+1}: {row}")

        else:
            print("Excel file has only one sheet")

    except Exception as e:
        print(f"Error during testing: {e}")
