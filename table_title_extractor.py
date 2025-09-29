"""
Table Title Extractor
A standalone module for extracting table titles from both PDF and Excel documents.
This module provides functions to extract table title information that appears
between reference numbers and column headers.
"""

import re
import logging
from typing import List, Dict, Optional, Tuple
import pandas as pd

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


def extract_pdf_table_title_items(table: List[List[str]], reference_row_idx: int, header_row_idx: int, kitakami_mode: bool = False,
                                  page_text: Optional[str] = None, reference_value: Optional[str] = None) -> Optional[Dict[str, str]]:
    """
    Extract table title items from PDF table that appear between reference number and column headers.

    Args:
        table: The table data as list of lists
        reference_row_idx: Index of the row containing the reference number
        header_row_idx: Index of the row containing column headers

    Returns:
        Dictionary with table title items or None if not found
    """
    try:
        # Check if we have a reference row
        if reference_row_idx >= len(table):
            return None

        reference_row = table[reference_row_idx]

        # Look for table title structure in the reference row itself
        # The table title is embedded within the reference number row
        # Structure can be either 7 cells or 8 cells:
        # 7 cells: [Reference, Item Name, 単位, Unit, 単位数量, Quantity, 単価]
        # 8 cells: [Reference, Item Name, Specification, 単位, Unit, 単位数量, Quantity, 単価]

        if len(reference_row) < 7:
            return None

        # Find positions of "単位" and "単位数量"
        unit_pos = None
        unit_qty_pos = None

        for i, cell in enumerate(reference_row):
            if cell and "単位" in str(cell) and "単位数量" not in str(cell):
                unit_pos = i
            elif cell and "単位数量" in str(cell):
                unit_qty_pos = i

        if (unit_pos is None or unit_qty_pos is None) and kitakami_mode:
            # Kitakami fallback: title is embedded in the SAME logical (spanned) reference row,
            # immediately on the next line within the same cell that contains the reference,
            # and appears BEFORE the column header row.
            try:
                # New: cell-wise scan within the same reference row (grid inside table)
                # Find a title-like cell on the left and a qty+unit cell on the right
                # Helper: strip dimension sequences like B1000×W1000×H1000 or 800×590×2000
                def _strip_dimensions(text: str) -> str:
                    if not text:
                        return text
                    return re.sub(r"\S*×\S*(?:×\S*)+", " ", text)

                # Require: quantity and unit separated by 1+ spaces; allow trailing after 当り/当たり
                # Require an "当り/当たり" marker near the right side to consider qty+unit valid.
                # This prevents picking numbers embedded in the item name area.
                qty_unit_regex = r"(?P<qty>-?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d+)?)[\s　]+(?P<unit>\S+?)[\s　]*(?:当り|当たり)(?:.*)?$"
                # Fallback for adjacent qty+unit like '1ｍ当り' or '10m3当り' (disallow dimensions '×') and keep unit short
                qty_unit_fallback = r"(?P<qty>-?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d+)?)[\s　]*(?P<unit>[^×\s　]{1,6})[\s　]*(?:当り|当たり)(?:.*)?$"
                header_tokens = ["名称", "数 量", "数量",
                                 "単位", "単 価", "金 額", "明細単価番号", "基 準"]

                def is_headerish(text: str) -> bool:
                    return any(tok in text for tok in header_tokens)

                # Collect candidate title (left) and qty/unit (right)
                left_title = None
                right_qty = None
                right_unit = None

                for idx, cell in enumerate(reference_row):
                    if cell is None:
                        continue
                    s = str(cell).strip()
                    if not s:
                        continue
                    # Skip reference tokens themselves
                    if re.search(r"第\s*[0-9０-９]+\s*号", s) or is_headerish(s):
                        continue
                    # Try qty+unit match first
                    try:
                        # Normalize width for matching
                        full_to_half = str.maketrans(
                            '０１２３４５６７８９　', '0123456789 ')
                        norm = s.translate(full_to_half)
                    except Exception:
                        norm = s
                    # Strip dimensions in the cell before matching
                    norm2 = _strip_dimensions(norm)
                    m = re.match(qty_unit_regex, norm2)
                    if not m:
                        m = re.match(qty_unit_fallback, norm2)
                    if m:
                        q = m.group('qty').strip()
                        u_raw = m.group('unit').strip()
                        u_clean = re.sub(r"[\s。、，,.]+$", "", u_raw)
                        u_clean = re.sub(r"(当り|当たり)$", "", u_clean)
                        right_qty = q
                        right_unit = u_clean
                        continue

                    # Otherwise treat as title candidate (prefer the longest meaningful one)
                    if not left_title or len(s) > len(left_title):
                        left_title = s

                if left_title and right_qty and right_unit:
                    # Clean trailing markers from title
                    left_title = re.sub(
                        r"\s*(当り|当たり)?\s*(明細書|単価表)?\s*$", "", left_title)
                    return {"item_name": left_title, "unit": right_unit, "unit_quantity": right_qty}

                # Find the cell within the reference row that contains the reference number
                ref_patterns = [
                    r"第\s*[0-9０-９]+\s*号\s*[一-龯]",  # 第12号明 / 第1号施
                    r"[一-龯々]+\s*[0-9０-９]+\s*号"      # 明12号 / 施1号 など
                ]

                def has_ref(text: str) -> bool:
                    if not text:
                        return False
                    for p in ref_patterns:
                        if re.search(p, str(text)):
                            return True
                    return False

                ref_cell_text = None
                for c in reference_row:
                    if c is None:
                        continue
                    if has_ref(str(c)):
                        ref_cell_text = str(c)
                        break

                if ref_cell_text:
                    # Build candidate lines from within the same spanned reference row
                    ref_lines = ref_cell_text.splitlines()
                    # Locate the line index that holds the reference
                    ref_idx = -1
                    for i, ln in enumerate(ref_lines):
                        if has_ref(ln):
                            ref_idx = i
                            break

                    candidates = []
                    # 1) Lines after the reference inside the same cell
                    for i in range(ref_idx + 1, len(ref_lines)):
                        candidates.append(ref_lines[i])

                    # 2) Lines from other cells in the same row (if any)
                    for c in reference_row:
                        if c is None:
                            continue
                        if str(c) == ref_cell_text:
                            continue
                        for ln in str(c).splitlines():
                            if ln and ln.strip():
                                candidates.append(ln)

                    # 3) Joined line across cells to capture split content
                    candidates.insert(0, "   ".join(
                        [str(c) for c in reference_row if c is not None]))

                    qty_pattern = r"-?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d+)?"

                    for cand in candidates:
                        if not cand or not cand.strip():
                            continue
                        try:
                            full_to_half = str.maketrans(
                                '０１２３４５６７８９　', '0123456789 ')
                            norm_line = str(cand).translate(full_to_half)
                        except Exception:
                            norm_line = str(cand)
                        # Skip header-like content (to ensure it's before headers)
                        if any(h in norm_line for h in ["名称", "数 量", "数量", "単位", "単 価", "金 額", "明細単価番号"]):
                            continue
                        # Strip dimensions in the joined row before matching
                        norm_line2 = _strip_dimensions(norm_line)
                        # Expect: title  (>=2 spaces)  quantity  (>=1 space)  unit
                        m = re.match(
                            rf"^(?P<title>.+?)\s{{2,}}(?P<qty>{qty_pattern})\s+(?P<unit>\S+)[\s　]*(?:当り|当たり)(?:\s+.*)?$", norm_line2)
                        if not m:
                            # Fallback: allow adjacent qty+unit with short unit
                            m = re.match(
                                rf"^(?P<title>.+?)\s+(?P<qty>{qty_pattern})\s*(?P<unit>[^×\s　]{{1,6}})[\s　]*(?:当り|当たり)(?:\s+.*)?$", norm_line2)
                        if m:
                            title = m.group('title').strip()
                            qty = m.group('qty').strip()
                            unit_raw = m.group('unit').strip()
                            # Clean unit: normalize width, strip trailing punctuation
                            unit_clean = re.sub(r"[\s。、，,.]+$", "", unit_raw)
                            canonical_unit = unit_clean
                            # Clean title: drop common trailing words after title
                            title = re.sub(
                                r"\s*(当り|当たり)?\s*(明細書|単価表)?\s*$", "", title)
                            if title and qty and canonical_unit:
                                return {"item_name": title, "unit": canonical_unit, "unit_quantity": qty}
            except Exception:
                pass
            # If not found within the same row, also scan rows between reference and header (some PDFs split the visual line)
            try:
                qty_pattern = r"-?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d+)?"
                for scan_idx in range(reference_row_idx + 1, max(reference_row_idx + 6, header_row_idx)):
                    if scan_idx >= len(table) or scan_idx >= header_row_idx:
                        break
                    scan_row = table[scan_idx]
                    if not scan_row:
                        continue
                    # Build a joined line across cells in the scan row
                    joined = "   ".join([str(c)
                                        for c in scan_row if c is not None])
                    if not joined.strip():
                        continue
                    try:
                        full_to_half = str.maketrans(
                            '０１２３４５６７８９　', '0123456789 ')
                        norm_line = str(joined).translate(full_to_half)
                    except Exception:
                        norm_line = str(joined)
                    if any(h in norm_line for h in ["名称", "数 量", "数量", "単位", "単 価", "金 額", "明細単価番号"]):
                        continue
                    norm_line2 = _strip_dimensions(norm_line)
                    m2 = re.match(
                        rf"^(?P<title>.+?)\s{{2,}}(?P<qty>{qty_pattern})\s+(?P<unit>\S+)[\s　]*(?:当り|当たり)(?:\s+.*)?$", norm_line2)
                    if not m2:
                        m2 = re.match(
                            rf"^(?P<title>.+?)\s+(?P<qty>{qty_pattern})\s*(?P<unit>[^×\s　]{{1,6}})[\s　]*(?:当り|当たり)(?:\s+.*)?$", norm_line2)
                    if m2:
                        title = m2.group('title').strip()
                        qty = m2.group('qty').strip()
                        unit_raw = m2.group('unit').strip()
                        unit_clean = re.sub(r"[\s。、，,.]+$", "", unit_raw)
                        canonical_unit = unit_clean
                        title = re.sub(
                            r"\s*(当り|当たり)?\s*(明細書|単価表)?\s*$", "", title)
                        if title and qty and canonical_unit:
                            return {"item_name": title, "unit": canonical_unit, "unit_quantity": qty}
            except Exception:
                pass
            # Also scan a few rows BEFORE the reference row (e.g., decorative title cell above)
            try:
                qty_pattern = r"-?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d+)?"
                for scan_idx in range(max(0, reference_row_idx - 3), reference_row_idx):
                    scan_row = table[scan_idx]
                    if not scan_row:
                        continue
                    joined = "   ".join([str(c)
                                        for c in scan_row if c is not None])
                    if not joined.strip():
                        continue
                    try:
                        full_to_half = str.maketrans(
                            '０１２３４５６７８９　', '0123456789 ')
                        norm_line = str(joined).translate(full_to_half)
                    except Exception:
                        norm_line = str(joined)
                    if any(h in norm_line for h in ["名称", "数 量", "数量", "単位", "単 価", "金 額", "明細単価番号"]):
                        continue
                    norm_line2 = _strip_dimensions(norm_line)
                    m2 = re.match(
                        rf"^(?P<title>.+?)\s{{2,}}(?P<qty>{qty_pattern})\s+(?P<unit>\S+)[\s　]*(?:当り|当たり)(?:\s+.*)?$", norm_line2)
                    if not m2:
                        m2 = re.match(
                            rf"^(?P<title>.+?)\s+(?P<qty>{qty_pattern})\s*(?P<unit>[^×\s　]{{1,6}})[\s　]*(?:当り|当たり)(?:\s+.*)?$", norm_line2)
                    if m2:
                        title = m2.group('title').strip()
                        qty = m2.group('qty').strip()
                        unit_raw = m2.group('unit').strip()
                        unit_clean = re.sub(r"[\s。、，,.]+$", "", unit_raw)
                        canonical_unit = unit_clean
                        title = re.sub(
                            r"\s*(当り|当たり)?\s*(明細書|単価表)?\s*$", "", title)
                        if title and qty and canonical_unit:
                            return {"item_name": title, "unit": canonical_unit, "unit_quantity": qty}
            except Exception:
                pass
            # Last resort: scan page_text near the reference for pattern '...  <qty> <unit> 当り'
            try:
                if page_text:
                    window = page_text
                    if reference_value:
                        idx = page_text.find(
                            str(reference_value).replace(' ', ''))
                        if idx != -1:
                            window = page_text[max(0, idx-50): idx+200]
                    try:
                        full_to_half = str.maketrans(
                            '０１２３４５６７８９　', '0123456789 ')
                        window = window.translate(full_to_half)
                    except Exception:
                        pass
                    m = None
                    window2 = _strip_dimensions(window)
                    m = re.search(
                        r"(?P<qty>-?\d+(?:,\d{3})*(?:\.\d+)?)[\s　]+(?P<unit>\S+?)[\s　]*(?:当り|当たり)", window2)
                    if not m:
                        m = re.search(
                            r"(?P<qty>-?\d+(?:,\d{3})*(?:\.\d+)?)[\s　]*(?P<unit>[^×\s　]{1,6})[\s　]*(?:当り|当たり)", window2)
                    if m:
                        q = m.group('qty').strip()
                        u = m.group('unit').strip()
                        # Derive a title-like prefix
                        prefix = window[:m.start()].strip()
                        parts = [seg.strip() for seg in re.split(
                            r"[\n\r]", prefix) if seg.strip()]
                        t = parts[-1] if parts else ''
                        t = re.sub(r"\s*(当り|当たり)?\s*(明細書|単価表)?\s*$", "", t)
                        if t:
                            return {"item_name": t, "unit": u, "unit_quantity": q}
            except Exception:
                pass
            return None
        elif (unit_pos is None or unit_qty_pos is None) and not kitakami_mode:
            # Do not attempt fallback for non-Kitakami projects
            return None

        # Extract item name from cells after reference number
        item_name_parts = []
        for i in range(1, unit_pos):  # Start from 1 to skip reference number
            if i < len(reference_row) and reference_row[i]:
                item_name_parts.append(str(reference_row[i]).strip())

        item_name = " ".join(item_name_parts) if item_name_parts else ""

        # Extract unit from cell after "単位"
        unit = ""
        if unit_pos + 1 < len(reference_row):
            unit = str(reference_row[unit_pos + 1]).strip()

        # Extract unit quantity from cell after "単位数量"
        unit_quantity = ""
        if unit_qty_pos + 1 < len(reference_row):
            unit_quantity = str(reference_row[unit_qty_pos + 1]).strip()

        # Validate that we have the required components
        if not unit or not unit_quantity:
            return None

        return {
            "item_name": item_name,
            "unit": unit,
            "unit_quantity": unit_quantity
        }

    except Exception as e:
        logger.error(f"Error extracting PDF table title: {e}")
        return None


def extract_excel_table_title_items(df: pd.DataFrame, reference_row: int, header_row: int) -> Optional[Dict[str, str]]:
    """
    Extract table title items from Excel subtable.

    Args:
        df: DataFrame containing the Excel data
        reference_row: Row index containing the reference number
        header_row: Row index containing the column headers

    Returns:
        Dictionary with item_name, unit, and unit_quantity, or None if no title found
    """
    try:
        # Find table boundaries
        prev_table_end = find_previous_table_end(df, reference_row)
        next_table_end = find_excel_table_end(df, reference_row)

        # Collect sentences from different areas
        sentences_before = []
        sentences_between = []
        sentences_after_table = []

        # Helper function to check if text is meaningful
        def is_meaningful_text(text):
            if not text or pd.isna(text):
                return False
            text = str(text).strip()
            if len(text) < 3:
                return False
            # Check if it contains Japanese characters, numbers, or Latin letters
            return bool(re.search(r'[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF0-9A-Za-z]', text))

        # Helper function to normalize text for comparison
        def normalize_text(text):
            if not text or pd.isna(text):
                return ""
            text = str(text).strip()
            # Convert full-width characters to half-width
            full_to_half = str.maketrans(
                '０１２３４５６７８９ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ'
                'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ　',
                '0123456789abcdefghijklmnopqrstuvwxyz'
                'ABCDEFGHIJKLMNOPQRSTUVWXYZ '
            )
            text = text.translate(full_to_half)
            # Remove spaces for comparison
            return re.sub(r'\s+', '', text)

        # Collect sentences before reference number
        if prev_table_end is not None:
            for row_idx in range(prev_table_end + 1, reference_row):
                row_text = " ".join(
                    [str(cell) for cell in df.iloc[row_idx] if pd.notna(cell) and str(cell).strip()])
                if is_meaningful_text(row_text):
                    sentences_before.append(
                        {'row': row_idx, 'text': row_text.strip()})

        # Collect sentences between reference and table end
        if next_table_end is not None:
            for row_idx in range(reference_row + 1, next_table_end):
                row_text = " ".join(
                    [str(cell) for cell in df.iloc[row_idx] if pd.notna(cell) and str(cell).strip()])
                if is_meaningful_text(row_text):
                    sentences_between.append(
                        {'row': row_idx, 'text': row_text.strip()})

        # Collect sentences after table number
        if next_table_end is not None:
            for row_idx in range(next_table_end + 1, min(next_table_end + 10, len(df))):
                row_text = " ".join(
                    [str(cell) for cell in df.iloc[row_idx] if pd.notna(cell) and str(cell).strip()])
                if is_meaningful_text(row_text):
                    sentences_after_table.append(
                        {'row': row_idx, 'text': row_text.strip()})

        # Select title based on priority order
        selected_title = None

        # Try sentences before reference number first (this is the main area for titles)
        if sentences_before:
            # If multiple sentences, prefer the one closest to the reference number
            if len(sentences_before) >= 2:
                # 2nd from the back
                selected_title = sentences_before[-2]['text']
            else:
                selected_title = sentences_before[0]['text']

        # If no sentence before, try sentences between reference and table end
        if not selected_title and sentences_between:
            selected_title = sentences_between[0]['text']

        # If still no title, try sentences after table number
        if not selected_title and sentences_after_table:
            selected_title = sentences_after_table[0]['text']

        if selected_title:
            return {
                "item_name": selected_title,
                "unit": "",  # Excel titles don't have separate unit fields
                "unit_quantity": ""  # Excel titles don't have separate quantity fields
            }

        return None

    except Exception as e:
        logger.error(f"Error extracting Excel table title: {e}")
        return None


def find_previous_table_end(df: pd.DataFrame, current_reference_row: int) -> int:
    """
    Find the end of the previous table by looking for a table number (just a number, no prefix or suffix)
    before the current reference row.

    Args:
        df: The DataFrame containing the Excel data
        current_reference_row: Current reference row to search backwards from

    Returns:
        Row index where the previous table ends, or 0 if no previous table found
    """
    try:
        # Search backwards from the current reference row
        for row_idx in range(current_reference_row - 1, -1, -1):
            row_data = df.iloc[row_idx]

            # Check if this row contains just a number (table end marker)
            non_empty_cells = [str(cell).strip() for cell in row_data if pd.notna(
                cell) and str(cell).strip()]

            if len(non_empty_cells) == 1:
                cell_value = non_empty_cells[0]
                # Check if it's just a number (no prefix or suffix)
                if re.match(r'^\d+$', cell_value):
                    return row_idx

        # If no previous table end found, return 0 (start of sheet)
        return 0

    except Exception as e:
        logger.error(f"Error finding previous table end: {e}")
        return 0


def find_excel_table_end(df: pd.DataFrame, start_row: int) -> int:
    """
    Find the end of a table in Excel by looking for a table number (just a number, no prefix or suffix).

    Args:
        df: The DataFrame containing the Excel data
        start_row: Starting row to search from

    Returns:
        Row index where the table ends
    """
    try:
        for row_idx in range(start_row, len(df)):
            row_data = df.iloc[row_idx]

            # Check if this row contains just a number (table end marker)
            non_empty_cells = [str(cell).strip() for cell in row_data if pd.notna(
                cell) and str(cell).strip()]

            if len(non_empty_cells) == 1:
                cell_value = non_empty_cells[0]
                # Check if it's just a number (no prefix or suffix)
                if re.match(r'^\d+$', cell_value):
                    return row_idx

        # If no table end found, return the last row
        return len(df) - 1

    except Exception as e:
        logger.error(f"Error finding Excel table end: {e}")
        return len(df) - 1
