from subtable_pdf_extractor import SubtablePDFExtractor
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

        # Nousei-specific column patterns (農政)
        self.nousei_column_patterns = {
            "工種・種目": [
                "工種・種目", "工種･種目", "工 種 ・ 種 目", "工 種 ･ 種 目", "工種種目", "工種  種目"
            ],
            "規格": ["規格", "規 格", "名称・規格", "名称", "項目", "品名"],
            "単位": ["単位", "単 位", "単　位"],
            "数量": ["数量", "数 量", "数　量"],
            "備考": ["備考", "摘 要", "摘要", "摘　要"]
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

                # Initialize Nousei global header mapping for this extraction session
                self._nousei_global_cols = None

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
            # 農政ではヘッダ検出に失敗しても、ドット行検出で継続する
            if project_area != "農政":
                return items
            header_row = table[0] if table else []
            header_idx = 0

        # Determine effective project area from header if possible
        effective_area = self._detect_project_area_from_header(
            header_row) or project_area

        # Build column mapping (attempt with effective area first, then fallback to other patterns)
        if effective_area == "農政" and getattr(self, "_nousei_global_cols", None):
            col_indices = self._nousei_global_cols
        else:
            col_indices = self._get_column_mapping(header_row, effective_area)
        # 農政: ヘッダの行が分割されている場合があるため、最初の数行をスキャンして
        # 『単位』と『数量』が同時に現れる行をヘッダとして再評価
        if effective_area == "農政" and not col_indices:
            try:
                import re as _re

                def _clean(s: str) -> str:
                    return _re.sub(r"[\s\u3000・･]+", "", s or "")
                scan_limit = min(6, len(table))
                best_idx = -1
                for i in range(0, scan_limit):
                    row = table[i]
                    if not row:
                        continue
                    texts = [_clean(str(c)) for c in row if c]
                    if any("単位" in t for t in texts) and any("数量" in t for t in texts):
                        best_idx = i
                        break
                if best_idx != -1:
                    header_row = table[best_idx]
                    header_idx = best_idx
                    col_indices = self._get_column_mapping(header_row, "農政")
            except Exception:
                pass
        if not col_indices:
            # 農政: 最小限のマッピングで続行（数量/単位が検出できない場合でも、名称列のみで進める）
            if effective_area == "農政":
                logger.info(
                    "Nousei: Using minimal header mapping %s due to missing 数量/単位.", "{'工種・種目': 0}")
                col_indices = {"工種・種目": 0}
            else:
                # Fallback: try the opposite area's patterns just in case
                if effective_area == "岩手":
                    fallback_area = "北上市"
                elif effective_area == "北上市":
                    fallback_area = "岩手"
                else:  # 農政など
                    fallback_area = "岩手"
                col_indices = self._get_column_mapping(
                    header_row, fallback_area)
                if not col_indices:
                    return items
                effective_area = fallback_area

        # Cache the first discovered mapping globally for Nousei
        # ただし最小マッピング（数量/単位なし）はキャッシュしない
        if (
            effective_area == "農政"
            and col_indices
            and not getattr(self, "_nousei_global_cols", None)
            and ("数量" in col_indices and "単位" in col_indices)
        ):
            self._nousei_global_cols = col_indices

        data_rows = table[header_idx + 1:]
        # For 農政, some dotted main rows may precede the detected header.
        # Process the entire table and rely on dotted markers to identify items.
        # exclude rows starting with '＊' or '*'; merge unit/quantity from up to next 2 non-item rows.
        if effective_area == "農政":
            data_rows = table
            # Reuse last mapping when available
            if not hasattr(self, "_nousei_last_cols"):
                self._nousei_last_cols = None
            if not col_indices and self._nousei_last_cols:
                col_indices = self._nousei_last_cols
            elif col_indices:
                self._nousei_last_cols = col_indices

            # Helper: trim all leading whitespace (Unicode aware)
            def _lstrip_all_ws(text: str) -> str:
                try:
                    return re.sub(r"^\s+", "", text or "")
                except Exception:
                    return (text or "").lstrip()

            def is_main_marker(text: str, in_subtable: bool) -> bool:
                if not text:
                    return False
                s_raw = str(text)
                if not s_raw:
                    return False
                # Trim ASCII and full-width leading spaces before checking
                s = _lstrip_all_ws(s_raw)
                if not s:
                    return False
                # Exclude explicit note rows
                if s.startswith('＊') or s.startswith('*'):
                    return False
                # Main items must start with at least one ideographic middle dot
                if s.startswith('・') or s.startswith('･'):
                    count = 0
                    for ch in s:
                        if ch == '・' or ch == '･':
                            count += 1
                        else:
                            break
                    return count in (1, 2, 3)
                # Zero-dot rows are never main items in Nousei
                return False

            def _get_marker_text(row: List) -> str:
                """Return a best-effort dotted marker cell text for this row.
                Prefer the mapped name column; otherwise scan the first few cells.
                Always returns a string trimmed of leading spaces (ASCII/full-width)."""
                try:
                    name_idx = col_indices.get('工種・種目', -1)
                    if name_idx != -1 and name_idx < len(row) and row[name_idx]:
                        cand = _lstrip_all_ws(str(row[name_idx]))
                        if cand.startswith('・') or cand.startswith('･'):
                            return cand
                    # Scan all cells for the first dotted start
                    for i, cell in enumerate(row):
                        if cell:
                            cand = _lstrip_all_ws(str(cell))
                            if cand.startswith('・') or cand.startswith('･'):
                                return cand
                    # Fallback: first cell trimmed
                    if row and row[0]:
                        return _lstrip_all_ws(str(row[0]))
                except Exception:
                    pass
                return ''

            in_subtable_after_triple = False
            for row_idx, row in enumerate(data_rows):
                try:
                    marker_text = _get_marker_text(row)
                    first = str(row[0]) if row and row[0] is not None else ""
                    # reset subtable region when a new dotted section begins
                    if marker_text.startswith('・') or marker_text.startswith('･'):
                        in_subtable_after_triple = False
                    if not is_main_marker(marker_text or first, in_subtable_after_triple):
                        continue

                    raw_fields: Dict[str, str] = {}
                    # name/spec/unit/qty from current row

                    def get_cell(col_name: str) -> str:
                        idx = col_indices.get(col_name, -1)
                        return (str(row[idx]).strip() if idx != -1 and idx < len(row) and row[idx] else "")

                    name = marker_text or get_cell("工種・種目") or (
                        str(row[0]).strip() if row and row[0] else "")
                    # Expose item name without leading dotted markers used only for classification
                    try:
                        display_name = str(name).lstrip(' \t\u3000')
                        display_name = re.sub(r"^[・･]+\s*", "", display_name)
                    except Exception:
                        display_name = name
                    raw_fields["工種・種目"] = display_name
                    # Also capture 備考 if available (metadata only; no logic change)
                    if "備考" in col_indices and not raw_fields.get("備考"):
                        try:
                            raw_fields["備考"] = get_cell("備考")
                        except Exception:
                            pass
                    if "規格" in col_indices:
                        raw_fields["規格"] = get_cell("規格")
                    if "単位" in col_indices:
                        raw_fields["単位"] = get_cell("単位")
                    if "数量" in col_indices:
                        raw_fields["数量"] = get_cell("数量")

                    # 規格補完は行わない（特に単位=「式」を規格へ昇格しない）

                    # 農政: 規格/単位/数量の列が見つからない、または空の場合はヒューリスティックで補完
                    # 右側のセルから、数量/単位でない最初の非空セルを規格候補とする
                    try:
                        def looks_like_unit(text: str) -> bool:
                            t = (text or "").strip()
                            return t in [
                                "m3", "m2", "m", "㎥", "㎡", "ｍ", "mm", "㎜", "cm", "㎝",
                                "枚", "箇所", "kg", "本", "人", "日", "ha", "掛㎡", "孔", "ton", "式", "基", "台",
                                "Ｌ", "L", "台･日", "台・日"
                            ]

                        def looks_like_quantity(text: str) -> bool:
                            import re as _re2
                            t = (text or "").replace(
                                ',', '').replace('，', '').strip()
                            return bool(_re2.match(r"^\d+(?:\.\d+)?$", t))

                        # index of the dotted name cell
                        dotted_idx = None
                        for ci, cell in enumerate(row):
                            if not cell:
                                continue
                            cand = _lstrip_all_ws(str(cell))
                            if cand.startswith('・') or cand.startswith('･'):
                                dotted_idx = ci
                                break

                        # 補完: 規格
                        if not raw_fields.get("規格") and dotted_idx is not None:
                            for ci in range(dotted_idx + 1, len(row)):
                                cell = row[ci]
                                if not cell:
                                    continue
                                t = str(cell).strip()
                                if not t:
                                    continue
                                if looks_like_unit(t) or looks_like_quantity(t) or t.startswith("算出数量"):
                                    continue
                                raw_fields["規格"] = t
                                break

                        # 補完: 単位
                        if ("単位" not in col_indices or not raw_fields.get("単位")) and dotted_idx is not None:
                            for ci in range(dotted_idx + 1, len(row)):
                                cell = row[ci]
                                if not cell:
                                    continue
                                t = str(cell).strip()
                                if looks_like_unit(t):
                                    raw_fields["単位"] = t
                                    break

                        # 補完: 数量
                        if ("数量" not in col_indices or not raw_fields.get("数量")) and dotted_idx is not None:
                            for ci in range(dotted_idx + 1, len(row)):
                                cell = row[ci]
                                if not cell:
                                    continue
                                t = str(cell).replace(
                                    ',', '').replace('，', '').strip()
                                if looks_like_quantity(t):
                                    raw_fields["数量"] = t
                                    break
                    except Exception:
                        pass

                    # No lookahead merge: zero-dot rows are subtable/detail and must be ignored
                    is_triple = (marker_text.startswith('・・・') or marker_text.startswith(
                        '･･･')) if marker_text else first.startswith('・・・')

                    # parse quantity (blank allowed)
                    qty_val = 0.0
                    qtext = raw_fields.get("数量") or ""
                    if qtext:
                        try:
                            qty_val = float(qtext.replace(
                                ',', '').replace('，', ''))
                        except Exception:
                            qty_val = 0.0
                    unit_val = raw_fields.get("単位") or None

                    item_key = self._create_item_key_from_fields(raw_fields)
                    # Log when creating item without 数量/単位 (minimal mapping case)
                    if ("数量" not in col_indices) or ("単位" not in col_indices):
                        try:
                            logger.info(
                                "Nousei: Creating item without 数量/単位 (minimal mapping) at page %d, table %d: %s",
                                page_num + 1,
                                table_num + 1,
                                name,
                            )
                        except Exception:
                            pass

                    # Mark whether this row is a triple-dot header (metadata only)
                    try:
                        is_triple_here = (marker_text.startswith('・・・') or marker_text.startswith(
                            '･･･')) if marker_text else first.startswith('・・・')
                        if is_triple_here:
                            raw_fields["_is_triple_dot"] = "1"
                    except Exception:
                        pass

                    items.append(TenderItem(
                        item_key=item_key,
                        raw_fields=raw_fields,
                        quantity=qty_val,
                        unit=unit_val,
                        source="PDF",
                        page_number=page_num + 1
                    ))

                    # After a triple-dot item, all following zero-dot lines belong to subtable until next dotted line
                    if is_triple:
                        in_subtable_after_triple = True
                except Exception as e:
                    logger.error(
                        f"Error processing Nousei main row {row_idx + 1} in table {table_num + 1}: {e}", exc_info=True)
            return items
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
        elif project_area == "農政":
            # 農政では最低限、工種・種目 or 摘要/備考/規格のいずれかで識別
            identifying_fields = ["工種・種目", "規格", "備考", "摘要"]
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
        key_fields = ["工種・種目", "工事区分・工種・種別・細別", "摘要", "備考"]
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
        elif project_area == "農政":
            # Restrict matching to 農政-specific patterns only, to avoid affecting other formats
            pattern_sets = [self.nousei_column_patterns]
        else:
            pattern_sets = [self.column_patterns,
                            self.kitakami_column_patterns]

        for patterns_to_use in pattern_sets:
            tentative = {}
            for col_name, patterns in patterns_to_use.items():
                for i, cell in enumerate(header_row):
                    if not cell:
                        continue
                    cell_text = str(cell)
                    # Direct inclusion match
                    if any(p in cell_text for p in patterns):
                        tentative[col_name] = i
                        break
                    # Normalized match only for 農政 (remove spaces/full-width spaces and middle dots)
                    if project_area == "農政":
                        try:
                            import re as _re

                            def _clean(s: str) -> str:
                                return _re.sub(r"[\s\u3000・･]+", "", s)
                            clean_cell = _clean(cell_text)
                            if any(_clean(p) in clean_cell or clean_cell in _clean(p) for p in patterns):
                                tentative[col_name] = i
                                break
                        except Exception:
                            pass
            # Require at least quantity and unit to proceed
            if ("数量" in tentative) and ("単位" in tentative):
                # Also keep any available name/spec/remarks columns
                col_indices = tentative
                # 農政: 工種・種目が見つからない場合は第1列をデフォルトに設定
                if project_area == "農政" and "工種・種目" not in col_indices:
                    col_indices["工種・種目"] = 0
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
            # 農政 detection
            try:
                nousei_has_name = any(
                    p in header_text for p in self.nousei_column_patterns.get("工種・種目", []))
                nousei_has_remarks = any(
                    p in header_text for p in self.nousei_column_patterns.get("備考", []))
                if nousei_has_name and nousei_has_remarks:
                    return "農政"
            except Exception:
                pass
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

    def extract_subtables_with_range(self, pdf_path: str, start_page: Optional[int] = None, end_page: Optional[int] = None, reference_numbers: Optional[List[str]] = None, project_area: str = "岩手") -> List[SubtableItem]:
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
        # 農政: subtables live in the same table; use special lightweight extractor
        if project_area == "農政":
            try:
                return self._extract_nousei_subtables(pdf_path, start_page, end_page)
            except Exception as e:
                logger.error(f"Nousei subtable extraction failed: {e}")
                return []

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

    def _extract_nousei_subtables(self, pdf_path: str, start_page: Optional[int], end_page: Optional[int]) -> List[SubtableItem]:
        items: List[SubtableItem] = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                total = len(pdf.pages)
                s = 0 if start_page is None else max(0, start_page - 1)
                e = (total - 1) if end_page is None else min(total - 1, end_page - 1)
                # Continuous block numbering across entire page range
                block_index = 0
                # Persist block state across page boundaries as well
                in_block = False
                current_reference = None
                pending_reference_index = None
                for p in range(s, e + 1):
                    tables = pdf.pages[p].extract_tables() or []
                    # Persist block state across tables within the same page (already persisted across pages)
                    for table in tables:
                        if not table or len(table) < 2:
                            continue
                        # Header mapping is not required for Nousei subtables
                        header, hidx = self._find_header_row(table)
                        hidx = hidx if header else 0
                        # Scan the entire table to handle cases where marker and items split across tables
                        # Titles are not built here; comparison handles title logic. Keep block state only.
                        pending_table_title = None
                        for row in table:
                            # Determine leading text from the first non-empty cell (name cell proxy)
                            lead_text = ""
                            if row:
                                for _c in row:
                                    if _c:
                                        lead_text = str(_c).lstrip(' \t\u3000')
                                        break
                            # Marker rows starting with one or more dotted markers (・ or ･)
                            if lead_text:
                                trimmed = lead_text.lstrip(' \t\u3000')
                                if trimmed.startswith('・') or trimmed.startswith('･'):
                                    # Count consecutive leading dots (both forms)
                                    cnt = 0
                                    for ch in trimmed:
                                        if ch == '・' or ch == '･':
                                            cnt += 1
                                        else:
                                            break
                                    # Any dotted row ends the current block
                                    if in_block:
                                        in_block = False
                                        current_reference = None
                                        pending_reference_index = None
                                    # Triple-dot starts a new block; reference increments only when a row is captured
                                    if cnt == 3:
                                        in_block = True
                                        pending_reference_index = (
                                            block_index + 1)
                                        current_reference = None
                                    # Skip marker rows themselves
                                    continue
                            if not in_block:
                                continue
                            # End current subtable when encountering a completely empty row
                            try:
                                if not any((c and str(c).strip()) for c in row):
                                    in_block = False
                                    current_reference = None
                                    pending_reference_index = None
                                    continue
                            except Exception:
                                pass
                            # Skip rows that are pure equipment/asterisk rows (first cell marker) – will also check name below
                            if lead_text and (lead_text.startswith('＊') or lead_text.startswith('*')):
                                continue
                            # Heuristic extraction for name/spec/unit/quantity within a block

                            def _text(c):
                                return str(c).strip() if c is not None else ""

                            def looks_like_unit(text: str) -> bool:
                                t = (text or "").strip()
                                return t in [
                                    "m3", "m2", "m", "㎥", "㎡", "ｍ", "mm", "㎜", "cm", "㎝",
                                    "枚", "箇所", "kg", "本", "人", "日", "ha", "掛㎡", "孔", "ton", "式", "基", "台",
                                    "Ｌ", "L", "台･日", "台・日"
                                ]

                            def looks_like_quantity(text: str) -> bool:
                                import re as _re2
                                t = (text or "").replace(
                                    ',', '').replace('，', '').strip()
                                return bool(_re2.match(r"^\d+(?:\.\d+)?$", t))

                            cells = [_text(c) for c in row]
                            # Name: first non-empty cell
                            name = next((c for c in cells if c), "")
                            if name.lstrip(' \t\u3000').startswith('＊') or name.lstrip(' \t\u3000').startswith('*'):
                                continue
                            if not name:
                                continue
                            # Spec: first non-empty after name that is not unit/quantity and not 算出数量*
                            spec = ""
                            try:
                                name_idx = cells.index(name)
                            except ValueError:
                                name_idx = -1
                            for c in cells[name_idx+1:] if name_idx != -1 else cells:
                                if not c:
                                    continue
                                if c.startswith("算出数量"):
                                    continue
                                if looks_like_unit(c) or looks_like_quantity(c):
                                    continue
                                spec = c
                                break
                            # Unit
                            unit = None
                            for c in cells:
                                if looks_like_unit(c):
                                    unit = c
                                    break
                            # Quantity
                            qty = 0.0
                            for c in cells:
                                if looks_like_quantity(c):
                                    try:
                                        qty = float(
                                            c.replace(',', '').replace('，', ''))
                                    except Exception:
                                        qty = 0.0
                                    break
                            remarks = ""
                            # Activate pending reference on first valid row in this block
                            if current_reference is None and pending_reference_index is not None:
                                block_index = pending_reference_index
                                current_reference = f"内{block_index}号"
                                pending_reference_index = None
                            # Reference number is constant within the current triple-dot block
                            reference = current_reference
                            if not reference:
                                # Safety: if somehow missing, skip this row
                                continue
                            raw = {"工種・種目": name, "規格": spec, "単位": unit or "", "数量": str(
                                qty) if qty else "", "備考": remarks, "参照番号": reference}
                            # Concatenate item_key as 工種・種目 + 規格 when 規格 is present (exactly these two only)
                            item_key_value = name
                            if spec:
                                item_key_value = f"{name} {spec}".strip()
                            items.append(SubtableItem(item_key=item_key_value, raw_fields=raw, quantity=qty, unit=unit, source="PDF",
                                         page_number=p+1, reference_number=reference, sheet_name=None, table_title=None))
        except Exception as e:
            logger.error(f"Error extracting Nousei subtables: {e}")
        return items
