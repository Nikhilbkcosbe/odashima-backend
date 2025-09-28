import logging
import re
from typing import List, Dict, Tuple
from rapidfuzz import process, fuzz
from ..schemas.tender import TenderItem, SubtableItem, ComparisonResult, ComparisonSummary, SubtableComparisonResult
from .normalizer import Normalizer

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class Matcher:
    def __init__(self):
        self.normalizer = Normalizer()
        self.min_confidence = 0.8  # 80% confidence threshold for fuzzy matching

    def compare_items(self, pdf_items: List[TenderItem], excel_items: List[TenderItem]) -> ComparisonSummary:
        """
        Compare items between PDF and Excel, focusing on mismatches and missing items.
        Priority: Items in PDF that are not found in Excel.
        """
        logger.info(
            f"Starting comparison (index-based): {len(pdf_items)} PDF items vs {len(excel_items)} Excel items")

        # Index-based 1-to-1 comparison
        results: List[ComparisonResult] = []
        max_len = max(len(pdf_items), len(excel_items))

        for idx in range(max_len):
            pdf_item = pdf_items[idx] if idx < len(pdf_items) else None
            excel_item = excel_items[idx] if idx < len(excel_items) else None

            if pdf_item and excel_item:
                # Name comparison categories
                pdf_name = self.normalizer.normalize_item(pdf_item.item_key)
                excel_name = self.normalizer.normalize_item(
                    excel_item.item_key)
                pdf_tokens = [
                    t for t in self.normalizer.tokenize_item_name(pdf_item.item_key)]

                exact_name_match = (pdf_name == excel_name)
                all_tokens_in_excel = all(
                    t in excel_name for t in pdf_tokens if t)
                any_overlap = any((t and t in excel_name) for t in pdf_tokens)

                # Quantity rules with blanks
                quantity_diff = (excel_item.quantity or 0) - \
                    (pdf_item.quantity or 0)
                pdf_qty_raw = (pdf_item.raw_fields.get('数量') if hasattr(
                    pdf_item, 'raw_fields') and isinstance(pdf_item.raw_fields, dict) else None)
                excel_qty_raw = (excel_item.raw_fields.get('数量') if hasattr(
                    excel_item, 'raw_fields') and isinstance(excel_item.raw_fields, dict) else None)
                pdf_qty_blank = (pdf_qty_raw is None) or (
                    str(pdf_qty_raw).strip() == '')
                excel_qty_blank = (excel_qty_raw is None) or (
                    str(excel_qty_raw).strip() == '')
                if pdf_qty_blank and excel_qty_blank:
                    has_quantity_mismatch = True
                elif pdf_qty_blank and not excel_qty_blank:
                    has_quantity_mismatch = False
                else:
                    has_quantity_mismatch = abs(quantity_diff) >= 0.001

                # Unit comparison
                pdf_unit = self._normalize_unit(pdf_item.unit)
                excel_unit = self._normalize_unit(excel_item.unit)
                has_unit_mismatch = pdf_unit != excel_unit

                # Decide status by priority
                if not exact_name_match and not any_overlap:
                    # No overlap at all => MISSING for this position, but attach excel_item for context
                    results.append(ComparisonResult(
                        status="MISSING",
                        pdf_item=pdf_item,
                        excel_item=excel_item,
                        match_confidence=0.0,
                        quantity_difference=None,
                        unit_mismatch=None,
                        type="Main Table"
                    ))
                elif has_quantity_mismatch:
                    results.append(ComparisonResult(
                        status="QUANTITY_MISMATCH",
                        pdf_item=pdf_item,
                        excel_item=excel_item,
                        match_confidence=1.0 if exact_name_match else 0.0,
                        quantity_difference=quantity_diff,
                        unit_mismatch=has_unit_mismatch,
                        type="Main Table"
                    ))
                elif has_unit_mismatch:
                    results.append(ComparisonResult(
                        status="UNIT_MISMATCH",
                        pdf_item=pdf_item,
                        excel_item=excel_item,
                        match_confidence=1.0 if exact_name_match else 0.0,
                        quantity_difference=None,
                        unit_mismatch=True,
                        type="Main Table"
                    ))
                elif not exact_name_match:
                    # Name mismatch categories 2 or 3
                    results.append(ComparisonResult(
                        status="NAME_MISMATCH",
                        pdf_item=pdf_item,
                        excel_item=excel_item,
                        match_confidence=0.0,
                        quantity_difference=None,
                        unit_mismatch=None,
                        type="Main Table"
                    ))
                else:
                    results.append(ComparisonResult(
                        status="OK",
                        pdf_item=pdf_item,
                        excel_item=excel_item,
                        match_confidence=1.0,
                        quantity_difference=None,
                        unit_mismatch=False,
                        type="Main Table"
                    ))
            elif pdf_item and not excel_item:
                # Excel ran out of items
                results.append(ComparisonResult(
                    status="MISSING",
                    pdf_item=pdf_item,
                    excel_item=None,
                    match_confidence=0.0,
                    quantity_difference=None,
                    unit_mismatch=None,
                    type="Main Table"
                ))
            elif excel_item and not pdf_item:
                results.append(ComparisonResult(
                    status="EXTRA",
                    pdf_item=None,
                    excel_item=excel_item,
                    match_confidence=0.0,
                    quantity_difference=None,
                    unit_mismatch=None,
                    type="Main Table"
                ))

        # Generate summary focusing on mismatches and missing items
        summary = self._generate_summary(results)

        # Log key findings
        self._log_summary(summary)

        return summary

    def compare_subtable_items(self, pdf_subtables: List[SubtableItem], excel_subtables: List[SubtableItem]) -> List[SubtableComparisonResult]:
        """
        Category-based subtable comparison within each reference (no index alignment):
        - Exact name match -> apply quantity/unit rules
        - All tokens from PDF present in an Excel name -> NAME_MISMATCH
        - Some token overlap -> NAME_MISMATCH
        - No overlap in ref -> MISSING
        Also emit EXTRA for unmatched Excel items per reference.
        """
        # Group by normalized reference number
        def norm_ref(ref: str) -> str:
            if not ref:
                return ''
            # Normalize spaces and common width differences
            normalized = ref.replace(' ', '').replace('　', '')
            # If Excel repeated-reference suffixes like "-2", "-3" were added, strip them for matching
            normalized = re.sub(r"-\d+$", "", normalized)
            return normalized

        # Kitakami reference equivalence: treat 明3号 == 第3号明, 施12号 == 第12号施, etc.
        # Rule: the trailing Kanji in PDF form (第N号X) or leading Kanji in Excel form (XN号) must match,
        # N号 must match, and the presence of 第 is optional.
        def kitakami_ref_key(ref: str) -> str:
            if not ref:
                return ''
            r = norm_ref(ref)
            # Extract patterns: 第?N号X  or  XN号
            m_pdf = re.match(r'^第?(\d+)号([一-龯])$', r)
            if m_pdf:
                num, tail = m_pdf.group(1), m_pdf.group(2)
                return f"{tail}:{num}"
            m_excel = re.match(r'^([一-龯])(\d+)号$', r)
            if m_excel:
                tail, num = m_excel.group(1), m_excel.group(2)
                return f"{tail}:{num}"
            # Fallback: return normalized as-is to keep grouping for non-Kitakami patterns
            return r

        from collections import defaultdict
        pdf_by_ref = defaultdict(list)
        excel_by_ref = defaultdict(list)
        for item in pdf_subtables:
            pdf_by_ref[kitakami_ref_key(
                getattr(item, 'reference_number', '')).strip()].append(item)
        for item in excel_subtables:
            excel_by_ref[kitakami_ref_key(
                getattr(item, 'reference_number', '')).strip()].append(item)

        all_refs = set(pdf_by_ref.keys()) | set(excel_by_ref.keys())
        results: List[SubtableComparisonResult] = []

        for ref in all_refs:
            pdf_list = pdf_by_ref.get(ref, [])
            excel_list = excel_by_ref.get(ref, [])

            matched_excel_indices = set()

            for pdf_item in pdf_list:
                pdf_name = self.normalizer.normalize_item(pdf_item.item_key)
                pdf_tokens = [
                    t for t in self.normalizer.tokenize_item_name(pdf_item.item_key)]

                # Exact match within same ref
                exact_idx = None
                for i, ex in enumerate(excel_list):
                    if i in matched_excel_indices:
                        continue
                    if self.normalizer.normalize_item(ex.item_key) == pdf_name:
                        exact_idx = i
                        break

                if exact_idx is not None:
                    ex = excel_list[exact_idx]
                    matched_excel_indices.add(exact_idx)
                    # Quantity with blank rules
                    quantity_diff = (ex.quantity or 0) - \
                        (pdf_item.quantity or 0)
                    pdf_qty_raw = (pdf_item.raw_fields.get('数量') if hasattr(
                        pdf_item, 'raw_fields') and isinstance(pdf_item.raw_fields, dict) else None)
                    excel_qty_raw = (ex.raw_fields.get('数量') if hasattr(
                        ex, 'raw_fields') and isinstance(ex.raw_fields, dict) else None)
                    pdf_qty_blank = (pdf_qty_raw is None) or (
                        str(pdf_qty_raw).strip() == '')
                    excel_qty_blank = (excel_qty_raw is None) or (
                        str(excel_qty_raw).strip() == '')
                    if pdf_qty_blank and excel_qty_blank:
                        has_quantity_mismatch = True
                    elif pdf_qty_blank and not excel_qty_blank:
                        has_quantity_mismatch = False
                    else:
                        has_quantity_mismatch = abs(quantity_diff) >= 0.001
                    pdf_unit = self._normalize_unit(pdf_item.unit)
                    excel_unit = self._normalize_unit(ex.unit)
                    has_unit_mismatch = pdf_unit != excel_unit

                    if has_quantity_mismatch:
                        results.append(SubtableComparisonResult(status="QUANTITY_MISMATCH", pdf_item=pdf_item, excel_item=ex,
                                       match_confidence=1.0, quantity_difference=quantity_diff, unit_mismatch=has_unit_mismatch, type="Sub Table"))
                    elif has_unit_mismatch:
                        results.append(SubtableComparisonResult(status="UNIT_MISMATCH", pdf_item=pdf_item, excel_item=ex,
                                       match_confidence=1.0, quantity_difference=None, unit_mismatch=True, type="Sub Table"))
                    else:
                        results.append(SubtableComparisonResult(status="OK", pdf_item=pdf_item, excel_item=ex,
                                       match_confidence=1.0, quantity_difference=None, unit_mismatch=False, type="Sub Table"))
                    continue

                # Category 2: all tokens present in some Excel name
                cat2_idx = None
                for i, ex in enumerate(excel_list):
                    if i in matched_excel_indices:
                        continue
                    ex_name = self.normalizer.normalize_item(ex.item_key)
                    if all(t in ex_name for t in pdf_tokens if t):
                        cat2_idx = i
                        break

                if cat2_idx is not None:
                    ex = excel_list[cat2_idx]
                    matched_excel_indices.add(cat2_idx)
                    # Promote to quantity/unit mismatch if applicable; otherwise name mismatch
                    quantity_diff = (ex.quantity or 0) - \
                        (pdf_item.quantity or 0)
                    pdf_qty_raw = (pdf_item.raw_fields.get('数量') if hasattr(
                        pdf_item, 'raw_fields') and isinstance(pdf_item.raw_fields, dict) else None)
                    excel_qty_raw = (ex.raw_fields.get('数量') if hasattr(
                        ex, 'raw_fields') and isinstance(ex.raw_fields, dict) else None)
                    pdf_qty_blank = (pdf_qty_raw is None) or (
                        str(pdf_qty_raw).strip() == '')
                    excel_qty_blank = (excel_qty_raw is None) or (
                        str(excel_qty_raw).strip() == '')
                    if pdf_qty_blank and excel_qty_blank:
                        has_quantity_mismatch = True
                    elif pdf_qty_blank and not excel_qty_blank:
                        has_quantity_mismatch = False
                    else:
                        has_quantity_mismatch = abs(quantity_diff) >= 0.001
                    pdf_unit = self._normalize_unit(pdf_item.unit)
                    excel_unit = self._normalize_unit(ex.unit)
                    has_unit_mismatch = pdf_unit != excel_unit
                    if has_quantity_mismatch:
                        results.append(SubtableComparisonResult(status="QUANTITY_MISMATCH", pdf_item=pdf_item, excel_item=ex,
                                       match_confidence=0.0, quantity_difference=quantity_diff, unit_mismatch=has_unit_mismatch, type="Sub Table"))
                    elif has_unit_mismatch:
                        results.append(SubtableComparisonResult(status="UNIT_MISMATCH", pdf_item=pdf_item, excel_item=ex,
                                       match_confidence=0.0, quantity_difference=None, unit_mismatch=True, type="Sub Table"))
                    else:
                        results.append(SubtableComparisonResult(status="NAME_MISMATCH", pdf_item=pdf_item, excel_item=ex,
                                       match_confidence=0.0, quantity_difference=None, unit_mismatch=None, type="Sub Table"))
                    continue

                # Category 3: any token overlap
                cat3_idx = None
                for i, ex in enumerate(excel_list):
                    if i in matched_excel_indices:
                        continue
                    ex_name = self.normalizer.normalize_item(ex.item_key)
                    if any((t and t in ex_name) for t in pdf_tokens):
                        cat3_idx = i
                        break

                if cat3_idx is not None:
                    ex = excel_list[cat3_idx]
                    matched_excel_indices.add(cat3_idx)
                    # Promote to quantity/unit mismatch if applicable; otherwise name mismatch
                    quantity_diff = (ex.quantity or 0) - \
                        (pdf_item.quantity or 0)
                    pdf_qty_raw = (pdf_item.raw_fields.get('数量') if hasattr(
                        pdf_item, 'raw_fields') and isinstance(pdf_item.raw_fields, dict) else None)
                    excel_qty_raw = (ex.raw_fields.get('数量') if hasattr(
                        ex, 'raw_fields') and isinstance(ex.raw_fields, dict) else None)
                    pdf_qty_blank = (pdf_qty_raw is None) or (
                        str(pdf_qty_raw).strip() == '')
                    excel_qty_blank = (excel_qty_raw is None) or (
                        str(excel_qty_raw).strip() == '')
                    if pdf_qty_blank and excel_qty_blank:
                        has_quantity_mismatch = True
                    elif pdf_qty_blank and not excel_qty_blank:
                        has_quantity_mismatch = False
                    else:
                        has_quantity_mismatch = abs(quantity_diff) >= 0.001
                    pdf_unit = self._normalize_unit(pdf_item.unit)
                    excel_unit = self._normalize_unit(ex.unit)
                    has_unit_mismatch = pdf_unit != excel_unit
                    if has_quantity_mismatch:
                        results.append(SubtableComparisonResult(status="QUANTITY_MISMATCH", pdf_item=pdf_item, excel_item=ex,
                                       match_confidence=0.0, quantity_difference=quantity_diff, unit_mismatch=has_unit_mismatch, type="Sub Table"))
                    elif has_unit_mismatch:
                        results.append(SubtableComparisonResult(status="UNIT_MISMATCH", pdf_item=pdf_item, excel_item=ex,
                                       match_confidence=0.0, quantity_difference=None, unit_mismatch=True, type="Sub Table"))
                    else:
                        results.append(SubtableComparisonResult(status="NAME_MISMATCH", pdf_item=pdf_item, excel_item=ex,
                                       match_confidence=0.0, quantity_difference=None, unit_mismatch=None, type="Sub Table"))
                    continue

                # No overlap: attach best-effort Excel candidate from same reference for context
                candidate = None
                try:
                    pdf_idx_in_ref = pdf_list.index(pdf_item)
                except ValueError:
                    pdf_idx_in_ref = -1

                # Prefer same-index excel if available and not already matched
                if 0 <= pdf_idx_in_ref < len(excel_list) and pdf_idx_in_ref not in matched_excel_indices:
                    candidate = excel_list[pdf_idx_in_ref]
                else:
                    # Otherwise, attach first unmatched excel in this reference (if any)
                    for i, ex in enumerate(excel_list):
                        if i not in matched_excel_indices:
                            candidate = ex
                            break

                results.append(SubtableComparisonResult(status="MISSING", pdf_item=pdf_item, excel_item=candidate,
                               match_confidence=0.0, quantity_difference=None, unit_mismatch=None, type="Sub Table"))

            # EXTRA for unmatched excel items
            for i, ex in enumerate(excel_list):
                if i not in matched_excel_indices:
                    results.append(SubtableComparisonResult(status="EXTRA", pdf_item=None, excel_item=ex,
                                   match_confidence=0.0, quantity_difference=None, unit_mismatch=None, type="Sub Table"))

        return results

    def _normalize_items(self, items: List[TenderItem], source: str) -> Dict[str, TenderItem]:
        """
        Normalize a list of items into a dictionary with normalized keys.
        """
        normalized = {}
        duplicates = 0

        for item in items:
            normalized_key = self.normalizer.normalize_item(item.item_key)

            if normalized_key in normalized:
                # Handle duplicates by appending source info
                duplicates += 1
                original_key = normalized_key
                counter = 1
                while normalized_key in normalized:
                    normalized_key = f"{original_key}_dup_{counter}"
                    counter += 1
                logger.warning(
                    f"Duplicate key found in {source}: {original_key} -> {normalized_key}")

            normalized[normalized_key] = item

        if duplicates > 0:
            logger.warning(f"Found {duplicates} duplicate keys in {source}")

        return normalized

    def _compare_single_pdf_item(self, pdf_key: str, pdf_item: TenderItem,
                                 excel_normalized: Dict[str, TenderItem],
                                 matched_excel_keys: set) -> ComparisonResult:
        """
        Compare a single PDF item against all Excel items.
        """
        # Try exact match first (Category 1)
        if pdf_key in excel_normalized:
            excel_item = excel_normalized[pdf_key]
            return self._create_comparison_result(pdf_item, excel_item, 1.0, "EXACT_MATCH")

        # Try fuzzy matching with high confidence threshold
        fuzzy_match = self._find_fuzzy_match(
            pdf_key, excel_normalized, matched_excel_keys)

        if fuzzy_match:
            excel_key, excel_item, confidence = fuzzy_match
            match_type = "FUZZY_MATCH" if confidence < 1.0 else "EXACT_MATCH"
            return self._create_comparison_result(pdf_item, excel_item, confidence, match_type)

        # New policy: do not emit MISSING; classify by name match category instead
        # Category 2: all words/characters from PDF present in some Excel item name
        pdf_tokens = [
            t for t in self.normalizer.tokenize_item_name(pdf_item.item_key)]
        for excel_key, excel_item in excel_normalized.items():
            excel_name = self.normalizer.normalize_item(excel_item.item_key)
            if all(t in excel_name for t in pdf_tokens if t):
                # Treat as name mismatch, not missing
                return ComparisonResult(
                    status="NAME_MISMATCH",
                    pdf_item=pdf_item,
                    excel_item=excel_item,
                    match_confidence=0.0,
                    quantity_difference=None,
                    unit_mismatch=None,
                    type="Main Table"
                )

        # Category 3: some substring or word overlap between PDF and Excel item names
        for excel_key, excel_item in excel_normalized.items():
            excel_name = self.normalizer.normalize_item(excel_item.item_key)
            if any((t and t in excel_name) for t in pdf_tokens):
                return ComparisonResult(
                    status="NAME_MISMATCH",
                    pdf_item=pdf_item,
                    excel_item=excel_item,
                    match_confidence=0.0,
                    quantity_difference=None,
                    unit_mismatch=None,
                    type="Main Table"
                )

        # If no overlap found at all, treat as MISSING per requirement
        logger.info("Missing item in Excel (no name overlap): %s",
                    pdf_item.item_key)
        # Attach a best-effort Excel candidate (closest by fuzzy without threshold) for context
        best_candidate = None
        best_score = -1
        try:
            for excel_key, excel_item in excel_normalized.items():
                score = fuzz.ratio(pdf_key, excel_key)
                if score > best_score:
                    best_score = score
                    best_candidate = excel_item
        except Exception:
            best_candidate = None

        return ComparisonResult(
            status="MISSING",
            pdf_item=pdf_item,
            excel_item=best_candidate,
            match_confidence=0.0,
            quantity_difference=None,
            unit_mismatch=None,
            type="Main Table"
        )

    def _find_fuzzy_match(self, pdf_key: str, excel_normalized: Dict[str, TenderItem],
                          matched_excel_keys: set) -> Tuple[str, TenderItem, float]:
        """
        Find the best fuzzy match for a PDF item in Excel items.
        Enhanced with better validation to prevent false positive matches.
        """
        # Get available Excel keys (not yet matched)
        available_excel_keys = [
            key for key in excel_normalized.keys() if key not in matched_excel_keys]

        if not available_excel_keys:
            return None

        # Try fuzzy matching
        matches = process.extractOne(
            pdf_key,
            available_excel_keys,
            scorer=fuzz.ratio,
            score_cutoff=self.min_confidence * 100
        )

        if matches:
            # Handle different return formats from rapidfuzz
            if isinstance(matches, tuple) and len(matches) >= 2:
                excel_key = matches[0]
                confidence_score = matches[1]
            else:
                excel_key = matches
                confidence_score = 100  # Default to high confidence

            excel_item = excel_normalized[excel_key]
            confidence = confidence_score / 100.0

            # ENHANCED: Check if items are significantly different despite high similarity score
            # This prevents false matches between similar but distinct items
            if self.normalizer.are_items_significantly_different(pdf_key, excel_key):
                logger.info(
                    f"Rejecting potential match due to significant differences:")
                logger.info(f"  PDF: '{pdf_key}'")
                logger.info(f"  Excel: '{excel_key}'")
                logger.info(
                    f"  Fuzzy score was {confidence:.2f} but items are significantly different")
                return None

            logger.debug(f"Fuzzy match found: {confidence:.2f} confidence")
            return excel_key, excel_item, confidence

        return None

    def _normalize_unit(self, unit: str) -> str:
        """
        Normalize unit values for comparison.
        Handles None values, common unit variations, and FULL-WIDTH vs HALF-WIDTH characters.
        """
        if not unit:
            return ""

        # Strip whitespace
        normalized = str(unit).strip()

        # CRITICAL FIX: Convert full-width characters to half-width BEFORE other processing
        # This handles cases like "ｍ" (full-width) vs "m" (half-width)
        full_to_half_map = str.maketrans(
            'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ'
            'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ'
            '０１２３４５６７８９'
            '！＂＃＄％＆＇（）＊＋，－．／：；＜＝＞？＠［＼］＾＿｀｛｜｝～',
            'abcdefghijklmnopqrstuvwxyz'
            'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
            '0123456789'
            '!"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~'
        )
        normalized = normalized.translate(full_to_half_map)

        # Convert to lowercase after width normalization
        normalized = normalized.lower()

        # Handle common variations
        unit_mappings = {
            "m2": "㎡",
            "m3": "㎥",
            "m²": "㎡",
            "m³": "㎥",
            "平方メートル": "㎡",
            "立方メートル": "㎥",
            "メートル": "m",
            "センチメートル": "cm",
            "ミリメートル": "mm",
            "キログラム": "kg",
            "グラム": "g",
            "リットル": "l",
            "ℓ": "l",
            # Additional mappings for full-width units that might remain
            "ｍ": "m",  # Full-width m -> half-width m
            "ｔ": "t",  # Full-width t -> half-width t
            "ｋｇ": "kg",  # Full-width kg -> half-width kg
        }

        result = unit_mappings.get(normalized, normalized)

        # Debug logging for unit mismatches
        if unit != result:
            logger.debug(f"Unit normalized: '{unit}' -> '{result}'")

        return result

    def _create_comparison_result(self, pdf_item: TenderItem, excel_item: TenderItem,
                                  confidence: float, match_type: str) -> ComparisonResult:
        """
        Create a comparison result for matched items, checking for quantity and unit differences.
        """
        # Check quantity difference with blank rules
        quantity_tolerance = 0.001
        quantity_diff = (excel_item.quantity or 0) - (pdf_item.quantity or 0)
        pdf_qty_raw = (pdf_item.raw_fields.get('数量') if hasattr(
            pdf_item, 'raw_fields') and isinstance(pdf_item.raw_fields, dict) else None)
        excel_qty_raw = (excel_item.raw_fields.get('数量') if hasattr(
            excel_item, 'raw_fields') and isinstance(excel_item.raw_fields, dict) else None)
        pdf_qty_blank = (pdf_qty_raw is None) or (
            str(pdf_qty_raw).strip() == '')
        excel_qty_blank = (excel_qty_raw is None) or (
            str(excel_qty_raw).strip() == '')
        if pdf_qty_blank and excel_qty_blank:
            has_quantity_mismatch = True
        elif pdf_qty_blank and not excel_qty_blank:
            has_quantity_mismatch = False
        else:
            has_quantity_mismatch = abs(quantity_diff) >= quantity_tolerance

        # Check unit difference
        pdf_unit = self._normalize_unit(pdf_item.unit)
        excel_unit = self._normalize_unit(excel_item.unit)
        has_unit_mismatch = pdf_unit != excel_unit

        # Determine status based on mismatches
        if has_quantity_mismatch and has_unit_mismatch:
            status = "QUANTITY_MISMATCH"  # Quantity mismatch takes priority
            actual_quantity_diff = quantity_diff
            logger.info(f"Quantity & Unit mismatch: {pdf_item.item_key[:30]}... "
                        f"(PDF: {pdf_item.quantity} {pdf_item.unit}, Excel: {excel_item.quantity} {excel_item.unit})")
        elif has_quantity_mismatch:
            status = "QUANTITY_MISMATCH"
            actual_quantity_diff = quantity_diff
            logger.info(f"Quantity mismatch: {pdf_item.item_key[:30]}... "
                        f"(PDF: {pdf_item.quantity}, Excel: {excel_item.quantity}, Diff: {quantity_diff})")
        elif has_unit_mismatch:
            status = "UNIT_MISMATCH"
            actual_quantity_diff = None
            logger.info(f"Unit mismatch: {pdf_item.item_key[:30]}... "
                        f"(PDF: '{pdf_item.unit}', Excel: '{excel_item.unit}')")
        else:
            status = "OK"
            actual_quantity_diff = None
            logger.debug(
                f"Perfect match: {pdf_item.item_key[:30]}... (quantities and units match)")

        return ComparisonResult(
            status=status,
            pdf_item=pdf_item,
            excel_item=excel_item,
            match_confidence=confidence,
            quantity_difference=actual_quantity_diff,
            unit_mismatch=has_unit_mismatch,
            type="Main Table"
        )

    def _find_extra_excel_items(self, excel_normalized: Dict[str, TenderItem],
                                matched_excel_keys: set) -> List[ComparisonResult]:
        """
        Find Excel items that have no corresponding PDF item.
        """
        extra_results = []

        for excel_key, excel_item in excel_normalized.items():
            if excel_key not in matched_excel_keys:
                logger.info(
                    f"Extra item in Excel (not in PDF): {excel_item.item_key}")
                extra_results.append(ComparisonResult(
                    status="EXTRA",
                    pdf_item=None,
                    excel_item=excel_item,
                    match_confidence=0.0,
                    quantity_difference=None,
                    unit_mismatch=None,
                    type="Main Table"
                ))

        return extra_results

    def _generate_summary(self, results: List[ComparisonResult]) -> ComparisonSummary:
        """
        Generate a comprehensive summary of the comparison results.
        """
        total_items = len(results)
        matched_items = sum(1 for r in results if r.status == "OK")
        quantity_mismatches = sum(
            1 for r in results if r.status == "QUANTITY_MISMATCH")
        unit_mismatches = sum(
            1 for r in results if r.status == "UNIT_MISMATCH")
        # Count genuine MISSING entries (main-table policy now allows missing if no overlap)
        missing_items = sum(1 for r in results if r.status == "MISSING")
        extra_items = sum(1 for r in results if r.status == "EXTRA")

        return ComparisonSummary(
            total_items=total_items,
            matched_items=matched_items,
            quantity_mismatches=quantity_mismatches,
            unit_mismatches=unit_mismatches,
            missing_items=missing_items,
            extra_items=extra_items,
            results=results
        )

    def _log_summary(self, summary: ComparisonSummary):
        """
        Log a detailed summary of the comparison results.
        """
        logger.info("=== COMPARISON SUMMARY ===")
        logger.info(f"Total items processed: {summary.total_items}")
        logger.info(f"Perfect matches: {summary.matched_items}")
        logger.info(f"Quantity mismatches: {summary.quantity_mismatches}")
        logger.info(f"Unit mismatches: {summary.unit_mismatches}")
        logger.info(f"Missing in Excel: {summary.missing_items}")
        logger.info(f"Extra in Excel: {summary.extra_items}")

        # Log details of missing items (high priority)
        if summary.missing_items > 0:
            logger.warning("=== MISSING ITEMS (PDF → Excel) ===")
            missing_count = 0
            for result in summary.results:
                if result.status == "MISSING" and missing_count < 10:  # Log first 10
                    logger.warning(f"Missing: {result.pdf_item.item_key}")
                    missing_count += 1
            if summary.missing_items > 10:
                logger.warning(
                    f"... and {summary.missing_items - 10} more missing items")

        # Log details of quantity mismatches
        if summary.quantity_mismatches > 0:
            logger.warning("=== QUANTITY MISMATCHES ===")
            mismatch_count = 0
            for result in summary.results:
                if result.status == "QUANTITY_MISMATCH" and mismatch_count < 5:  # Log first 5
                    logger.warning(f"Quantity diff: {result.pdf_item.item_key[:30]}... "
                                   f"(Diff: {result.quantity_difference})")
                    mismatch_count += 1
            if summary.quantity_mismatches > 5:
                logger.warning(
                    f"... and {summary.quantity_mismatches - 5} more quantity mismatches")

        # Log details of unit mismatches
        if summary.unit_mismatches > 0:
            logger.warning("=== UNIT MISMATCHES ===")
            unit_mismatch_count = 0
            for result in summary.results:
                if result.status == "UNIT_MISMATCH" and unit_mismatch_count < 5:  # Log first 5
                    logger.warning(f"Unit diff: {result.pdf_item.item_key[:30]}... "
                                   f"(PDF: '{result.pdf_item.unit}', Excel: '{result.excel_item.unit}')")
                    unit_mismatch_count += 1
            if summary.unit_mismatches > 5:
                logger.warning(
                    f"... and {summary.unit_mismatches - 5} more unit mismatches")

        logger.info("=== END SUMMARY ===")

    def get_missing_items_only(self, pdf_items: List[TenderItem], excel_items: List[TenderItem]) -> List[TenderItem]:
        """
        Quick method to get only the missing items (PDF items not found in Excel).
        """
        logger.info("Getting missing items only...")

        comparison_summary = self.compare_items(pdf_items, excel_items)

        missing_items = []
        for result in comparison_summary.results:
            if result.status == "MISSING":
                missing_items.append(result.pdf_item)

        logger.info(f"Found {len(missing_items)} missing items")
        return missing_items

    def get_missing_items_by_name_only_strict(self, pdf_items: List[TenderItem], excel_items: List[TenderItem]) -> List[TenderItem]:
        """
        Return missing items considering ONLY item name (strict match), ignoring quantity and unit.
        Strict means exact string equality after trimming outer whitespace.
        No fuzzy/normalization is applied.
        """
        logger.info("Getting missing items by strict name-only matching...")

        # Use the same normalizer as the rest of the system to eliminate
        # whitespace/newline/full-width differences while staying name-only
        excel_names = set()
        for item in excel_items:
            if item and item.item_key is not None:
                name = self.normalizer.normalize_item(str(item.item_key))
                if name:
                    excel_names.add(name)

        missing: List[TenderItem] = []
        for pdf_item in pdf_items:
            if pdf_item and pdf_item.item_key is not None:
                pdf_name = self.normalizer.normalize_item(
                    str(pdf_item.item_key))
                if pdf_name and pdf_name not in excel_names:
                    missing.append(pdf_item)

        logger.info(f"Found {len(missing)} strict name-only missing items")
        return missing

    def get_mismatched_items_only(self, pdf_items: List[TenderItem], excel_items: List[TenderItem]) -> List[ComparisonResult]:
        """
        Quick method to get only the mismatched items (quantity differences).
        """
        logger.info("Getting mismatched items only...")

        comparison_summary = self.compare_items(pdf_items, excel_items)

        mismatched_results = []
        for result in comparison_summary.results:
            if result.status == "QUANTITY_MISMATCH":
                mismatched_results.append(result)

        logger.info(f"Found {len(mismatched_results)} quantity mismatches")
        return mismatched_results

    def get_unit_mismatched_items_only(self, pdf_items: List[TenderItem], excel_items: List[TenderItem]) -> List[ComparisonResult]:
        """
        Quick method to get only the unit mismatched items.
        """
        logger.info("Getting unit mismatched items only...")

        comparison_summary = self.compare_items(pdf_items, excel_items)

        unit_mismatched_results = []
        for result in comparison_summary.results:
            if result.status == "UNIT_MISMATCH":
                unit_mismatched_results.append(result)

        logger.info(f"Found {len(unit_mismatched_results)} unit mismatches")
        return unit_mismatched_results

    def get_extra_items_only(self, pdf_items: List[TenderItem], excel_items: List[TenderItem]) -> List[TenderItem]:
        """
        Quick method to get only the extra items (Excel items not found in PDF).
        This compares main table items only.
        """
        logger.info("Getting extra main table items only...")

        comparison_summary = self.compare_items(pdf_items, excel_items)

        extra_items = []
        for result in comparison_summary.results:
            if result.status == "EXTRA":
                extra_items.append(result.excel_item)

        logger.info(f"Found {len(extra_items)} extra main table items")
        return extra_items

    def get_extra_subtable_items_only(self, pdf_subtables: List[SubtableItem], excel_subtables: List[SubtableItem]) -> List[SubtableItem]:
        """
        Quick method to get only the extra subtable items (Excel subtable items not found in PDF subtables).
        """
        logger.info("Getting extra subtable items only...")

        subtable_results = self.compare_subtable_items(
            pdf_subtables, excel_subtables)

        extra_subtable_items = []
        for result in subtable_results:
            if result.status == "EXTRA":
                extra_subtable_items.append(result.excel_item)

        logger.info(f"Found {len(extra_subtable_items)} extra subtable items")
        return extra_subtable_items

    def get_extra_items_only_simplified(self, pdf_items: List[TenderItem], excel_items: List[TenderItem]) -> List[TenderItem]:
        """
        Simplified method to get only the extra items (Excel items not found in PDF).
        Compares ONLY item name, quantity, and unit as requested by the user.
        Uses the improved normalizer to handle complex characters properly.
        """
        logger.info(
            "Getting extra main table items with simplified matching...")

        def items_match(pdf_item: TenderItem, excel_item: TenderItem) -> bool:
            """Check if items match based only on name, quantity, and unit"""
            # Compare normalized item names using improved normalizer
            pdf_name = self.normalizer.normalize_item(pdf_item.item_key)
            excel_name = self.normalizer.normalize_item(excel_item.item_key)

            # Check name similarity (exact match or very close)
            names_match = (pdf_name == excel_name or
                           pdf_name in excel_name or
                           excel_name in pdf_name or
                           (len(pdf_name) > 10 and len(excel_name) > 10 and
                            abs(len(pdf_name) - len(excel_name)) <= 2 and
                            pdf_name[:10] == excel_name[:10]))

            if not names_match:
                logger.debug(
                    f"Names don't match: PDF='{pdf_name}' vs Excel='{excel_name}'")
                return False

            # Compare quantities (with small tolerance)
            quantity_match = abs(pdf_item.quantity -
                                 excel_item.quantity) < 0.001

            if not quantity_match:
                logger.debug(
                    f"Quantities don't match: PDF={pdf_item.quantity} vs Excel={excel_item.quantity}")
                return False

            # Compare units (normalize and compare) - FIXED to use unit-specific normalization
            pdf_unit = self._normalize_unit(pdf_item.unit)
            excel_unit = self._normalize_unit(excel_item.unit)
            unit_match = pdf_unit == excel_unit

            if not unit_match:
                logger.debug(
                    f"Units don't match: PDF='{pdf_unit}' vs Excel='{excel_unit}'")
                return False

            return True

        # Find Excel items that don't have a matching PDF item
        extra_items = []

        for excel_item in excel_items:
            found_match = False

            for pdf_item in pdf_items:
                if items_match(pdf_item, excel_item):
                    found_match = True
                    logger.debug(
                        f"Found match: Excel '{excel_item.item_key}' matches PDF '{pdf_item.item_key}'")
                    break

            if not found_match:
                extra_items.append(excel_item)
                logger.info(
                    f"Extra item found: '{excel_item.item_key}' (qty: {excel_item.quantity}, unit: {excel_item.unit})")

        logger.info(
            f"Found {len(extra_items)} extra main table items using simplified matching")
        return extra_items
