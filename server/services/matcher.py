import logging
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
            f"Starting comparison: {len(pdf_items)} PDF items vs {len(excel_items)} Excel items")

        # Normalize all items for comparison
        pdf_normalized = self._normalize_items(pdf_items, "PDF")
        excel_normalized = self._normalize_items(excel_items, "Excel")

        logger.info(
            f"Normalized: {len(pdf_normalized)} PDF items, {len(excel_normalized)} Excel items")

        # Compare items iteratively
        results = []
        matched_excel_keys = set()

        # Process each PDF item to find mismatches and missing items
        for pdf_idx, (pdf_key, pdf_item) in enumerate(pdf_normalized.items()):
            logger.debug(
                f"Processing PDF item {pdf_idx + 1}/{len(pdf_normalized)}: {pdf_key[:50]}...")

            comparison_result = self._compare_single_pdf_item(
                pdf_key, pdf_item, excel_normalized, matched_excel_keys
            )

            results.append(comparison_result)

            # Track which Excel items have been matched
            if comparison_result.excel_item and comparison_result.status in ["OK", "QUANTITY_MISMATCH"]:
                excel_normalized_key = self.normalizer.normalize_item(
                    comparison_result.excel_item.item_key)
                matched_excel_keys.add(excel_normalized_key)

        # Find extra items in Excel (not requested but included for completeness)
        extra_excel_results = self._find_extra_excel_items(
            excel_normalized, matched_excel_keys)
        results.extend(extra_excel_results)

        # Generate summary focusing on mismatches and missing items
        summary = self._generate_summary(results)

        # Log key findings
        self._log_summary(summary)

        return summary

    def compare_subtable_items(self, pdf_subtables: List[SubtableItem], excel_subtables: List[SubtableItem]) -> List[SubtableComparisonResult]:
        """
        Compare subtable items between PDF and Excel, focusing on mismatches and missing items.
        ENHANCED: Uses normalization for better matching (similar to main table logic).
        """
        # ENHANCED: Normalize keys for matching using the normalizer
        def make_key(item):
            # Normalize item name using the normalizer
            normalized_item_key = self.normalizer.normalize_item(item.item_key) if item.item_key else ''
            
            # Normalize reference number (handle spacing issues like "単 1号" vs "単1号")
            normalized_ref = item.reference_number.strip().replace(' ', '').replace('　', '') if item.reference_number else ''
            
            return (normalized_item_key, normalized_ref)

        pdf_dict = {make_key(item): item for item in pdf_subtables}
        excel_dict = {make_key(item): item for item in excel_subtables}

        results = []
        matched_excel_keys = set()

        # Compare PDF subtables to Excel subtables
        for key, pdf_item in pdf_dict.items():
            if key in excel_dict:
                excel_item = excel_dict[key]
                # Compare quantity and unit
                quantity_diff = (excel_item.quantity or 0) - \
                    (pdf_item.quantity or 0)
                has_quantity_mismatch = abs(quantity_diff) >= 0.001
                # FIXED: Use normalized unit comparison to handle full-width vs half-width characters
                pdf_unit = self._normalize_unit(pdf_item.unit)
                excel_unit = self._normalize_unit(excel_item.unit)
                has_unit_mismatch = pdf_unit != excel_unit
                if has_quantity_mismatch:
                    status = "QUANTITY_MISMATCH"
                elif has_unit_mismatch:
                    status = "UNIT_MISMATCH"
                else:
                    status = "OK"
                results.append(SubtableComparisonResult(
                    status=status,
                    pdf_item=pdf_item,
                    excel_item=excel_item,
                    match_confidence=1.0,
                    quantity_difference=quantity_diff if has_quantity_mismatch else None,
                    unit_mismatch=has_unit_mismatch if has_unit_mismatch else None,
                    type="Sub Table"
                ))
                matched_excel_keys.add(key)
            else:
                # Missing in Excel
                results.append(SubtableComparisonResult(
                    status="MISSING",
                    pdf_item=pdf_item,
                    excel_item=None,
                    match_confidence=0.0,
                    quantity_difference=None,
                    unit_mismatch=None,
                    type="Sub Table"
                ))
        # Find extra items in Excel
        for key, excel_item in excel_dict.items():
            if key not in matched_excel_keys:
                results.append(SubtableComparisonResult(
                    status="EXTRA",
                    pdf_item=None,
                    excel_item=excel_item,
                    match_confidence=0.0,
                    quantity_difference=None,
                    unit_mismatch=None,
                    type="Sub Table"
                ))
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
        # Try exact match first
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

        # No match found - this is a missing item
        logger.info(f"Missing item in Excel: {pdf_item.item_key}")
        return ComparisonResult(
            status="MISSING",
            pdf_item=pdf_item,
            excel_item=None,
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
            "ｋｇ": "kg", # Full-width kg -> half-width kg
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
        # Check quantity difference
        quantity_tolerance = 0.001
        quantity_diff = excel_item.quantity - pdf_item.quantity
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

        subtable_results = self.compare_subtable_items(pdf_subtables, excel_subtables)

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
        logger.info("Getting extra main table items with simplified matching...")

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
                logger.debug(f"Names don't match: PDF='{pdf_name}' vs Excel='{excel_name}'")
                return False
                
            # Compare quantities (with small tolerance)
            quantity_match = abs(pdf_item.quantity - excel_item.quantity) < 0.001
            
            if not quantity_match:
                logger.debug(f"Quantities don't match: PDF={pdf_item.quantity} vs Excel={excel_item.quantity}")
                return False
            
            # Compare units (normalize and compare) - FIXED to use unit-specific normalization
            pdf_unit = self._normalize_unit(pdf_item.unit)
            excel_unit = self._normalize_unit(excel_item.unit)
            unit_match = pdf_unit == excel_unit
            
            if not unit_match:
                logger.debug(f"Units don't match: PDF='{pdf_unit}' vs Excel='{excel_unit}'")
                return False
            
            return True

        # Find Excel items that don't have a matching PDF item
        extra_items = []
        
        for excel_item in excel_items:
            found_match = False
            
            for pdf_item in pdf_items:
                if items_match(pdf_item, excel_item):
                    found_match = True
                    logger.debug(f"Found match: Excel '{excel_item.item_key}' matches PDF '{pdf_item.item_key}'")
                    break
            
            if not found_match:
                extra_items.append(excel_item)
                logger.info(f"Extra item found: '{excel_item.item_key}' (qty: {excel_item.quantity}, unit: {excel_item.unit})")

        logger.info(f"Found {len(extra_items)} extra main table items using simplified matching")
        return extra_items
