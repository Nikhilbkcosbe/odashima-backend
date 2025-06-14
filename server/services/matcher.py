from typing import List, Dict, Tuple
from rapidfuzz import process, fuzz
from ..schemas.tender import TenderItem, ComparisonResult, ComparisonSummary
from .normalizer import Normalizer


class Matcher:
    def __init__(self):
        self.normalizer = Normalizer()
        self.min_confidence = 0.8

    def compare_items(self, pdf_items: List[TenderItem], excel_items: List[TenderItem]) -> ComparisonSummary:
        """
        Compare items between PDF and Excel, returning a detailed comparison summary.
        """
        # Normalize all items
        pdf_normalized = {self.normalizer.normalize_item(
            item.item_key): item for item in pdf_items}
        excel_normalized = {self.normalizer.normalize_item(
            item.item_key): item for item in excel_items}

        results = []
        matched_excel_keys = set()

        # Compare each PDF item
        for pdf_key, pdf_item in pdf_normalized.items():
            # Try exact match first
            if pdf_key in excel_normalized:
                excel_item = excel_normalized[pdf_key]
                matched_excel_keys.add(pdf_key)

                # Check quantity
                if abs(pdf_item.quantity - excel_item.quantity) < 0.001:
                    status = "OK"
                    quantity_diff = None
                else:
                    status = "QUANTITY_MISMATCH"
                    quantity_diff = excel_item.quantity - pdf_item.quantity

                results.append(ComparisonResult(
                    status=status,
                    pdf_item=pdf_item,
                    excel_item=excel_item,
                    match_confidence=1.0,
                    quantity_difference=quantity_diff
                ))
                continue

            # Try fuzzy match
            matches = process.extractOne(
                pdf_key,
                excel_normalized.keys(),
                scorer=fuzz.ratio,
                score_cutoff=self.min_confidence * 100
            )

            if matches:
                # Handle different return formats from rapidfuzz
                if len(matches) >= 2:
                    excel_key = matches[0]
                    confidence = matches[1]
                else:
                    excel_key = matches[0]
                    confidence = 100  # Default to high confidence for exact matches
                
                excel_item = excel_normalized[excel_key]
                matched_excel_keys.add(excel_key)

                # Check quantity
                if abs(pdf_item.quantity - excel_item.quantity) < 0.001:
                    status = "OK"
                    quantity_diff = None
                else:
                    status = "QUANTITY_MISMATCH"
                    quantity_diff = excel_item.quantity - pdf_item.quantity

                results.append(ComparisonResult(
                    status=status,
                    pdf_item=pdf_item,
                    excel_item=excel_item,
                    match_confidence=confidence / 100,
                    quantity_difference=quantity_diff
                ))
            else:
                # No match found
                results.append(ComparisonResult(
                    status="MISSING",
                    pdf_item=pdf_item,
                    excel_item=None,
                    match_confidence=0.0,
                    quantity_difference=None
                ))

        # Find extra items in Excel
        for excel_key, excel_item in excel_normalized.items():
            if excel_key not in matched_excel_keys:
                results.append(ComparisonResult(
                    status="EXTRA",
                    pdf_item=None,
                    excel_item=excel_item,
                    match_confidence=0.0,
                    quantity_difference=None
                ))

        # Calculate summary
        total_items = len(results)
        matched_items = sum(1 for r in results if r.status == "OK")
        quantity_mismatches = sum(
            1 for r in results if r.status == "QUANTITY_MISMATCH")
        missing_items = sum(1 for r in results if r.status == "MISSING")
        extra_items = sum(1 for r in results if r.status == "EXTRA")

        return ComparisonSummary(
            total_items=total_items,
            matched_items=matched_items,
            quantity_mismatches=quantity_mismatches,
            missing_items=missing_items,
            extra_items=extra_items,
            results=results
        )
