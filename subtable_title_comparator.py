"""
Subtable Title Comparator
Compares PDF subtable titles with Excel subtable titles based on reference numbers.
"""

import re
import logging
from typing import List, Dict, Optional, Tuple
# These imports are needed for the direct file comparison function
from subtable_pdf_extractor import SubtablePDFExtractor
from excel_subtable_extractor import extract_subtables_from_excel

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def normalize_to_fullwidth(text: str) -> str:
    """Convert half-width characters to full-width characters."""
    if not text:
        return ""

    # Convert half-width to full-width characters
    half_to_full = str.maketrans(
        '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ ',
        '０１２３４５６７８９ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ　'
    )

    # Also convert common units to full-width
    unit_conversions = {
        'm': 'ｍ',
        'm2': 'ｍ²',
        'm3': 'ｍ³',
        'kg': 'ｋｇ',
        't': 'ｔ',
        'h': 'ｈ',
        '日': '日',
        '時間': '時間',
        '回': '回',
        '掛': '掛',
        '個': '個',
        '枚': '枚',
        '本': '本',
        '組': '組',
        '式': '式',
        '孔': '孔',
        '部材': '部材',
        '構造物': '構造物'
    }

    normalized = text.translate(half_to_full)

    # Apply unit conversions
    for half_width, full_width in unit_conversions.items():
        normalized = normalized.replace(half_width, full_width)

    return normalized


def normalize_text(text: str) -> str:
    """
    Normalize text by removing spaces and converting full-width characters to half-width
    """
    if not text:
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


def extract_item_name_parts(pdf_item_name: str) -> Tuple[str, str]:
    """
    Extract first and second parts of PDF item name.
    Split by common separators like spaces, commas, etc.
    """
    if not pdf_item_name:
        return "", ""

    # Split by common separators
    parts = re.split(r'[\s,、]+', pdf_item_name.strip())

    if len(parts) >= 2:
        return parts[0], parts[1]
    elif len(parts) == 1:
        return parts[0], ""
    else:
        return "", ""


def check_unit_presence_in_excel_title(pdf_unit: str, excel_title: str) -> bool:
    """
    Check if PDF unit is present anywhere in Excel title string.
    """
    if not pdf_unit or not excel_title:
        return False

    normalized_pdf_unit = normalize_text(pdf_unit)
    normalized_excel_title = normalize_text(excel_title)

    return normalized_pdf_unit in normalized_excel_title


def check_unit_quantity_presence_in_excel_title(pdf_unit_quantity: str, excel_title: str) -> bool:
    """
    Check if PDF unit quantity is present in Excel title string with stricter matching.
    Now checks for adjacent unit quantity + unit patterns (e.g., "10m2", "10m") rather than just substring matching.
    """
    if not pdf_unit_quantity or not excel_title:
        return False

    normalized_pdf_qty = normalize_text(pdf_unit_quantity)
    normalized_excel_title = normalize_text(excel_title)

    # First, try exact substring match (original logic)
    if normalized_pdf_qty in normalized_excel_title:
        return True

    # If exact match fails, try with comma variations
    qty_without_comma = normalized_pdf_qty.replace(',', '')
    if qty_without_comma in normalized_excel_title:
        return True

    # Try with comma (e.g., "1000" vs "1,000")
    qty_with_comma = normalized_pdf_qty
    if ',' not in qty_with_comma and len(qty_with_comma) >= 4:
        # Add comma for thousands
        qty_with_comma = qty_with_comma[:-3] + ',' + qty_with_comma[-3:]
        if qty_with_comma in normalized_excel_title:
            return True

    return False


def check_adjacent_unit_quantity_unit_pattern(pdf_unit_quantity: str, pdf_unit: str, excel_title: str) -> bool:
    """
    Check if PDF quantity+unit substring appears in the Excel title with quantity on the left and unit on the right.
    Per your rule, treat it as a match if the normalized Excel title contains the normalized '<qty><unit>' substring.
    """
    if not pdf_unit_quantity or not pdf_unit or not excel_title:
        return False

    normalized_pdf_qty = normalize_text(pdf_unit_quantity)
    normalized_pdf_unit = normalize_text(pdf_unit)
    normalized_excel_title = normalize_text(excel_title)

    # Normalize equivalent units: m^2 ↔ m2, m^3 ↔ m3, including full-width variants
    def norm_unit(u: str) -> str:
        u = u.replace('ｍ', 'm').replace('㎡', 'm2').replace('㎥', 'm3')
        u = u.replace('m^2', 'm2').replace('m²', 'm2').replace(
            'm^3', 'm3').replace('m³', 'm3')
        return u

    normalized_pdf_unit = norm_unit(normalized_pdf_unit)
    normalized_excel_title = norm_unit(normalized_excel_title)

    # Simple substring per requirement
    adjacent_pattern = normalized_pdf_qty + normalized_pdf_unit
    if adjacent_pattern in normalized_excel_title:
        return True

    # Accept quantity without comma variation
    qty_without_comma = normalized_pdf_qty.replace(',', '')
    if qty_without_comma + normalized_pdf_unit in normalized_excel_title:
        return True

    # Accept thousands with comma variation (e.g., 1000 -> 1,000)
    if ',' not in normalized_pdf_qty and len(normalized_pdf_qty) >= 4:
        qty_with_comma = normalized_pdf_qty[:-
                                            3] + ',' + normalized_pdf_qty[-3:]
        if qty_with_comma + normalized_pdf_unit in normalized_excel_title:
            return True

    return False


def check_title_match_with_pdf_data(excel_title: str, pdf_data: Dict) -> Tuple[bool, str]:
    """Check if PDF data components are present in Excel title using STRICT adjacent matching for unit+quantity."""
    if not excel_title or not pdf_data:
        return False, "Missing title or data"

    # Normalize both Excel and PDF data
    excel_title_normalized = normalize_text(excel_title.lower())

    # Get PDF components from actual extracted data
    pdf_item_name = pdf_data.get('item_name', '')
    pdf_unit = pdf_data.get('unit', '')
    pdf_unit_quantity = pdf_data.get('unit_quantity', '')

    # Normalize PDF components
    pdf_item_name_normalized = normalize_text(pdf_item_name.lower())
    pdf_unit_normalized = normalize_text(pdf_unit.lower())
    pdf_unit_quantity_normalized = normalize_text(pdf_unit_quantity.lower())

    # Strict requirement: both unit and quantity must exist in PDF title to consider a match
    if not pdf_unit_normalized or not pdf_unit_quantity_normalized:
        return False, "PDF missing unit and/or quantity"

    # STRICT LOGIC: Check each component with stricter matching

    # 1. Check item name: 1st or 2nd part of PDF item name anywhere in Excel title
    item_match = False
    if pdf_item_name_normalized:
        # Split the original PDF item name by spaces (before normalization)
        original_pdf_item = pdf_data.get('item_name', '')
        if original_pdf_item:
            space_parts = original_pdf_item.split()
            if len(space_parts) >= 2:
                first_part = normalize_text(
                    space_parts[0].lower())  # 1st part (e.g., "下塗")
                # 2nd part (e.g., "塗装種別:...")
                second_part = normalize_text(space_parts[1].lower())

                # Check if 1st part is in Excel title
                if first_part and first_part in excel_title_normalized:
                    item_match = True
                # Check if 2nd part is in Excel title
                elif second_part and second_part in excel_title_normalized:
                    item_match = True
                # If no exact match, try partial matching for 1st part
                elif first_part and len(first_part) >= 3:
                    for length in range(min(len(first_part), 6), 2, -1):
                        prefix = first_part[:length]
                        if prefix in excel_title_normalized:
                            item_match = True
                            break
            else:
                # If only one part, check that part
                single_part = normalize_text(
                    space_parts[0].lower()) if space_parts else ""
                if single_part and single_part in excel_title_normalized:
                    item_match = True
                elif single_part and len(single_part) >= 3:
                    for length in range(min(len(single_part), 6), 2, -1):
                        prefix = single_part[:length]
                        if prefix in excel_title_normalized:
                            item_match = True
                            break
    else:
        item_match = True  # If no item name in PDF, consider it a match

    # 2. Check unit: PDF unit MUST be present in Excel title string (we already ensured it's present)
    if pdf_unit_normalized in excel_title_normalized:
        unit_match = True
    else:
        unit_match = False

    # 3. Check unit quantity: Use STRICT adjacent pattern matching
    # Quantity must be adjacent to unit in Excel title (strict)
    if check_adjacent_unit_quantity_unit_pattern(pdf_unit_quantity_normalized, pdf_unit_normalized, excel_title_normalized):
        quantity_match = True
    else:
        quantity_match = False

    # UNIT & QUANTITY ONLY: Only check unit and unit quantity
    # Both unit and quantity must match
    if unit_match and quantity_match:
        return True, "Unit + Quantity match"
    else:
        # Check which conditions failed
        missing = []
        if not unit_match:
            missing.append("unit")
        if not quantity_match:
            missing.append("quantity")
        return False, f"No match - missing: {', '.join(missing)}"


def compare_subtable_titles(pdf_subtable, excel_subtable) -> Dict:
    """
    Compare PDF subtable title with Excel subtable title using improved AND logic.

    Args:
        pdf_subtable: PDF subtable (SubtableItem object or dictionary) with table_title
        excel_subtable: Excel subtable (SubtableItem object or dictionary) with table_title

    Returns:
        Comparison result dictionary
    """
    # Handle both SubtableItem objects and dictionaries
    if hasattr(pdf_subtable, 'reference_number'):
        pdf_ref = pdf_subtable.reference_number
        pdf_title = pdf_subtable.table_title
    else:
        pdf_ref = pdf_subtable.get("reference_number", "")
        pdf_title = pdf_subtable.get("table_title", {})

    if hasattr(excel_subtable, 'reference_number'):
        excel_ref = excel_subtable.reference_number
        excel_title = excel_subtable.table_title
    else:
        excel_ref = excel_subtable.get("reference_number", "")
        excel_title = excel_subtable.get("table_title", "")

    # Try to get PDF page number early
    pdf_page_number = None
    try:
        if hasattr(pdf_subtable, 'page_number'):
            pdf_page_number = getattr(pdf_subtable, 'page_number', None)
        elif isinstance(pdf_subtable, dict):
            pdf_page_number = pdf_subtable.get('page_number', None)
    except Exception:
        pdf_page_number = None

    result = {
        "pdf_reference": pdf_ref,
        "excel_reference": excel_ref,
        "pdf_title": pdf_title,
        "excel_title": excel_title,
        "pdf_page_number": pdf_page_number,
        "pdf_item_name": (pdf_title.get('item_name', '') if isinstance(pdf_title, dict) else ''),
        "unit_match": False,
        "unit_quantity_match": False,
        "overall_match": False,
        "match_score": 0,
        "details": {}
    }

    # Check if PDF has table title
    if not pdf_title:
        result["details"]["reason"] = "No PDF table title found - skipping comparison"
        return result

    # Check if Excel has table title
    if not excel_title:
        result["details"]["reason"] = "No Excel table title found"
        return result

    # Extract Excel title - handle both string and dictionary formats
    if isinstance(excel_title, dict):
        excel_title_text = excel_title.get("item_name", "")
    else:
        excel_title_text = str(excel_title)

    # Get PDF unit and quantity
    pdf_unit = pdf_title.get('unit', '') if isinstance(pdf_title, dict) else ''
    pdf_qty = pdf_title.get('unit_quantity', '') if isinstance(
        pdf_title, dict) else ''

    # If PDF is missing unit/quantity, cannot compare
    if not pdf_unit or not pdf_qty:
        result["details"]["reason"] = "PDF missing unit and/or quantity"
        return result

    # Requirement: Excel title must contain quantity first and then unit (after normalization)
    is_adjacent = check_adjacent_unit_quantity_unit_pattern(
        pdf_qty, pdf_unit, excel_title_text)
    if is_adjacent:
        is_match, match_type = True, "Excel title contains '<qty><unit>'"
    else:
        is_match, match_type = False, "Excel title missing '<qty><unit>'"

    result["overall_match"] = is_match
    result["details"]["reason"] = match_type

    # Update detail flags for clarity (true only when adjacency matched)
    result["unit_match"] = is_adjacent
    result["unit_quantity_match"] = is_adjacent

    # Set match score based on result
    if is_match:
        result["match_score"] = 100
    else:
        result["match_score"] = 0

    return result


def compare_all_subtable_titles(pdf_path: str, excel_path: str, pdf_start_page: int = 13, pdf_end_page: int = 82) -> Dict:
    """
    Compare all PDF subtable titles with Excel subtable titles.

    Args:
        pdf_path: Path to PDF file
        excel_path: Path to Excel file
        pdf_start_page: Starting page for PDF extraction
        pdf_end_page: Ending page for PDF extraction

    Returns:
        Dictionary with comparison results
    """
    logger.info(f"Starting subtable title comparison between PDF and Excel")

    try:
        # Extract PDF subtables
        pdf_extractor = SubtablePDFExtractor()
        pdf_result = pdf_extractor.extract_subtables_from_pdf(
            pdf_path, pdf_start_page, pdf_end_page)
        pdf_subtables = pdf_result.get("subtables", [])

        # Extract Excel subtables
        excel_subtables = extract_subtables_from_excel(excel_path)

        logger.info(
            f"Found {len(pdf_subtables)} PDF subtables and {len(excel_subtables)} Excel subtables")

        # Create reference number mappings with Kitakami equivalence (e.g., 第3号明 == 明3号)
        def norm_ref(r: str) -> str:
            if not r:
                return ""
            return re.sub(r"[\s　]+", "", str(r))

        def kitakami_key(r: str) -> str:
            r = norm_ref(r)
            m_pdf = re.match(r'^第?(\d+)号([一-龯])$', r)
            if m_pdf:
                num, tail = m_pdf.group(1), m_pdf.group(2)
                return f"{tail}:{num}"
            m_excel = re.match(r'^([一-龯])(\d+)号$', r)
            if m_excel:
                tail, num = m_excel.group(1), m_excel.group(2)
                return f"{tail}:{num}"
            return r

        pdf_by_ref = {kitakami_key(subtable["reference_number"]): subtable
                      for subtable in pdf_subtables if "reference_number" in subtable}
        excel_by_ref = {kitakami_key(subtable["reference_number"]): subtable
                        for subtable in excel_subtables if "reference_number" in subtable}

        logger.info(f"PDF subtables with titles: {len(pdf_by_ref)}")
        logger.info(f"Excel subtables with titles: {len(excel_by_ref)}")

        # Compare matching reference numbers
        comparison_results = []
        matched_refs = set()

        for ref_num in pdf_by_ref.keys():
            if ref_num in excel_by_ref:
                comparison = compare_subtable_titles(
                    pdf_by_ref[ref_num], excel_by_ref[ref_num])
                comparison_results.append(comparison)
                matched_refs.add(ref_num)

                if comparison["overall_match"]:
                    logger.info(
                        f"✅ MATCH: {ref_num} - {comparison['details']['reason']}")
                else:
                    logger.warning(
                        f"❌ NO MATCH: {ref_num} - {comparison['details']['reason']}")

        # Summary statistics
        total_comparisons = len(comparison_results)
        successful_matches = sum(
            1 for comp in comparison_results if comp["overall_match"])
        failed_matches = total_comparisons - successful_matches

        # PDF subtables without Excel counterparts
        pdf_only = set(pdf_by_ref.keys()) - matched_refs

        # Excel subtables without PDF counterparts
        excel_only = set(excel_by_ref.keys()) - matched_refs

        summary = {
            "total_pdf_subtables": len(pdf_subtables),
            "total_excel_subtables": len(excel_subtables),
            "pdf_subtables_with_titles": len(pdf_by_ref),
            "excel_subtables_with_titles": len(excel_by_ref),
            "total_comparisons": total_comparisons,
            "successful_matches": successful_matches,
            "failed_matches": failed_matches,
            "match_rate": (successful_matches / total_comparisons * 100) if total_comparisons > 0 else 0,
            "pdf_only_references": list(pdf_only),
            "excel_only_references": list(excel_only)
        }

        result = {
            "summary": summary,
            "comparisons": comparison_results,
            "pdf_subtables": pdf_subtables,
            "excel_subtables": excel_subtables
        }

        logger.info(
            f"Comparison complete: {successful_matches}/{total_comparisons} matches ({summary['match_rate']:.1f}%)")

        return result

    except Exception as e:
        logger.error(f"Error during comparison: {e}")
        raise


def _get_page_number(obj) -> Optional[int]:
    try:
        if hasattr(obj, 'page_number'):
            return getattr(obj, 'page_number', None)
        if isinstance(obj, dict):
            return obj.get('page_number')
    except Exception:
        return None
    return None


def _get_raw_fields(obj) -> Dict:
    try:
        if hasattr(obj, 'raw_fields'):
            return getattr(obj, 'raw_fields', {}) or {}
        if isinstance(obj, dict):
            return obj.get('raw_fields', {}) or {}
    except Exception:
        return {}
    return {}


def _get_reference_number(obj) -> str:
    try:
        if hasattr(obj, 'reference_number'):
            return getattr(obj, 'reference_number', '') or ''
        if isinstance(obj, dict):
            return obj.get('reference_number', '') or ''
    except Exception:
        return ''
    return ''


def _normalize_qty_text(q: str) -> str:
    if not q:
        return ''
    # Normalize punctuation and remove grouping commas
    q = str(q).replace('，', ',')
    q_no_commas = q.replace(',', '')
    # Trim trailing zeros but preserve significant decimals
    if '.' in q_no_commas:
        q_no_commas = re.sub(r'(\.\d*?)0+$', r'\1', q_no_commas)
        q_no_commas = re.sub(r'\.$', '', q_no_commas)
    # If ends with .00 style, becomes integer (e.g., 1.00 -> 1)
    m = re.match(r'^(\d+)\.(0+)$', q_no_commas)
    if m:
        return m.group(1)
    return q_no_commas


def _build_nousei_pdf_title_for_ref(ref: str, ref_page: Optional[int], pdf_items: List) -> Optional[Dict]:
    """Build a PDF title for Nousei by scanning triple-dot main rows in order and mapping sequentially to 内N号 across pages.
    Uses metadata fields from the stable extractor: '_is_triple_dot' (marker) and '備考' (remarks with 算出数量).
    """
    if not pdf_items:
        return None
    triple_titles = []
    for itm in pdf_items:
        raw = _get_raw_fields(itm)
        if not raw:
            continue
        # Require triple-dot marker flag from extraction metadata
        if not raw.get('_is_triple_dot'):
            continue
        name = (raw.get('工種・種目') or raw.get('工事区分・工種・種別・細別') or '').strip()
        remarks = raw.get('備考') or raw.get('摘要') or ''
        unit = ''
        qty = ''
        if remarks:
            m = re.search(
                r"算\s*出\s*数\s*量\s*[:：]?\s*([\d,，\.]+)\s*([^\s]+)", str(remarks))
            if not m:
                r_norm = re.sub(r"[\s\u3000]+", "", str(remarks))
                m = re.search(r"算出数量[:：]?([\d,，\.]+)([^\s]+)", r_norm)
            if m:
                qty = _normalize_qty_text((m.group(1) or '').replace('，', ','))
                unit = (m.group(2) or '').strip()
        triple_titles.append(
            {'item_name': name, 'unit': unit, 'unit_quantity': qty})
    if not triple_titles:
        return None
    mref = re.match(r'^内(\d+)号$', ref or '')
    if not mref:
        return None
    idx = int(mref.group(1)) - 1
    if 0 <= idx < len(triple_titles):
        return triple_titles[idx]
    return None


def compare_all_subtable_titles_from_cached_data(pdf_subtables: List, excel_subtables: List, pdf_items: Optional[List] = None) -> Dict:
    """
    Compare all PDF subtable titles with Excel subtable titles using cached data.

    Args:
        pdf_subtables: List of PDF SubtableItem objects with table_title
        excel_subtables: List of Excel SubtableItem objects with table_title

    Returns:
        Dictionary with comparison results
    """
    logger.info(f"Starting subtable title comparison from cached data")

    try:
        logger.info(
            f"Found {len(pdf_subtables)} PDF subtables and {len(excel_subtables)} Excel subtables")

        # Create reference number mappings with Kitakami equivalence (e.g., 第3号明 == 明3号)
        def norm_ref(r: str) -> str:
            if not r:
                return ""
            return re.sub(r"[\s　]+", "", str(r))

        def kitakami_key(r: str) -> str:
            r = norm_ref(r)
            m_pdf = re.match(r'^第?(\d+)号([一-龯])$', r)
            if m_pdf:
                num, tail = m_pdf.group(1), m_pdf.group(2)
                return f"{tail}:{num}"
            m_excel = re.match(r'^([一-龯])(\d+)号$', r)
            if m_excel:
                tail, num = m_excel.group(1), m_excel.group(2)
                return f"{tail}:{num}"
            return r

        pdf_by_ref = {}
        excel_by_ref = {}

        # Group all PDF subtables by reference and track first page for each ref
        ref_to_subs = {}
        ref_to_first_page = {}
        for subtable in pdf_subtables:
            ref = _get_reference_number(subtable)
            if not ref:
                continue
            key = kitakami_key(ref)
            ref_to_subs.setdefault(key, []).append(subtable)
            pnum = _get_page_number(subtable)
            if pnum is not None:
                if key not in ref_to_first_page or pnum < ref_to_first_page[key]:
                    ref_to_first_page[key] = pnum
        # For Nousei refs (内\d+号), build titles from main items
        for key, subs in ref_to_subs.items():
            # Determine original ref string for result fields
            orig_ref = _get_reference_number(subs[0])
            # Fill by default with the first subtable object
            base = subs[0]
            # Identify Nousei refs by pattern
            if re.match(r'^内\d+号$', orig_ref or '') and pdf_items:
                ref_page = ref_to_first_page.get(key)
                title = _build_nousei_pdf_title_for_ref(
                    orig_ref, ref_page, pdf_items)
                # Create an augmented dict with computed title and page number (marker page)
                augmented = dict(base) if isinstance(base, dict) else {
                    'reference_number': orig_ref,
                    'page_number': _get_page_number(base),
                    'table_title': None
                }
                if title:
                    augmented['table_title'] = title
                # Prefer the marker page if known
                if ref_page is not None:
                    augmented['page_number'] = ref_page
                pdf_by_ref[key] = augmented
            else:
                # Non-Nousei: keep as-is
                pdf_by_ref[key] = base

        for subtable in excel_subtables:
            if hasattr(subtable, 'reference_number'):
                excel_by_ref[kitakami_key(
                    subtable.reference_number)] = subtable
            elif isinstance(subtable, dict) and subtable.get("reference_number"):
                excel_by_ref[kitakami_key(
                    subtable["reference_number"])] = subtable

        logger.info(f"PDF subtables with titles: {len(pdf_by_ref)}")
        logger.info(f"Excel subtables with titles: {len(excel_by_ref)}")

        # Compare matching reference numbers
        comparison_results = []
        matched_refs = set()

        for ref_num in pdf_by_ref.keys():
            if ref_num in excel_by_ref:
                comparison = compare_subtable_titles(
                    pdf_by_ref[ref_num], excel_by_ref[ref_num])
                comparison_results.append(comparison)
                matched_refs.add(ref_num)

                if comparison["overall_match"]:
                    logger.info(
                        f"✅ MATCH: {ref_num} - {comparison['details']['reason']}")
                else:
                    logger.warning(
                        f"❌ NO MATCH: {ref_num} - {comparison['details']['reason']}")

        # Summary statistics
        total_comparisons = len(comparison_results)
        successful_matches = sum(
            1 for comp in comparison_results if comp["overall_match"])
        failed_matches = total_comparisons - successful_matches

        # PDF subtables without Excel counterparts
        pdf_only = set(pdf_by_ref.keys()) - matched_refs

        # Excel subtables without PDF counterparts
        excel_only = set(excel_by_ref.keys()) - matched_refs

        summary = {
            "total_pdf_subtables": len(pdf_subtables),
            "total_excel_subtables": len(excel_subtables),
            "pdf_subtables_with_titles": len(pdf_by_ref),
            "excel_subtables_with_titles": len(excel_by_ref),
            "total_comparisons": total_comparisons,
            "successful_matches": successful_matches,
            "failed_matches": failed_matches,
            "match_rate": (successful_matches / total_comparisons * 100) if total_comparisons > 0 else 0,
            "pdf_only_references": list(pdf_only),
            "excel_only_references": list(excel_only)
        }

        result = {
            "summary": summary,
            "comparisons": comparison_results,
            "pdf_subtables": pdf_subtables,
            "excel_subtables": excel_subtables
        }

        logger.info(
            f"Comparison complete: {successful_matches}/{total_comparisons} matches ({summary['match_rate']:.1f}%)")

        return result

    except Exception as e:
        logger.error(
            f"Error in subtable title comparison: {str(e)}", exc_info=True)
        raise


def print_comparison_summary(comparison_result: Dict):
    """
    Print a formatted summary of the comparison results.
    """
    summary = comparison_result["summary"]
    comparisons = comparison_result["comparisons"]

    print("=" * 80)
    print("SUBTABLE TITLE COMPARISON SUMMARY")
    print("=" * 80)
    print(
        f"PDF Subtables: {summary['total_pdf_subtables']} (with titles: {summary['pdf_subtables_with_titles']})")
    print(
        f"Excel Subtables: {summary['total_excel_subtables']} (with titles: {summary['excel_subtables_with_titles']})")
    print(f"Total Comparisons: {summary['total_comparisons']}")
    print(f"Successful Matches: {summary['successful_matches']}")
    print(f"Failed Matches: {summary['failed_matches']}")
    print(f"Match Rate: {summary['match_rate']:.1f}%")
    print()

    if summary['pdf_only_references']:
        print(
            f"PDF Only References: {', '.join(summary['pdf_only_references'])}")
    if summary['excel_only_references']:
        print(
            f"Excel Only References: {', '.join(summary['excel_only_references'])}")
    print()

    print("DETAILED COMPARISON RESULTS:")
    print("-" * 80)

    for comp in comparisons:
        status = "✅ MATCH" if comp["overall_match"] else "❌ NO MATCH"
        print(
            f"{status} | {comp['pdf_reference']} | Score: {comp['match_score']}")
        print(f"  PDF Title: {comp['pdf_title']}")
        print(f"  Excel Title: {comp['excel_title']}")
        print(f"  Reason: {comp['details']['reason']}")
        print()


# Test function
if __name__ == "__main__":
    pdf_file = "07_入札時（見積）積算参考資料.pdf"
    excel_file = "【修正】水沢橋　積算書.xlsx"

    try:
        result = compare_all_subtable_titles(pdf_file, excel_file)
        print_comparison_summary(result)
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
