import os
import sys
import json
from io import BytesIO


def main():
    # Resolve project root and add backend to path
    script_dir = os.path.dirname(os.path.abspath(__file__))
    backend_dir = os.path.abspath(os.path.join(script_dir, '..'))
    sys.path.insert(0, backend_dir)

    # Import after sys.path setup
    from server.services.pdf_parser import PDFParser
    from server.services.excel_table_extractor_service import ExcelTableExtractorService
    from server.services.matcher import Matcher
    from server.services.normalizer import Normalizer

    # Inputs (hardcoded per user request)
    project_root = os.path.abspath(os.path.join(backend_dir, '..'))
    pdf_path = os.path.join(project_root, "07+入札時（見積）積算参考資料 (2).pdf")
    excel_path = os.path.join(project_root, "【52標準】更木地区_単19号3パターン.xlsx")
    main_start, main_end = 3, 7
    sub_start, sub_end = 8, 25
    main_sheet = "52標準 15行本工事内訳書"

    if not os.path.exists(pdf_path):
        print(json.dumps(
            {"error": f"PDF not found: {pdf_path}"}, ensure_ascii=False))
        return
    if not os.path.exists(excel_path):
        print(json.dumps(
            {"error": f"Excel not found: {excel_path}"}, ensure_ascii=False))
        return

    # Extract PDF main and subtable items
    pdf_parser = PDFParser()
    pdf_main_items = pdf_parser.extract_tables_with_range(
        pdf_path, main_start, main_end)
    pdf_subtable_items = pdf_parser.extract_subtables_with_range(
        pdf_path, sub_start, sub_end)

    # Extract Excel main and subtable items
    with open(excel_path, 'rb') as f:
        excel_bytes = f.read()
    excel_buffer = BytesIO(excel_bytes)

    excel_service = ExcelTableExtractorService()
    excel_main_items = excel_service.extract_main_table_from_buffer(
        excel_buffer, main_sheet)

    # For subtables, the service expects buffer + main items
    # Create a fresh buffer since previous was consumed by openpyxl
    excel_buffer2 = BytesIO(excel_bytes)
    excel_subtable_items = excel_service.extract_subtables_from_buffer(
        excel_buffer2, main_sheet, excel_main_items)

    # Compute name mismatches using same categorization as API
    matcher = Matcher()
    normalizer = Normalizer()

    # Main table: use index-based compare, then collect NAME_MISMATCH and classify as Cat-2/Cat-3
    main_summary = matcher.compare_items(pdf_main_items, excel_main_items)
    main_name_mismatches = []
    for r in (main_summary.results or []):
        if r.status == 'NAME_MISMATCH' and r.pdf_item is not None and r.excel_item is not None:
            pdf_tokens = [t for t in normalizer.tokenize_item_name(
                r.pdf_item.item_key)]
            excel_name = normalizer.normalize_item(r.excel_item.item_key)
            is_cat2 = all(t in excel_name for t in pdf_tokens if t)
            is_cat3 = any((t and t in excel_name)
                          for t in pdf_tokens) and not is_cat2
            if is_cat2 or is_cat3:
                main_name_mismatches.append({
                    "pdf_item_name": r.pdf_item.item_key,
                    "excel_item_name": r.excel_item.item_key,
                    "category": 2 if is_cat2 else 3,
                    "pdf_item": {
                        "item_key": r.pdf_item.item_key,
                        "page_number": getattr(r.pdf_item, 'page_number', None)
                    },
                    "excel_item": {
                        "item_key": r.excel_item.item_key,
                        "table_number": getattr(r.excel_item, 'table_number', None),
                        "logical_line_number": getattr(r.excel_item, 'logical_line_number', None)
                    },
                    "type": "Main Table"
                })

    # Subtables: use category-based comparator and filter NAME_MISMATCH
    subtable_results = matcher.compare_subtable_items(
        pdf_subtable_items, excel_subtable_items)
    subtable_name_mismatches = []
    for r in subtable_results:
        if r.status == 'NAME_MISMATCH' and r.pdf_item is not None and r.excel_item is not None:
            pdf_tokens = [t for t in normalizer.tokenize_item_name(
                r.pdf_item.item_key)]
            excel_name = normalizer.normalize_item(r.excel_item.item_key)
            is_cat2 = all(t in excel_name for t in pdf_tokens if t)
            is_cat3 = any((t and t in excel_name)
                          for t in pdf_tokens) and not is_cat2
            if is_cat2 or is_cat3:
                subtable_name_mismatches.append({
                    "pdf_item_name": r.pdf_item.item_key,
                    "excel_item_name": r.excel_item.item_key,
                    "category": 2 if is_cat2 else 3,
                    "pdf_item": {
                        "item_key": r.pdf_item.item_key,
                        "page_number": getattr(r.pdf_item, 'page_number', None),
                        "reference_number": getattr(r.pdf_item, 'reference_number', None)
                    },
                    "excel_item": {
                        "item_key": r.excel_item.item_key,
                        "reference_number": getattr(r.excel_item, 'reference_number', None),
                        "logical_line_number": getattr(r.excel_item, 'logical_line_number', None)
                    },
                    "type": "Sub Table"
                })

    result = {
        "params": {
            "pdf_main_pages": [main_start, main_end],
            "pdf_subtable_pages": [sub_start, sub_end],
            "excel_sheet": main_sheet,
            "pdf": pdf_path,
            "excel": excel_path
        },
        "main_name_mismatches_count": len(main_name_mismatches),
        "subtable_name_mismatches_count": len(subtable_name_mismatches),
        "main_name_mismatches": main_name_mismatches,
        "subtable_name_mismatches": subtable_name_mismatches,
    }

    # Write to UTF-8 JSON file to avoid console encoding issues on Windows
    out_path = os.path.join(project_root, "name_mismatches_result.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    # Print ASCII-only confirmation
    print(f"WROTE: {out_path}")


if __name__ == "__main__":
    main()
