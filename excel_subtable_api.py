"""
API-Ready Excel Subtable Extractor
Extracts all subtables from all remaining sheets (except main sheet) of an Excel file
"""

import pandas as pd
import logging
import sys
import os
from typing import List, Dict, Any, Optional

# Add parent directory to path to import excel_subtable_extractor
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, parent_dir)
from excel_subtable_extractor import extract_subtables_from_excel

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def extract_all_subtables_api(excel_file_path: str) -> Dict[str, Any]:
    """
    API-ready function to extract all subtables from all remaining sheets (except main sheet)

    Args:
        excel_file_path (str): Path to the Excel file

    Returns:
        Dict[str, Any]: API response containing:
            - success (bool): Whether extraction was successful
            - message (str): Success/error message
            - total_sheets_processed (int): Number of sheets processed
            - total_subtables (int): Total number of subtables found
            - total_data_rows (int): Total number of data rows across all subtables
            - sheets (List[Dict]): Detailed results for each sheet
            - all_subtables (List[Dict]): All subtables combined from all sheets
            - reference_patterns (Dict): Statistics of reference number patterns found
    """

    try:
        # Load Excel file and get sheet names
        xl_file = pd.ExcelFile(excel_file_path)
        all_sheet_names = xl_file.sheet_names

        if len(all_sheet_names) < 2:
            return {
                "success": False,
                "message": "Excel file must contain at least 2 sheets (main sheet + remaining sheets)",
                "total_sheets_processed": 0,
                "total_subtables": 0,
                "total_data_rows": 0,
                "sheets": [],
                "all_subtables": [],
                "reference_patterns": {}
            }

        # Skip the main sheet (first sheet) and process remaining sheets
        main_sheet = all_sheet_names[0]
        remaining_sheets = all_sheet_names[1:]

        logger.info(f"Processing Excel file: {excel_file_path}")
        logger.info(f"Main sheet (skipped): {main_sheet}")
        logger.info(f"Remaining sheets to process: {remaining_sheets}")

        # Initialize results
        all_sheets_results = []
        all_subtables_combined = []
        total_subtables = 0
        total_data_rows = 0
        reference_patterns = {}

        # Process each remaining sheet
        for sheet_index, sheet_name in enumerate(remaining_sheets, 1):
            try:
                logger.info(
                    f"Processing sheet {sheet_index}/{len(remaining_sheets)}: {sheet_name}")

                # Extract subtables from current sheet
                subtables = extract_subtables_from_excel(
                    excel_file_path, sheet_name)

                # Calculate sheet statistics
                sheet_subtables = len(subtables)
                sheet_data_rows = sum(st['total_rows'] for st in subtables)

                # Analyze reference patterns for this sheet
                sheet_patterns = {}
                for subtable in subtables:
                    ref_num = subtable['reference_number']
                    if ref_num.startswith('内'):
                        pattern = '内X号'
                    elif ref_num.startswith('単'):
                        pattern = '単X号'
                    elif ref_num.startswith('代'):
                        pattern = '代X号'
                    elif ref_num.startswith('施'):
                        pattern = '施X号'
                    else:
                        pattern = 'Other'

                    sheet_patterns[pattern] = sheet_patterns.get(
                        pattern, 0) + 1
                    reference_patterns[pattern] = reference_patterns.get(
                        pattern, 0) + 1

                # Add sheet metadata to each subtable
                for subtable in subtables:
                    subtable['sheet_name'] = sheet_name
                    subtable['sheet_index'] = sheet_index

                # Store sheet results
                sheet_result = {
                    "sheet_name": sheet_name,
                    "sheet_index": sheet_index,
                    "subtables_count": sheet_subtables,
                    "data_rows_count": sheet_data_rows,
                    "reference_patterns": sheet_patterns,
                    "subtables": subtables,
                    "success": True,
                    "message": f"Successfully extracted {sheet_subtables} subtables with {sheet_data_rows} data rows"
                }

                all_sheets_results.append(sheet_result)
                all_subtables_combined.extend(subtables)
                total_subtables += sheet_subtables
                total_data_rows += sheet_data_rows

                logger.info(
                    f"Sheet '{sheet_name}': {sheet_subtables} subtables, {sheet_data_rows} data rows")

            except Exception as sheet_error:
                logger.error(
                    f"Error processing sheet '{sheet_name}': {sheet_error}")

                # Add failed sheet result
                sheet_result = {
                    "sheet_name": sheet_name,
                    "sheet_index": sheet_index,
                    "subtables_count": 0,
                    "data_rows_count": 0,
                    "reference_patterns": {},
                    "subtables": [],
                    "success": False,
                    "message": f"Error processing sheet: {str(sheet_error)}"
                }
                all_sheets_results.append(sheet_result)

        # Create final API response
        api_response = {
            "success": True,
            "message": f"Successfully processed {len(remaining_sheets)} sheets, extracted {total_subtables} subtables with {total_data_rows} total data rows",
            "excel_file": excel_file_path,
            "main_sheet_skipped": main_sheet,
            "total_sheets_processed": len(remaining_sheets),
            "total_subtables": total_subtables,
            "total_data_rows": total_data_rows,
            "reference_patterns": reference_patterns,
            "sheets": all_sheets_results,
            "all_subtables": all_subtables_combined
        }

        logger.info(
            f"API extraction complete: {total_subtables} subtables from {len(remaining_sheets)} sheets")
        return api_response

    except Exception as e:
        error_message = f"Failed to process Excel file: {str(e)}"
        logger.error(error_message)

        return {
            "success": False,
            "message": error_message,
            "excel_file": excel_file_path,
            "total_sheets_processed": 0,
            "total_subtables": 0,
            "total_data_rows": 0,
            "sheets": [],
            "all_subtables": [],
            "reference_patterns": {}
        }


def get_subtables_summary(api_response: Dict[str, Any]) -> Dict[str, Any]:
    """
    Get a concise summary of the extraction results

    Args:
        api_response (Dict): Response from extract_all_subtables_api

    Returns:
        Dict[str, Any]: Summary statistics
    """

    if not api_response.get("success", False):
        return {
            "success": False,
            "message": api_response.get("message", "Unknown error")
        }

    # Calculate per-sheet summary
    sheet_summaries = []
    for sheet in api_response.get("sheets", []):
        if sheet.get("success", False):
            sheet_summaries.append({
                "sheet_name": sheet["sheet_name"],
                "subtables": sheet["subtables_count"],
                "data_rows": sheet["data_rows_count"],
                "patterns": sheet["reference_patterns"]
            })

    return {
        "success": True,
        "excel_file": api_response.get("excel_file", ""),
        "total_sheets": api_response.get("total_sheets_processed", 0),
        "total_subtables": api_response.get("total_subtables", 0),
        "total_data_rows": api_response.get("total_data_rows", 0),
        "reference_patterns": api_response.get("reference_patterns", {}),
        "sheet_summaries": sheet_summaries
    }


# Example usage and testing
if __name__ == "__main__":
    # Test the API function
    excel_file = "【修正】水沢橋　積算書.xlsx"

    print("🚀 TESTING API-READY SUBTABLE EXTRACTOR")
    print("=" * 80)

    # Extract all subtables from all remaining sheets
    result = extract_all_subtables_api(excel_file)

    if result["success"]:
        print(f"✅ SUCCESS: {result['message']}")
        print(f"📁 Excel file: {result['excel_file']}")
        print(f"📄 Main sheet skipped: {result['main_sheet_skipped']}")
        print(f"📊 Sheets processed: {result['total_sheets_processed']}")
        print(f"📋 Total subtables: {result['total_subtables']}")
        print(f"📝 Total data rows: {result['total_data_rows']}")

        print(f"\n📊 Reference patterns found:")
        for pattern, count in result['reference_patterns'].items():
            print(f"  - {pattern}: {count} subtables")

        print(f"\n📄 Per-sheet breakdown:")
        for sheet in result['sheets']:
            status = "✅" if sheet['success'] else "❌"
            print(
                f"  {status} {sheet['sheet_name']}: {sheet['subtables_count']} subtables, {sheet['data_rows_count']} rows")

        # Get summary
        summary = get_subtables_summary(result)
        print(f"\n📋 SUMMARY:")
        print(
            f"Total extraction: {summary['total_subtables']} subtables from {summary['total_sheets']} sheets")

    else:
        print(f"❌ FAILED: {result['message']}")

    print("=" * 80)
