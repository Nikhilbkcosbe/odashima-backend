from fastapi import APIRouter, UploadFile, File, HTTPException, Form
from typing import List, Optional
import tempfile
import os
import gc
import logging
import time
from io import BytesIO
from ..schemas.tender import ComparisonSummary, SubtableComparisonSummary
from ..services.pdf_parser import PDFParser
from ..services.excel_parser import ExcelParser
from ..services.matcher import Matcher
from ..services.excel_table_extractor_service import ExcelTableExtractorService
from ..services.normalizer import Normalizer
from ..services.extraction_cache_service import get_extraction_cache

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

router = APIRouter()


@router.get("/test")
async def test_endpoint():
    """Simple test endpoint to verify the server is working"""
    return {"status": "ok", "message": "Tender API is working"}


@router.post("/test-new-extraction")
async def test_new_extraction_endpoint(excel_file: UploadFile = File(...)):
    """
    NEW: Test endpoint to verify the new API-ready subtable extraction is working
    """
    logger.info("=== TESTING NEW API-READY SUBTABLE EXTRACTION ===")
    logger.info(f"Excel file: {excel_file.filename}")

    # Validate file type
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Excel file required")

    try:
        # Handle Excel file (in-memory processing)
        excel_content = await excel_file.read()
        logger.info(f"Excel file size: {len(excel_content)} bytes")

        # Create in-memory buffer for Excel
        excel_buffer = BytesIO(excel_content)

        # Initialize the Excel table extractor service
        excel_table_extractor = ExcelTableExtractorService()

        # Test the new API-ready subtable extraction
        logger.info("Testing new API-ready subtable extraction...")
        excel_subtables = excel_table_extractor.extract_subtables_with_new_api(excel_content)
        logger.info(f"NEW API extracted {len(excel_subtables)} Excel subtable items")

        # Close Excel buffer
        excel_buffer.close()

        # TEST: Unit normalization fix
        logger.info("=== TESTING UNIT NORMALIZATION FIX ===")
        from ..services.matcher import Matcher
        matcher = Matcher()
        
        # Test cases for the reported unit mismatch issue
        test_cases = [
            ("m", "ｍ"),      # half-width vs full-width m
            ("t", "ｔ"),      # half-width vs full-width t  
            ("kg", "ｋｇ"),   # half-width vs full-width kg
            ("m", "m"),      # identical units
            ("t", "t"),      # identical units
        ]
        
        unit_test_results = []
        for pdf_unit, excel_unit in test_cases:
            pdf_normalized = matcher._normalize_unit(pdf_unit)
            excel_normalized = matcher._normalize_unit(excel_unit)
            match = pdf_normalized == excel_normalized
            
            test_result = {
                "pdf_unit": pdf_unit,
                "excel_unit": excel_unit,
                "pdf_normalized": pdf_normalized,
                "excel_normalized": excel_normalized,
                "units_match": match,
                "issue_type": "FIXED" if match else "STILL_MISMATCH"
            }
            unit_test_results.append(test_result)
            logger.info(f"Unit test: '{pdf_unit}' vs '{excel_unit}' -> {pdf_normalized} vs {excel_normalized} = {'MATCH' if match else 'MISMATCH'}")

        # Create summary response
        summary = {
            "success": True,
            "message": f"Successfully tested new extraction with {len(excel_subtables)} subtable items",
            "excel_file": excel_file.filename,
            "total_subtable_items": len(excel_subtables),
            "extraction_method": "NEW API-ready extraction",
            "unit_normalization_tests": unit_test_results,
            "sample_items": []
        }

        # Add sample items (first 3) to the response for verification
        for i, item in enumerate(excel_subtables[:3]):
            sample_item = {
                "item_key": item.item_key,
                "reference_number": item.reference_number,
                "sheet_name": item.sheet_name,
                "quantity": item.quantity,
                "unit": item.unit,
                "raw_fields": item.raw_fields
            }
            summary["sample_items"].append(sample_item)

        logger.info("=== NEW EXTRACTION TEST COMPLETED ===")
        return summary

    except Exception as e:
        logger.error(f"Error during new extraction test: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Test extraction error: {str(e)}")

    finally:
        gc.collect()


@router.post("/extract-and-cache")
async def extract_and_cache_files(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    start_page: Optional[int] = Form(None),
    end_page: Optional[int] = Form(None),
    sheet_name: Optional[str] = Form(None),
    pdf_subtable_start_page: Optional[int] = Form(None),
    pdf_subtable_end_page: Optional[int] = Form(None)
):
    """
    OPTIMIZED EXTRACTION ENDPOINT: Extract PDF and Excel data once and cache for reuse.
    This eliminates redundant extraction across multiple comparison operations.

    Args:
        pdf_file: PDF tender document
        excel_file: Excel proposal document  
        start_page: Starting page number for PDF main table extraction
        end_page: Ending page number for PDF main table extraction
        sheet_name: Specific Excel sheet name for main table extraction
        pdf_subtable_start_page: Starting page number for PDF subtable extraction
        pdf_subtable_end_page: Ending page number for PDF subtable extraction

    Returns:
        session_id: Unique identifier for cached extraction results
        extraction_summary: Summary of extracted data
    """
    logger.info("=== STARTING OPTIMIZED EXTRACTION AND CACHING ===")
    logger.info(f"PDF file: {pdf_file.filename}")
    logger.info(f"Excel file: {excel_file.filename}")
    logger.info(f"PDF main table page range: {start_page} to {end_page}")
    logger.info(
        f"PDF subtable page range: {pdf_subtable_start_page} to {pdf_subtable_end_page}")
    logger.info(f"Excel sheet: {sheet_name or 'Auto-detect'}")

    # Validate file types
    if not pdf_file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Excel file required")

    # Validate page ranges
    if start_page is not None and start_page < 1:
        raise HTTPException(status_code=400, detail="Start page must be >= 1")
    if end_page is not None and end_page < 1:
        raise HTTPException(status_code=400, detail="End page must be >= 1")
    if start_page is not None and end_page is not None and start_page > end_page:
        raise HTTPException(
            status_code=400, detail="Start page cannot be greater than end page")

    if pdf_subtable_start_page is not None and pdf_subtable_start_page < 1:
        raise HTTPException(
            status_code=400, detail="Subtable start page must be >= 1")
    if pdf_subtable_end_page is not None and pdf_subtable_end_page < 1:
        raise HTTPException(
            status_code=400, detail="Subtable end page must be >= 1")
    if pdf_subtable_start_page is not None and pdf_subtable_end_page is not None and pdf_subtable_start_page > pdf_subtable_end_page:
        raise HTTPException(
            status_code=400, detail="Subtable start page cannot be greater than end page")

    pdf_fd = None
    pdf_path = None

    try:
        logger.info("=== PROCESSING FILES FOR EXTRACTION ===")

        # Handle PDF file (requires temporary file for pdfplumber)
        pdf_fd, pdf_path = tempfile.mkstemp(suffix='.pdf', prefix='optimized_')
        pdf_content = await pdf_file.read()
        logger.info(f"PDF file size: {len(pdf_content)} bytes")

        with os.fdopen(pdf_fd, 'wb') as f:
            f.write(pdf_content)
        pdf_fd = None  # File descriptor is closed now

        # Handle Excel file (in-memory processing)
        excel_content = await excel_file.read()
        logger.info(f"Excel file size: {len(excel_content)} bytes")
        excel_buffer = BytesIO(excel_content)

        logger.info("=== STARTING COMPREHENSIVE EXTRACTION ===")

        # Initialize parsers
        pdf_parser = PDFParser()
        excel_table_extractor = ExcelTableExtractorService()

        # Extract PDF main table items
        logger.info("Extracting PDF main table items...")
        pdf_items = pdf_parser.extract_tables_with_range(
            pdf_path, start_page, end_page)
        logger.info(f"Extracted {len(pdf_items)} PDF main table items")

        # Extract Excel main table items
        logger.info("Extracting Excel main table items...")
        if sheet_name:
            excel_items = excel_table_extractor.extract_main_table_from_buffer(
                excel_buffer, sheet_name)
        else:
            # Use original parser for all sheets
            excel_parser = ExcelParser()
            excel_items = excel_parser.extract_items_from_buffer_with_sheet(
                excel_buffer, sheet_name)
        logger.info(f"Extracted {len(excel_items)} Excel main table items")

        # Extract subtables from both PDF and Excel
        logger.info("Extracting subtables from PDF...")
        pdf_subtables = pdf_parser.extract_subtables_with_range(
            pdf_path, pdf_subtable_start_page, pdf_subtable_end_page
        )
        logger.info(f"Extracted {len(pdf_subtables)} PDF subtable items")

        logger.info("Extracting subtables from Excel using NEW API-ready extraction...")
        excel_subtables = excel_table_extractor.extract_subtables_from_buffer(
            excel_buffer, sheet_name or '', excel_items
        )
        logger.info(f"NEW API extracted {len(excel_subtables)} Excel subtable items")

        # Close Excel buffer
        excel_buffer.close()

        # Store extraction parameters for reference
        extraction_params = {
            'pdf_filename': pdf_file.filename,
            'excel_filename': excel_file.filename,
            'start_page': start_page,
            'end_page': end_page,
            'sheet_name': sheet_name,
            'pdf_subtable_start_page': pdf_subtable_start_page,
            'pdf_subtable_end_page': pdf_subtable_end_page,
            'extracted_at': time.time()
        }

        # Cache the extraction results
        cache_service = get_extraction_cache()
        session_id = cache_service.store_extraction_results(
            pdf_items=pdf_items,
            excel_items=excel_items,
            pdf_subtables=pdf_subtables,
            excel_subtables=excel_subtables,
            extraction_params=extraction_params
        )

        logger.info("=== EXTRACTION COMPLETED AND CACHED ===")
        logger.info(f"Session ID: {session_id}")

        # Store counts before cleanup
        pdf_main_count = len(pdf_items)
        excel_main_count = len(excel_items)
        pdf_subtable_count = len(pdf_subtables)
        excel_subtable_count = len(excel_subtables)
        total_count = pdf_main_count + excel_main_count + \
            pdf_subtable_count + excel_subtable_count

        # Clean up parser references
        del pdf_parser, excel_table_extractor, pdf_items, excel_items, pdf_subtables, excel_subtables
        gc.collect()

        return {
            "session_id": session_id,
            "extraction_summary": {
                "pdf_main_items": pdf_main_count,
                "excel_main_items": excel_main_count,
                "pdf_subtable_items": pdf_subtable_count,
                "excel_subtable_items": excel_subtable_count,
                "total_items_extracted": total_count,
                "extraction_parameters": extraction_params,
                "cache_expires_in_minutes": 30
            }
        }

    except Exception as e:
        logger.error(
            f"Error during optimized extraction: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Extraction error: {str(e)}")

    finally:
        # Cleanup temporary PDF file
        try:
            if pdf_fd is not None:
                os.close(pdf_fd)
        except Exception:
            pass

        if pdf_path and os.path.exists(pdf_path):
            try:
                os.unlink(pdf_path)
            except Exception:
                pass

        gc.collect()


@router.post("/compare-cached-extra-items")
async def compare_cached_extra_items(session_id: str = Form(...)):
    """
    OPTIMIZED COMPARISON: Get both extra items (Excel not in PDF) and missing items (PDF not in Excel)
    using cached extraction results. This is much faster as it skips the extraction phase entirely.
    
    Args:
        session_id: Session identifier from extract-and-cache endpoint
        
    Returns:
        Complete comparison results including both extra and missing items
    """
    logger.info(
        f"=== COMPARING EXTRA & MISSING ITEMS FROM CACHED SESSION {session_id} ===")

    try:
        # Get cached extraction results
        cache_service = get_extraction_cache()
        cached_data = cache_service.get_extraction_results(session_id)

        if not cached_data:
            raise HTTPException(
                status_code=404,
                detail="Session not found or expired. Please re-extract files first."
            )

        # Extract cached data
        pdf_items = cached_data['pdf_items']
        excel_items = cached_data['excel_items']
        pdf_subtables = cached_data['pdf_subtables']
        excel_subtables = cached_data['excel_subtables']

        logger.info(f"Using cached data: {len(pdf_items)} PDF items, {len(excel_items)} Excel items, "
                    f"{len(pdf_subtables)} PDF subtables, {len(excel_subtables)} Excel subtables")

        # Perform comparison using cached data
        matcher = Matcher()

        # Main table extra items (Excel items not in PDF - simplified matching)
        extra_main_items = matcher.get_extra_items_only_simplified(
            pdf_items, excel_items)

        # Main table missing items (PDF items not in Excel - simplified matching)
        missing_main_items = matcher.get_extra_items_only_simplified(
            excel_items, pdf_items)

        # Subtable extra items (Excel subtables not in PDF)
        extra_subtable_items = matcher.get_extra_subtable_items_only(
            pdf_subtables, excel_subtables)

        # Subtable missing items (PDF subtables not in Excel)
        missing_subtable_items = matcher.get_extra_subtable_items_only(
            excel_subtables, pdf_subtables)
        
        # QUANTITY AND UNIT MISMATCH ANALYSIS
        # Get quantity mismatches for main table items
        main_quantity_mismatches = matcher.get_mismatched_items_only(pdf_items, excel_items)
        
        # Get unit mismatches for main table items
        main_unit_mismatches = matcher.get_unit_mismatched_items_only(pdf_items, excel_items)
        
        # For subtables, we need to use the comprehensive comparison
        subtable_results = matcher.compare_subtable_items(pdf_subtables, excel_subtables)
        subtable_quantity_mismatches = [r for r in subtable_results if r.status == 'QUANTITY_MISMATCH']
        subtable_unit_mismatches = [r for r in subtable_results if r.status == 'UNIT_MISMATCH']

        # Combine EXTRA items (Excel items not in PDF) for frontend display
        combined_extra_items = []

        # Add main table extra items with type indicator
        for item in extra_main_items:
            combined_extra_items.append({
                "item_key": item.item_key,
                "raw_fields": item.raw_fields,
                "quantity": item.quantity,
                "unit": item.unit,
                "source": item.source,
                "page_number": item.page_number,
                "type": "Main Table"
            })

        # Add subtable extra items with type indicator
        for item in extra_subtable_items:
            combined_extra_items.append({
                "item_key": item.item_key,
                "raw_fields": item.raw_fields,
                "quantity": item.quantity,
                "unit": item.unit,
                "source": item.source,
                "page_number": item.page_number,
                "reference_number": getattr(item, 'reference_number', None),
                "type": "Sub Table"
            })

        # Combine MISSING items (PDF items not in Excel) for frontend display
        combined_missing_items = []

        # Add main table missing items with type indicator
        for item in missing_main_items:
            combined_missing_items.append({
                "item_key": item.item_key,
                "raw_fields": item.raw_fields,
                "quantity": item.quantity,
                "unit": item.unit,
                "source": item.source,
                "page_number": item.page_number,
                "type": "Main Table"
            })

        # Add subtable missing items with type indicator
        for item in missing_subtable_items:
            combined_missing_items.append({
                "item_key": item.item_key,
                "raw_fields": item.raw_fields,
                "quantity": item.quantity,
                "unit": item.unit,
                "source": item.source,
                "page_number": item.page_number,
                "reference_number": getattr(item, 'reference_number', None),
                "type": "Sub Table"
            })

        logger.info(f"=== CACHED COMPARISON COMPLETED ===")
        logger.info(f"Extra items: {len(extra_main_items)} main, {len(extra_subtable_items)} subtable")
        logger.info(f"Missing items: {len(missing_main_items)} main, {len(missing_subtable_items)} subtable")

        # Calculate detailed breakdown for comprehensive summary
        pdf_main_total = len(pdf_items)
        pdf_subtable_total = len(pdf_subtables)
        excel_main_total = len(excel_items)
        excel_subtable_total = len(excel_subtables)
        
        return {
            "session_id": session_id,
            # Basic totals
            "total_pdf_items": pdf_main_total,
            "total_excel_items": excel_main_total,
            "total_pdf_subtables": pdf_subtable_total,
            "total_excel_subtables": excel_subtable_total,
            
            # Extra items (Excel items not in PDF)
            "extra_items_count": len(combined_extra_items),
            "extra_main_items_count": len(extra_main_items),
            "extra_subtable_items_count": len(extra_subtable_items),
            "extra_items": combined_extra_items,
            
            # Missing items (PDF items not in Excel) - FIXED: Now included!
            "missing_items_count": len(combined_missing_items),
            "missing_main_items_count": len(missing_main_items),
            "missing_subtable_items_count": len(missing_subtable_items),
            "missing_items": combined_missing_items,
            
            # Quantity and Unit mismatch data
            "quantity_mismatches_count": len(main_quantity_mismatches) + len(subtable_quantity_mismatches),
            "quantity_mismatches": main_quantity_mismatches + subtable_quantity_mismatches,
            "unit_mismatches_count": len(main_unit_mismatches) + len(subtable_unit_mismatches),
            "unit_mismatches": main_unit_mismatches + subtable_unit_mismatches,
            
            # DETAILED BREAKDOWN FOR SUMMARY CARDS
            "detailed_breakdown": {
                "pdf_analysis": {
                    "main_table": {
                        "total_items": pdf_main_total,
                        "missing_items": len(missing_main_items),
                        "missing_percentage": round((len(missing_main_items) / pdf_main_total * 100), 1) if pdf_main_total > 0 else 0,
                        "quantity_mismatches": len(main_quantity_mismatches),
                        "quantity_mismatch_percentage": round((len(main_quantity_mismatches) / pdf_main_total * 100), 1) if pdf_main_total > 0 else 0,
                        "unit_mismatches": len(main_unit_mismatches),
                        "unit_mismatch_percentage": round((len(main_unit_mismatches) / pdf_main_total * 100), 1) if pdf_main_total > 0 else 0
                    },
                    "subtable": {
                        "total_items": pdf_subtable_total,
                        "missing_items": len(missing_subtable_items),
                        "missing_percentage": round((len(missing_subtable_items) / pdf_subtable_total * 100), 1) if pdf_subtable_total > 0 else 0,
                        "quantity_mismatches": len(subtable_quantity_mismatches),
                        "quantity_mismatch_percentage": round((len(subtable_quantity_mismatches) / pdf_subtable_total * 100), 1) if pdf_subtable_total > 0 else 0,
                        "unit_mismatches": len(subtable_unit_mismatches),
                        "unit_mismatch_percentage": round((len(subtable_unit_mismatches) / pdf_subtable_total * 100), 1) if pdf_subtable_total > 0 else 0
                    },
                    "overall": {
                        "total_items": pdf_main_total + pdf_subtable_total,
                        "missing_items": len(missing_main_items) + len(missing_subtable_items),
                        "missing_percentage": round(((len(missing_main_items) + len(missing_subtable_items)) / (pdf_main_total + pdf_subtable_total) * 100), 1) if (pdf_main_total + pdf_subtable_total) > 0 else 0,
                        "quantity_mismatches": len(main_quantity_mismatches) + len(subtable_quantity_mismatches),
                        "quantity_mismatch_percentage": round(((len(main_quantity_mismatches) + len(subtable_quantity_mismatches)) / (pdf_main_total + pdf_subtable_total) * 100), 1) if (pdf_main_total + pdf_subtable_total) > 0 else 0,
                        "unit_mismatches": len(main_unit_mismatches) + len(subtable_unit_mismatches),
                        "unit_mismatch_percentage": round(((len(main_unit_mismatches) + len(subtable_unit_mismatches)) / (pdf_main_total + pdf_subtable_total) * 100), 1) if (pdf_main_total + pdf_subtable_total) > 0 else 0
                    }
                },
                "excel_analysis": {
                    "main_table": {
                        "total_items": excel_main_total,
                        "extra_items": len(extra_main_items),
                        "extra_percentage": round((len(extra_main_items) / excel_main_total * 100), 1) if excel_main_total > 0 else 0,
                        "quantity_mismatches": len(main_quantity_mismatches),
                        "quantity_mismatch_percentage": round((len(main_quantity_mismatches) / excel_main_total * 100), 1) if excel_main_total > 0 else 0,
                        "unit_mismatches": len(main_unit_mismatches),
                        "unit_mismatch_percentage": round((len(main_unit_mismatches) / excel_main_total * 100), 1) if excel_main_total > 0 else 0
                    },
                    "subtable": {
                        "total_items": excel_subtable_total,
                        "extra_items": len(extra_subtable_items),
                        "extra_percentage": round((len(extra_subtable_items) / excel_subtable_total * 100), 1) if excel_subtable_total > 0 else 0,
                        "quantity_mismatches": len(subtable_quantity_mismatches),
                        "quantity_mismatch_percentage": round((len(subtable_quantity_mismatches) / excel_subtable_total * 100), 1) if excel_subtable_total > 0 else 0,
                        "unit_mismatches": len(subtable_unit_mismatches),
                        "unit_mismatch_percentage": round((len(subtable_unit_mismatches) / excel_subtable_total * 100), 1) if excel_subtable_total > 0 else 0
                    },
                    "overall": {
                        "total_items": excel_main_total + excel_subtable_total,
                        "extra_items": len(extra_main_items) + len(extra_subtable_items),
                        "extra_percentage": round(((len(extra_main_items) + len(extra_subtable_items)) / (excel_main_total + excel_subtable_total) * 100), 1) if (excel_main_total + excel_subtable_total) > 0 else 0,
                        "quantity_mismatches": len(main_quantity_mismatches) + len(subtable_quantity_mismatches),
                        "quantity_mismatch_percentage": round(((len(main_quantity_mismatches) + len(subtable_quantity_mismatches)) / (excel_main_total + excel_subtable_total) * 100), 1) if (excel_main_total + excel_subtable_total) > 0 else 0,
                        "unit_mismatches": len(main_unit_mismatches) + len(subtable_unit_mismatches),
                        "unit_mismatch_percentage": round(((len(main_unit_mismatches) + len(subtable_unit_mismatches)) / (excel_main_total + excel_subtable_total) * 100), 1) if (excel_main_total + excel_subtable_total) > 0 else 0
                    }
                }
            },
            # CRITICAL: Add extracted items arrays for frontend table display
            "pdf_extracted_items": [
                {
                    "item_key": item.item_key,
                    "raw_fields": item.raw_fields,
                    "quantity": item.quantity,
                    "unit": item.unit,
                    "source": item.source,
                    "page_number": item.page_number
                }
                for item in pdf_items
            ],
            "excel_extracted_items": [
                {
                    "item_key": item.item_key,
                    "raw_fields": item.raw_fields,
                    "quantity": item.quantity,
                    "unit": item.unit,
                    "source": item.source,
                    "page_number": item.page_number
                }
                for item in excel_items
            ],
            # Add subtable items arrays for separate tables
            "pdf_subtable_items": [
                {
                    "item_key": item.item_key,
                    "raw_fields": item.raw_fields,
                    "quantity": item.quantity,
                    "unit": item.unit,
                    "source": item.source,
                    "page_number": item.page_number,
                    "reference_number": getattr(item, 'reference_number', None)
                }
                for item in pdf_subtables
            ],
            "excel_subtable_items": [
                {
                    "item_key": item.item_key,
                    "raw_fields": item.raw_fields,
                    "quantity": item.quantity,
                    "unit": item.unit,
                    "source": item.source,
                    "page_number": item.page_number,
                    "reference_number": getattr(item, 'reference_number', None)
                }
                for item in excel_subtables
            ],
            "extraction_parameters": cached_data['extraction_params'],
            "performance_note": "Used cached extraction results - much faster!"
        }

    except HTTPException:
        raise
    except Exception as e:
        logger.error(
            f"Error in cached extra items comparison: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Comparison error: {str(e)}")


@router.post("/compare-cached-subtables")
async def compare_cached_subtables(session_id: str = Form(...)):
    """
    OPTIMIZED SUBTABLE COMPARISON: Get subtables using cached extraction results.

    Args:
        session_id: Session identifier from extract-and-cache endpoint

    Returns:
        Subtable comparison results
    """
    logger.info(
        f"=== COMPARING SUBTABLES FROM CACHED SESSION {session_id} ===")

    try:
        # Get cached extraction results
        cache_service = get_extraction_cache()
        cached_data = cache_service.get_extraction_results(session_id)

        if not cached_data:
            raise HTTPException(
                status_code=404,
                detail="Session not found or expired. Please re-extract files first."
            )

        # Extract cached subtable data
        pdf_subtables = cached_data['pdf_subtables']
        excel_subtables = cached_data['excel_subtables']

        logger.info(
            f"Using cached subtable data: {len(pdf_subtables)} PDF subtables, {len(excel_subtables)} Excel subtables")

        # Create response using SubtableComparisonSummary schema
        result = SubtableComparisonSummary(
            total_pdf_subtables=len(pdf_subtables),
            total_excel_subtables=len(excel_subtables),
            pdf_subtables=pdf_subtables,
            excel_subtables=excel_subtables
        )

        logger.info("=== CACHED SUBTABLE COMPARISON COMPLETED ===")

        return result

    except HTTPException:
        raise
    except Exception as e:
        logger.error(
            f"Error in cached subtable comparison: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Subtable comparison error: {str(e)}")


@router.post("/cleanup-session")
async def cleanup_session(session_id: str = Form(...)):
    """
    CLEANUP ENDPOINT: Manually cleanup a cached session to free memory.

    Args:
        session_id: Session identifier to cleanup

    Returns:
        Cleanup status
    """
    logger.info(f"=== CLEANING UP SESSION {session_id} ===")

    try:
        cache_service = get_extraction_cache()
        cleaned = cache_service.cleanup_session(session_id)

        if cleaned:
            logger.info(f"Successfully cleaned up session {session_id}")
            return {
                "status": "success",
                "message": f"Session {session_id} has been cleaned up",
                "session_id": session_id
            }
        else:
            logger.warning(
                f"Session {session_id} not found or already cleaned")
            return {
                "status": "not_found",
                "message": f"Session {session_id} not found or already cleaned",
                "session_id": session_id
            }

    except Exception as e:
        logger.error(
            f"Error cleaning up session {session_id}: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Cleanup error: {str(e)}")


@router.get("/cache-stats")
async def get_cache_stats():
    """
    Get cache statistics for monitoring and debugging.

    Returns:
        Cache statistics
    """
    try:
        cache_service = get_extraction_cache()

        # Clean up expired sessions first
        expired_cleaned = cache_service.cleanup_expired_sessions()

        # Get current stats
        stats = cache_service.get_cache_stats()
        stats['expired_sessions_cleaned'] = expired_cleaned

        logger.info(
            f"Cache stats requested - {stats['active_sessions']} active sessions")

        return {
            "status": "success",
            "cache_statistics": stats,
            "performance_tip": "Use cached comparisons for better performance!"
        }

    except Exception as e:
        logger.error(f"Error getting cache stats: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Stats error: {str(e)}")


@router.post("/compare", response_model=ComparisonSummary)
async def compare_tender_files(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    start_page: Optional[int] = Form(None),
    end_page: Optional[int] = Form(None),
    sheet_name: Optional[str] = Form(None)
) -> ComparisonSummary:
    """
    Compare a tender PDF with an Excel proposal iteratively.
    Focus on finding mismatches and items present in PDF but not in Excel.

    Args:
        pdf_file: PDF tender document
        excel_file: Excel proposal document
        start_page: Starting page number for PDF extraction (optional)
        end_page: Ending page number for PDF extraction (optional)
        sheet_name: Specific Excel sheet name to extract from (optional)
    """
    logger.info("=== STARTING TENDER COMPARISON ===")
    logger.info(f"PDF file: {pdf_file.filename}")
    logger.info(f"Excel file: {excel_file.filename}")
    logger.info(f"PDF page range: {start_page} to {end_page}")
    logger.info(f"Excel sheet: {sheet_name or 'All sheets'}")

    # Validate file types
    if not pdf_file.filename.lower().endswith('.pdf'):
        logger.error(f"Invalid PDF file: {pdf_file.filename}")
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        logger.error(f"Invalid Excel file: {excel_file.filename}")
        raise HTTPException(status_code=400, detail="Excel file required")

    # Validate page range
    if start_page is not None and start_page < 1:
        raise HTTPException(status_code=400, detail="Start page must be >= 1")
    if end_page is not None and end_page < 1:
        raise HTTPException(status_code=400, detail="End page must be >= 1")
    if start_page is not None and end_page is not None and start_page > end_page:
        raise HTTPException(
            status_code=400, detail="Start page cannot be greater than end page")

    # For PDF, we still need a temporary file (pdfplumber requires file path)
    # For Excel, we can use in-memory processing
    pdf_fd = None
    pdf_path = None

    try:
        logger.info("=== PROCESSING PDF FILE ===")
        # Handle PDF file (requires temporary file)
        pdf_fd, pdf_path = tempfile.mkstemp(suffix='.pdf', prefix='tender_')

        # Save PDF to temporary file
        pdf_content = await pdf_file.read()
        logger.info(f"PDF file size: {len(pdf_content)} bytes")

        with os.fdopen(pdf_fd, 'wb') as f:
            f.write(pdf_content)
        pdf_fd = None  # File descriptor is closed now

        logger.info("=== PROCESSING EXCEL FILE ===")
        # Handle Excel file (in-memory processing)
        excel_content = await excel_file.read()
        logger.info(f"Excel file size: {len(excel_content)} bytes")

        # Create in-memory buffer for Excel
        excel_buffer = BytesIO(excel_content)

        logger.info("=== STARTING EXTRACTION PROCESS ===")
        # Parse files iteratively with parameters
        pdf_parser = PDFParser()
        excel_parser = ExcelParser()

        # Extract items from PDF with page range
        logger.info("Extracting items from PDF with specified parameters...")
        pdf_items = pdf_parser.extract_tables_with_range(
            pdf_path, start_page, end_page)
        logger.info(f"Total PDF items extracted: {len(pdf_items)}")

        # Extract items from Excel with sheet filter
        logger.info("Extracting items from Excel with specified parameters...")
        if sheet_name:
            # Use the new standalone corrected main table extraction for specific sheet
            excel_table_extractor = ExcelTableExtractorService()
            excel_items = excel_table_extractor.extract_main_table_from_buffer(
                excel_buffer, sheet_name)
            logger.info(
                f"Total Excel items extracted using standalone corrected logic: {len(excel_items)}")
        else:
            # Use original method for all sheets
            excel_items = excel_parser.extract_items_from_buffer_with_sheet(
                excel_buffer, sheet_name)
            logger.info(f"Total Excel items extracted: {len(excel_items)}")

        # Close Excel buffer
        excel_buffer.close()

        if not pdf_items:
            logger.warning("No items found in PDF!")
            raise HTTPException(
                status_code=400, detail="No extractable items found in PDF")

        if not excel_items:
            logger.warning("No items found in Excel!")
            raise HTTPException(
                status_code=400, detail="No extractable items found in Excel")

        logger.info("=== STARTING COMPARISON PROCESS ===")
        # Compare items focusing on mismatches and missing items
        matcher = Matcher()
        result = matcher.compare_items(pdf_items, excel_items)

        logger.info("=== COMPARISON COMPLETED ===")
        logger.info(f"Summary: {result.matched_items} matches, {result.quantity_mismatches} mismatches, "
                    f"{result.missing_items} missing, {result.extra_items} extra")

        # Clean up references
        del pdf_parser, excel_parser, matcher, pdf_items, excel_items, excel_buffer
        gc.collect()

        return result

    except HTTPException:
        # Re-raise HTTP exceptions
        raise
    except Exception as e:
        logger.error(
            f"Error during comparison process: {str(e)}", exc_info=True)
        gc.collect()
        raise HTTPException(
            status_code=500, detail=f"Processing error: {str(e)}")

    finally:
        # Clean up PDF temporary file
        try:
            if pdf_fd is not None:
                os.close(pdf_fd)
        except Exception as e:
            logger.error(f"Error closing PDF file descriptor: {e}")

        gc.collect()

        # Remove PDF temporary file
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.unlink(pdf_path)
                logger.info("Temporary PDF file cleaned up")
            except Exception as e:
                logger.warning(f"Could not remove temporary PDF file: {e}")
                try:
                    import time
                    time.sleep(0.5)
                    os.unlink(pdf_path)
                    logger.info("Temporary PDF file cleaned up (retry)")
                except Exception as e2:
                    logger.error(
                        f"Failed to clean up temporary PDF file: {e2}")


@router.post("/compare-main-table-corrected", response_model=ComparisonSummary)
async def compare_main_table_corrected(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    start_page: Optional[int] = Form(None),
    end_page: Optional[int] = Form(None),
    sheet_name: str = Form(...)
) -> ComparisonSummary:
    """
    Compare a tender PDF with an Excel main table using the corrected extraction logic.
    This endpoint specifically uses the new ExcelTableExtractorCorrected logic for main table extraction.
    Focus on finding mismatches and items present in PDF but not in Excel.

    Args:
        pdf_file: PDF tender document
        excel_file: Excel proposal document
        start_page: Starting page number for PDF extraction (optional)
        end_page: Ending page number for PDF extraction (optional)
        sheet_name: Specific Excel sheet name to extract from (required for corrected logic)
    """
    logger.info("=== STARTING CORRECTED MAIN TABLE COMPARISON ===")
    logger.info(f"PDF file: {pdf_file.filename}")
    logger.info(f"Excel file: {excel_file.filename}")
    logger.info(f"PDF page range: {start_page} to {end_page}")
    logger.info(f"Excel sheet: {sheet_name}")

    # Validate file types
    if not pdf_file.filename.lower().endswith('.pdf'):
        logger.error(f"Invalid PDF file: {pdf_file.filename}")
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        logger.error(f"Invalid Excel file: {excel_file.filename}")
        raise HTTPException(status_code=400, detail="Excel file required")

    # Validate page range
    if start_page is not None and start_page < 1:
        raise HTTPException(status_code=400, detail="Start page must be >= 1")
    if end_page is not None and end_page < 1:
        raise HTTPException(status_code=400, detail="End page must be >= 1")
    if start_page is not None and end_page is not None and start_page > end_page:
        raise HTTPException(
            status_code=400, detail="Start page cannot be greater than end page")

    # For PDF, we still need a temporary file (pdfplumber requires file path)
    # For Excel, we can use in-memory processing
    pdf_fd = None
    pdf_path = None

    try:
        logger.info("=== PROCESSING PDF FILE ===")
        # Handle PDF file (requires temporary file)
        pdf_fd, pdf_path = tempfile.mkstemp(suffix='.pdf', prefix='tender_')

        # Save PDF to temporary file
        pdf_content = await pdf_file.read()
        logger.info(f"PDF file size: {len(pdf_content)} bytes")

        with os.fdopen(pdf_fd, 'wb') as f:
            f.write(pdf_content)
        pdf_fd = None  # File descriptor is closed now

        logger.info("=== PROCESSING EXCEL FILE ===")
        # Handle Excel file (in-memory processing)
        excel_content = await excel_file.read()
        logger.info(f"Excel file size: {len(excel_content)} bytes")

        # Create in-memory buffer for Excel
        excel_buffer = BytesIO(excel_content)

        logger.info("=== STARTING CORRECTED EXTRACTION PROCESS ===")
        # Parse files iteratively with parameters
        pdf_parser = PDFParser()
        excel_parser = ExcelParser()

        # Extract items from PDF with page range
        logger.info("Extracting items from PDF with specified parameters...")
        pdf_items = pdf_parser.extract_tables_with_range(
            pdf_path, start_page, end_page)
        logger.info(f"Total PDF items extracted: {len(pdf_items)}")

        # Extract items from Excel using corrected main table extraction
        logger.info(
            "Extracting items from Excel using corrected main table logic...")
        excel_table_extractor = ExcelTableExtractorService()
        excel_items = excel_table_extractor.extract_main_table_from_buffer(
            excel_buffer, sheet_name)
        logger.info(
            f"Total Excel items extracted using standalone corrected logic: {len(excel_items)}")

        # Close Excel buffer
        excel_buffer.close()

        if not pdf_items:
            logger.warning("No items found in PDF!")
            raise HTTPException(
                status_code=400, detail="No extractable items found in PDF")

        if not excel_items:
            logger.warning("No items found in Excel!")
            raise HTTPException(
                status_code=400, detail="No extractable items found in Excel")

        logger.info("=== STARTING COMPARISON PROCESS ===")
        # Compare items focusing on mismatches and missing items
        matcher = Matcher()
        result = matcher.compare_items(pdf_items, excel_items)

        logger.info("=== CORRECTED COMPARISON COMPLETED ===")
        logger.info(f"Summary: {result.matched_items} matches, {result.quantity_mismatches} mismatches, "
                    f"{result.missing_items} missing, {result.extra_items} extra")

        # Clean up references
        del pdf_parser, excel_parser, matcher, pdf_items, excel_items, excel_buffer
        gc.collect()

        return result

    except HTTPException:
        # Re-raise HTTP exceptions
        raise
    except Exception as e:
        logger.error(
            f"Error during corrected comparison process: {str(e)}", exc_info=True)
        gc.collect()
        raise HTTPException(
            status_code=500, detail=f"Processing error: {str(e)}")

    finally:
        # Clean up PDF temporary file
        try:
            if pdf_fd is not None:
                os.close(pdf_fd)
        except Exception as e:
            logger.error(f"Error closing PDF file descriptor: {e}")

        gc.collect()

        # Remove PDF temporary file
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.unlink(pdf_path)
                logger.info("Temporary PDF file cleaned up")
            except Exception as e:
                logger.warning(f"Could not remove temporary PDF file: {e}")
                try:
                    import time
                    time.sleep(0.5)
                    os.unlink(pdf_path)
                    logger.info("Temporary PDF file cleaned up (retry)")
                except Exception as e2:
                    logger.error(
                        f"Failed to clean up temporary PDF file: {e2}")


@router.post("/compare-missing-only")
async def compare_tender_files_missing_only(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    start_page: Optional[int] = Form(None),
    end_page: Optional[int] = Form(None),
    sheet_name: Optional[str] = Form(None),
    pdf_subtable_start_page: Optional[int] = Form(None),
    pdf_subtable_end_page: Optional[int] = Form(None)
):
    """
    Compare tender files and return only the missing items (PDF items not in Excel), including both main table and sub table.
    """
    logger.info("=== STARTING MISSING ITEMS COMPARISON ===")
    logger.info(f"PDF main table page range: {start_page} to {end_page}")
    logger.info(
        f"PDF subtable page range: {pdf_subtable_start_page} to {pdf_subtable_end_page}")
    logger.info(f"Excel sheet: {sheet_name or 'All sheets'}")

    if not pdf_file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Excel file required")

    # Validate main table page range
    if start_page is not None and start_page < 1:
        raise HTTPException(status_code=400, detail="Start page must be >= 1")
    if end_page is not None and end_page < 1:
        raise HTTPException(status_code=400, detail="End page must be >= 1")
    if start_page is not None and end_page is not None and start_page > end_page:
        raise HTTPException(
            status_code=400, detail="Start page cannot be greater than end page")

    # Validate subtable page range
    if pdf_subtable_start_page is not None and pdf_subtable_start_page < 1:
        raise HTTPException(
            status_code=400, detail="Subtable start page must be >= 1")
    if pdf_subtable_end_page is not None and pdf_subtable_end_page < 1:
        raise HTTPException(
            status_code=400, detail="Subtable end page must be >= 1")
    if pdf_subtable_start_page is not None and pdf_subtable_end_page is not None and pdf_subtable_start_page > pdf_subtable_end_page:
        raise HTTPException(
            status_code=400, detail="Subtable start page cannot be greater than end page")

    pdf_fd = None
    pdf_path = None

    try:
        pdf_fd, pdf_path = tempfile.mkstemp(suffix='.pdf', prefix='tender_')
        pdf_content = await pdf_file.read()
        with os.fdopen(pdf_fd, 'wb') as f:
            f.write(pdf_content)
        pdf_fd = None

        excel_content = await excel_file.read()
        excel_buffer = BytesIO(excel_content)

        pdf_parser = PDFParser()
        excel_parser = ExcelParser()

        pdf_items = pdf_parser.extract_tables_with_range(
            pdf_path, start_page, end_page)

        if sheet_name:
            excel_table_extractor = ExcelTableExtractorService()
            excel_items = excel_table_extractor.extract_main_table_from_buffer(
                excel_buffer, sheet_name)
        else:
            excel_items = excel_parser.extract_items_from_buffer_with_sheet(
                excel_buffer, sheet_name)

        # --- Subtable extraction ---
        # For subtables, use the same sheet_name as main table for Excel
        excel_table_extractor = ExcelTableExtractorService()
        main_table_items = excel_items
        pdf_subtables = pdf_parser.extract_subtables_with_range(
            pdf_path,
            pdf_subtable_start_page,
            pdf_subtable_end_page
        )
        logger.info("Using NEW API-ready Excel subtable extraction...")
        excel_subtables = excel_table_extractor.extract_subtables_from_buffer(
            excel_buffer, sheet_name or (excel_buffer and getattr(excel_buffer, 'name', None)) or '', main_table_items)
        logger.info(f"NEW API extracted {len(excel_subtables)} Excel subtables")

        excel_buffer.close()

        matcher = Matcher()
        # Main table missing items
        main_missing_items = matcher.get_missing_items_only(
            pdf_items, excel_items)
        # Subtable missing items (use only those with status == 'MISSING')
        subtable_results = matcher.compare_subtable_items(
            pdf_subtables, excel_subtables)
        subtable_missing_items = [
            r for r in subtable_results if r.status == 'MISSING']

        # Format for frontend: add 'type' field
        missing_items = [
            {
                "item_key": item.item_key,
                "raw_fields": item.raw_fields,
                "quantity": item.quantity,
                "unit": item.unit,
                "page_number": item.page_number,
                "type": "Main Table"
            }
            for item in main_missing_items
        ] + [
            {
                "item_key": r.pdf_item.item_key,
                "raw_fields": r.pdf_item.raw_fields,
                "quantity": r.pdf_item.quantity,
                "unit": r.pdf_item.unit,
                "page_number": r.pdf_item.page_number,
                "reference_number": r.pdf_item.reference_number,
                "type": "Sub Table"
            }
            for r in subtable_missing_items if r.pdf_item is not None
        ]

        return {
            "total_pdf_items": len(pdf_items),
            "total_excel_items": len(excel_items),
            "missing_items_count": len(missing_items),
            "missing_items": missing_items,
            "pdf_extracted_items": [
                {
                    "item_key": item.item_key,
                    "raw_fields": item.raw_fields,
                    "quantity": item.quantity,
                    "unit": item.unit,
                    "source": item.source,
                    "page_number": item.page_number
                }
                for item in pdf_items
            ],
            "excel_extracted_items": [
                {
                    "item_key": item.item_key,
                    "raw_fields": item.raw_fields,
                    "quantity": item.quantity,
                    "unit": item.unit,
                    "source": item.source,
                    "page_number": item.page_number
                }
                for item in excel_items
            ]
        }

    except Exception as e:
        logger.error(
            f"Error in missing items comparison: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Processing error: {str(e)}")

    finally:
        # Cleanup
        try:
            if pdf_fd is not None:
                os.close(pdf_fd)
        except Exception:
            pass

        if pdf_path and os.path.exists(pdf_path):
            try:
                os.unlink(pdf_path)
            except Exception:
                pass

        gc.collect()


@router.post("/compare-mismatches-only")
async def compare_tender_files_mismatches_only(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    start_page: Optional[int] = Form(None),
    end_page: Optional[int] = Form(None),
    sheet_name: Optional[str] = Form(None),
    pdf_subtable_start_page: Optional[int] = Form(None),
    pdf_subtable_end_page: Optional[int] = Form(None)
):
    """
    Compare tender files and return only the quantity mismatches, including both main table and sub table.
    """
    logger.info("=== STARTING QUANTITY MISMATCHES COMPARISON ===")
    logger.info(f"PDF main table page range: {start_page} to {end_page}")
    logger.info(
        f"PDF subtable page range: {pdf_subtable_start_page} to {pdf_subtable_end_page}")
    logger.info(f"Excel sheet: {sheet_name or 'All sheets'}")

    if not pdf_file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Excel file required")

    # Validate main table page range
    if start_page is not None and start_page < 1:
        raise HTTPException(status_code=400, detail="Start page must be >= 1")
    if end_page is not None and end_page < 1:
        raise HTTPException(status_code=400, detail="End page must be >= 1")
    if start_page is not None and end_page is not None and start_page > end_page:
        raise HTTPException(
            status_code=400, detail="Start page cannot be greater than end page")

    # Validate subtable page range
    if pdf_subtable_start_page is not None and pdf_subtable_start_page < 1:
        raise HTTPException(
            status_code=400, detail="Subtable start page must be >= 1")
    if pdf_subtable_end_page is not None and pdf_subtable_end_page < 1:
        raise HTTPException(
            status_code=400, detail="Subtable end page must be >= 1")
    if pdf_subtable_start_page is not None and pdf_subtable_end_page is not None and pdf_subtable_start_page > pdf_subtable_end_page:
        raise HTTPException(
            status_code=400, detail="Subtable start page cannot be greater than end page")

    pdf_fd = None
    pdf_path = None

    try:
        pdf_fd, pdf_path = tempfile.mkstemp(suffix='.pdf', prefix='tender_')
        pdf_content = await pdf_file.read()
        with os.fdopen(pdf_fd, 'wb') as f:
            f.write(pdf_content)
        pdf_fd = None

        excel_content = await excel_file.read()
        excel_buffer = BytesIO(excel_content)

        pdf_parser = PDFParser()
        excel_parser = ExcelParser()

        pdf_items = pdf_parser.extract_tables_with_range(
            pdf_path, start_page, end_page)

        if sheet_name:
            excel_table_extractor = ExcelTableExtractorService()
            excel_items = excel_table_extractor.extract_main_table_from_buffer(
                excel_buffer, sheet_name)
        else:
            excel_items = excel_parser.extract_items_from_buffer_with_sheet(
                excel_buffer, sheet_name)

        # --- Subtable extraction ---
        excel_table_extractor = ExcelTableExtractorService()
        main_table_items = excel_items
        pdf_subtables = pdf_parser.extract_subtables_with_range(
            pdf_path,
            pdf_subtable_start_page,
            pdf_subtable_end_page
        )
        logger.info("Using NEW API-ready Excel subtable extraction...")
        excel_subtables = excel_table_extractor.extract_subtables_from_buffer(
            excel_buffer, sheet_name or (excel_buffer and getattr(excel_buffer, 'name', None)) or '', main_table_items)
        logger.info(f"NEW API extracted {len(excel_subtables)} Excel subtables")

        excel_buffer.close()

        matcher = Matcher()
        # Main table quantity mismatches
        main_mismatched_results = matcher.get_mismatched_items_only(
            pdf_items, excel_items)
        # Subtable quantity mismatches
        subtable_results = matcher.compare_subtable_items(
            pdf_subtables, excel_subtables)
        subtable_mismatches = [
            r for r in subtable_results if r.status == 'QUANTITY_MISMATCH']

        # Format for frontend: add 'type' field
        quantity_mismatches = [
            {
                "pdf_item": {
                    "item_key": r.pdf_item.item_key,
                    "quantity": r.pdf_item.quantity,
                    "unit": r.pdf_item.unit,
                    "raw_fields": r.pdf_item.raw_fields,
                    "page_number": r.pdf_item.page_number
                } if r.pdf_item else None,
                "excel_item": {
                    "item_key": r.excel_item.item_key,
                    "quantity": r.excel_item.quantity,
                    "unit": r.excel_item.unit,
                    "raw_fields": r.excel_item.raw_fields,
                    "page_number": r.excel_item.page_number
                } if r.excel_item else None,
                "quantity_difference": r.quantity_difference,
                "match_confidence": r.match_confidence,
                "type": r.type
            }
            for r in subtable_mismatches
        ] + [
            {
                "pdf_item": {
                    "item_key": result.pdf_item.item_key,
                    "quantity": result.pdf_item.quantity,
                    "unit": result.pdf_item.unit,
                    "raw_fields": result.pdf_item.raw_fields,
                    "page_number": result.pdf_item.page_number
                } if result.pdf_item else None,
                "excel_item": {
                    "item_key": result.excel_item.item_key,
                    "quantity": result.excel_item.quantity,
                    "unit": result.excel_item.unit,
                    "raw_fields": result.excel_item.raw_fields,
                    "page_number": result.excel_item.page_number
                } if result.excel_item else None,
                "quantity_difference": result.quantity_difference,
                "match_confidence": result.match_confidence,
                "type": "Main Table"
            }
            for result in main_mismatched_results
        ]

        return {
            "total_pdf_items": len(pdf_items),
            "total_excel_items": len(excel_items),
            "quantity_mismatches_count": len(quantity_mismatches),
            "quantity_mismatches": quantity_mismatches,
            "pdf_extracted_items": [
                {
                    "item_key": item.item_key,
                    "raw_fields": item.raw_fields,
                    "quantity": item.quantity,
                    "unit": item.unit,
                    "source": item.source,
                    "page_number": item.page_number
                }
                for item in pdf_items
            ],
            "excel_extracted_items": [
                {
                    "item_key": item.item_key,
                    "raw_fields": item.raw_fields,
                    "quantity": item.quantity,
                    "unit": item.unit,
                    "source": item.source,
                    "page_number": item.page_number
                }
                for item in excel_items
            ]
        }

    except Exception as e:
        logger.error(
            f"Error in mismatches comparison: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Processing error: {str(e)}")

    finally:
        # Cleanup
        try:
            if pdf_fd is not None:
                os.close(pdf_fd)
        except Exception:
            pass

        if pdf_path and os.path.exists(pdf_path):
            try:
                os.unlink(pdf_path)
            except Exception:
                pass

        gc.collect()


@router.post("/compare-unit-mismatches-only")
async def compare_tender_files_unit_mismatches_only(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    start_page: Optional[int] = Form(None),
    end_page: Optional[int] = Form(None),
    sheet_name: Optional[str] = Form(None)
):
    """
    Compare tender files and return only the unit mismatches.
    Optimized endpoint for specific use case.

    Args:
        pdf_file: PDF tender document
        excel_file: Excel proposal document
        start_page: Starting page number for PDF extraction (optional)
        end_page: Ending page number for PDF extraction (optional)
        sheet_name: Specific Excel sheet name to extract from (optional)
    """
    logger.info("=== STARTING UNIT MISMATCHES COMPARISON ===")
    logger.info(f"PDF page range: {start_page} to {end_page}")
    logger.info(f"Excel sheet: {sheet_name or 'All sheets'}")

    # Same file processing as main endpoint
    if not pdf_file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Excel file required")

    # Validate page range
    if start_page is not None and start_page < 1:
        raise HTTPException(status_code=400, detail="Start page must be >= 1")
    if end_page is not None and end_page < 1:
        raise HTTPException(status_code=400, detail="End page must be >= 1")
    if start_page is not None and end_page is not None and start_page > end_page:
        raise HTTPException(
            status_code=400, detail="Start page cannot be greater than end page")

    pdf_fd = None
    pdf_path = None

    try:
        # Process files
        pdf_fd, pdf_path = tempfile.mkstemp(suffix='.pdf', prefix='tender_')
        pdf_content = await pdf_file.read()

        with os.fdopen(pdf_fd, 'wb') as f:
            f.write(pdf_content)
        pdf_fd = None

        excel_content = await excel_file.read()
        excel_buffer = BytesIO(excel_content)

        # Parse files with parameters
        pdf_parser = PDFParser()
        excel_parser = ExcelParser()

        pdf_items = pdf_parser.extract_tables_with_range(
            pdf_path, start_page, end_page)

        # Extract items from Excel with sheet filter
        logger.info("Extracting items from Excel with specified parameters...")
        if sheet_name:
            # Use the new standalone corrected main table extraction for specific sheet
            excel_table_extractor = ExcelTableExtractorService()
            excel_items = excel_table_extractor.extract_main_table_from_buffer(
                excel_buffer, sheet_name)
            logger.info(
                f"Total Excel items extracted using standalone corrected logic: {len(excel_items)}")
        else:
            # Use original method for all sheets
            excel_items = excel_parser.extract_items_from_buffer_with_sheet(
                excel_buffer, sheet_name)
            logger.info(f"Total Excel items extracted: {len(excel_items)}")

        excel_buffer.close()

        # Get only unit mismatched items
        matcher = Matcher()
        unit_mismatched_results = matcher.get_unit_mismatched_items_only(
            pdf_items, excel_items)

        # Return simplified response
        return {
            "total_pdf_items": len(pdf_items),
            "total_excel_items": len(excel_items),
            "unit_mismatches_count": len(unit_mismatched_results),
            "unit_mismatches": [
                {
                    "pdf_item": {
                        "item_key": result.pdf_item.item_key,
                        "quantity": result.pdf_item.quantity,
                        "unit": result.pdf_item.unit,
                        "raw_fields": result.pdf_item.raw_fields,
                        "page_number": result.pdf_item.page_number
                    },
                    "excel_item": {
                        "item_key": result.excel_item.item_key,
                        "quantity": result.excel_item.quantity,
                        "unit": result.excel_item.unit,
                        "raw_fields": result.excel_item.raw_fields,
                        "page_number": result.excel_item.page_number
                    },
                    "match_confidence": result.match_confidence
                }
                for result in unit_mismatched_results
            ],
            "pdf_extracted_items": [
                {
                    "item_key": item.item_key,
                    "raw_fields": item.raw_fields,
                    "quantity": item.quantity,
                    "unit": item.unit,
                    "source": item.source,
                    "page_number": item.page_number
                }
                for item in pdf_items
            ],
            "excel_extracted_items": [
                {
                    "item_key": item.item_key,
                    "raw_fields": item.raw_fields,
                    "quantity": item.quantity,
                    "unit": item.unit,
                    "source": item.source,
                    "page_number": item.page_number
                }
                for item in excel_items
            ]
        }

    except Exception as e:
        logger.error(
            f"Error in unit mismatches comparison: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Processing error: {str(e)}")

    finally:
        # Cleanup
        try:
            if pdf_fd is not None:
                os.close(pdf_fd)
        except Exception:
            pass

        if pdf_path and os.path.exists(pdf_path):
            try:
                os.unlink(pdf_path)
            except Exception:
                pass

        gc.collect()


@router.post("/compare-extra-items-only")
async def compare_tender_files_extra_items_only(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    start_page: Optional[int] = Form(None),
    end_page: Optional[int] = Form(None),
    sheet_name: Optional[str] = Form(None),
    pdf_subtable_start_page: Optional[int] = Form(None),
    pdf_subtable_end_page: Optional[int] = Form(None)
):
    """
    Compare tender files and return only the extra items (Excel items not in PDF).
    Uses improved simplified matching that compares only item name, quantity, and unit.
    Compares main table items separately from subtable items.

    Args:
        pdf_file: PDF tender document
        excel_file: Excel proposal document
        start_page: Starting page number for PDF main table extraction (optional)
        end_page: Ending page number for PDF main table extraction (optional)
        sheet_name: Specific Excel sheet name to extract from (optional)
        pdf_subtable_start_page: Starting page number for PDF subtable extraction (optional)
        pdf_subtable_end_page: Ending page number for PDF subtable extraction (optional)
    """
    logger.info("=== STARTING EXTRA ITEMS COMPARISON ===")
    logger.info(f"PDF main table page range: {start_page} to {end_page}")
    logger.info(
        f"PDF subtable page range: {pdf_subtable_start_page} to {pdf_subtable_end_page}")
    logger.info(f"Excel sheet: {sheet_name or 'All sheets'}")

    # Same file processing as main endpoint
    if not pdf_file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Excel file required")

    # Validate main table page range
    if start_page is not None and start_page < 1:
        raise HTTPException(status_code=400, detail="Start page must be >= 1")
    if end_page is not None and end_page < 1:
        raise HTTPException(status_code=400, detail="End page must be >= 1")
    if start_page is not None and end_page is not None and start_page > end_page:
        raise HTTPException(
            status_code=400, detail="Start page cannot be greater than end page")

    # Validate subtable page range
    if pdf_subtable_start_page is not None and pdf_subtable_start_page < 1:
        raise HTTPException(
            status_code=400, detail="Subtable start page must be >= 1")
    if pdf_subtable_end_page is not None and pdf_subtable_end_page < 1:
        raise HTTPException(
            status_code=400, detail="Subtable end page must be >= 1")
    if pdf_subtable_start_page is not None and pdf_subtable_end_page is not None and pdf_subtable_start_page > pdf_subtable_end_page:
        raise HTTPException(
            status_code=400, detail="Subtable start page cannot be greater than end page")

    pdf_fd = None
    pdf_path = None

    try:
        # Process files
        pdf_fd, pdf_path = tempfile.mkstemp(suffix='.pdf', prefix='tender_')
        pdf_content = await pdf_file.read()

        with os.fdopen(pdf_fd, 'wb') as f:
            f.write(pdf_content)
        pdf_fd = None

        excel_content = await excel_file.read()
        excel_buffer = BytesIO(excel_content)

        # Parse files with parameters
        pdf_parser = PDFParser()
        excel_parser = ExcelParser()

        # Extract main table items from PDF
        pdf_items = pdf_parser.extract_tables_with_range(
            pdf_path, start_page, end_page)

        # Extract main table items from Excel with sheet filter
        logger.info(
            "Extracting main table items from Excel with specified parameters...")
        if sheet_name:
            # Use the new standalone corrected main table extraction for specific sheet
            excel_table_extractor = ExcelTableExtractorService()
            excel_items = excel_table_extractor.extract_main_table_from_buffer(
                excel_buffer, sheet_name)
            logger.info(
                f"Total Excel main table items extracted using standalone corrected logic: {len(excel_items)}")
        else:
            # Use original method for all sheets
            excel_items = excel_parser.extract_items_from_buffer_with_sheet(
                excel_buffer, sheet_name)
            logger.info(
                f"Total Excel main table items extracted: {len(excel_items)}")

        # --- Subtable extraction ---
        logger.info("Extracting subtable items from both PDF and Excel...")
        excel_table_extractor = ExcelTableExtractorService()
        main_table_items = excel_items
        pdf_subtables = pdf_parser.extract_subtables_with_range(
            pdf_path,
            pdf_subtable_start_page,
            pdf_subtable_end_page
        )
        logger.info("Using NEW API-ready Excel subtable extraction...")
        excel_subtables = excel_table_extractor.extract_subtables_from_buffer(
            excel_buffer, sheet_name or (excel_buffer and getattr(excel_buffer, 'name', None)) or '', main_table_items)
        logger.info(f"NEW API extracted {len(excel_subtables)} Excel subtables")

        excel_buffer.close()

        # Get extra items for both main table and subtables separately
        matcher = Matcher()
        # Always use simplified matching method for main table items (improved accuracy)
        logger.info("Using simplified matching method for main table items")
        extra_main_items = matcher.get_extra_items_only_simplified(
            pdf_items, excel_items)
        extra_subtable_items = matcher.get_extra_subtable_items_only(
            pdf_subtables, excel_subtables)

        # Combine main table and subtable extra items for frontend display
        combined_extra_items = []

        # Add main table extra items with type indicator
        for item in extra_main_items:
            combined_extra_items.append({
                "item_key": item.item_key,
                "raw_fields": item.raw_fields,
                "quantity": item.quantity,
                "unit": item.unit,
                "source": item.source,
                "page_number": item.page_number,
                "type": "Main Table"
            })

        # Add subtable extra items with type indicator
        for item in extra_subtable_items:
            combined_extra_items.append({
                "item_key": item.item_key,
                "raw_fields": item.raw_fields,
                "quantity": item.quantity,
                "unit": item.unit,
                "source": item.source,
                "page_number": item.page_number,
                "reference_number": item.reference_number,
                "type": "Sub Table"
            })

        # Return comprehensive response
        return {
            "total_pdf_items": len(pdf_items),
            "total_excel_items": len(excel_items),
            "total_pdf_subtables": len(pdf_subtables),
            "total_excel_subtables": len(excel_subtables),
            "extra_items_count": len(combined_extra_items),
            "extra_main_items_count": len(extra_main_items),
            "extra_subtable_items_count": len(extra_subtable_items),
            "extra_items": combined_extra_items,
            "pdf_extracted_items": [
                {
                    "item_key": item.item_key,
                    "raw_fields": item.raw_fields,
                    "quantity": item.quantity,
                    "unit": item.unit,
                    "source": item.source,
                    "page_number": item.page_number
                }
                for item in pdf_items
            ],
            "excel_extracted_items": [
                {
                    "item_key": item.item_key,
                    "raw_fields": item.raw_fields,
                    "quantity": item.quantity,
                    "unit": item.unit,
                    "source": item.source,
                    "page_number": item.page_number
                }
                for item in excel_items
            ]
        }

    except Exception as e:
        logger.error(
            f"Error in extra items comparison: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Processing error: {str(e)}")

    finally:
        # Cleanup
        try:
            if pdf_fd is not None:
                os.close(pdf_fd)
        except Exception:
            pass

        if pdf_path and os.path.exists(pdf_path):
            try:
                os.unlink(pdf_path)
            except Exception:
                pass

        gc.collect()


@router.post("/compare-subtables", response_model=SubtableComparisonSummary)
async def compare_subtables(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    main_sheet_name: str = Form(...),
    pdf_subtable_start_page: Optional[int] = Form(None),
    pdf_subtable_end_page: Optional[int] = Form(None)
) -> SubtableComparisonSummary:
    """
    Extract subtables from both PDF and Excel files.

    Subtable extraction for PDF:
    1. Dynamically discovers reference patterns from main table
    2. Ignores rows with only "合計" and "単価" without quantities  
    3. Looks for specific column headers: 名称・規格, 条件, 単位, 数量, etc.
    4. Finds reference numbers like "内 X号", "単 Y号" and associates table data with them

    Subtable extraction for Excel:
    1. Dynamically discovers reference patterns from main sheet
    2. Scans all non-main sheets for subtables
    3. Applies same filtering logic as PDF
    4. Supports row spanning for items with name+unit but no quantity

    Args:
        pdf_file: PDF document containing subtables
        excel_file: Excel document containing subtables in multiple sheets
        pdf_subtable_start_page: Starting page number for PDF subtable extraction
        pdf_subtable_end_page: Ending page number for PDF subtable extraction
    """
    logger.info("=== STARTING SUBTABLE COMPARISON ===")
    logger.info(f"PDF file: {pdf_file.filename}")
    logger.info(f"Excel file: {excel_file.filename}")
    logger.info(f"Main sheet name (if provided): {main_sheet_name}")
    logger.info(
        f"PDF subtable page range: {pdf_subtable_start_page} to {pdf_subtable_end_page}")

    # Validate file types
    if not pdf_file.filename.lower().endswith('.pdf'):
        logger.error(f"Invalid PDF file: {pdf_file.filename}")
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        logger.error(f"Invalid Excel file: {excel_file.filename}")
        raise HTTPException(status_code=400, detail="Excel file required")

    # Validate page range
    if pdf_subtable_start_page is not None and pdf_subtable_start_page < 1:
        raise HTTPException(
            status_code=400, detail="PDF subtable start page must be >= 1")
    if pdf_subtable_end_page is not None and pdf_subtable_end_page < 1:
        raise HTTPException(
            status_code=400, detail="PDF subtable end page must be >= 1")
    if (pdf_subtable_start_page is not None and pdf_subtable_end_page is not None and
            pdf_subtable_start_page > pdf_subtable_end_page):
        raise HTTPException(
            status_code=400, detail="PDF subtable start page cannot be greater than end page")

    pdf_fd = None
    pdf_path = None

    try:
        logger.info("=== PROCESSING PDF FILE FOR SUBTABLES ===")
        # Handle PDF file (requires temporary file)
        pdf_fd, pdf_path = tempfile.mkstemp(suffix='.pdf', prefix='subtable_')

        # Save PDF to temporary file
        pdf_content = await pdf_file.read()
        logger.info(f"PDF file size: {len(pdf_content)} bytes")

        with os.fdopen(pdf_fd, 'wb') as f:
            f.write(pdf_content)
        pdf_fd = None  # File descriptor is closed now

        logger.info("=== PROCESSING EXCEL FILE FOR SUBTABLES ===")
        # Handle Excel file (in-memory processing)
        excel_content = await excel_file.read()
        logger.info(f"Excel file size: {len(excel_content)} bytes")

        # Create in-memory buffer for Excel
        excel_buffer = BytesIO(excel_content)

        logger.info("=== STARTING SUBTABLE EXTRACTION PROCESS ===")
        # Parse files for subtables
        pdf_parser = PDFParser()
        excel_parser = ExcelParser()

        # First, extract the main table to get reference numbers from 摘要 column
        logger.info("Extracting main table to get reference numbers...")
        excel_table_extractor = ExcelTableExtractorService()
        main_table_items = excel_table_extractor.extract_main_table_from_buffer(
            excel_buffer, main_sheet_name)
        logger.info(
            f"Extracted {len(main_table_items)} main table items for reference discovery")

        # Extract reference numbers from main table's 摘要 column
        reference_numbers = []
        for item in main_table_items:
            remarks = item.raw_fields.get('摘要', '')
            if remarks and remarks.strip():
                reference_numbers.append(remarks.strip())

        # Remove duplicates and filter valid reference numbers
        unique_references = list(set(reference_numbers))
        valid_references = [
            ref for ref in unique_references if ref and ('号' in ref or '単' in ref)]

        logger.info(
            f"Found {len(valid_references)} unique reference numbers from main table: {valid_references[:10]}...")

        # IMPORTANT: If no valid references found but page range is specified,
        # we should still only process the specified page range without discovery
        if not valid_references and (pdf_subtable_start_page is not None or pdf_subtable_end_page is not None):
            logger.warning(
                "No valid reference numbers found from main table, but page range is specified.")
            logger.warning(
                "Will process specified page range without reference filtering.")
            # Use None to indicate no specific reference filtering, but still respect page range
            valid_references = None
        elif not valid_references:
            logger.warning(
                "No valid reference numbers found from main table. PDF extraction may use discovery mode.")

        # Extract subtables from PDF with page range and reference numbers
        logger.info("Extracting subtables from PDF with reference numbers...")
        pdf_subtables = pdf_parser.extract_subtables_with_range(
            pdf_path, pdf_subtable_start_page, pdf_subtable_end_page, valid_references)
        logger.info(f"Total PDF subtables extracted: {len(pdf_subtables)}")

        # Extract subtables from Excel using the NEW API-ready logic
        logger.info(
            "Extracting subtables from Excel using NEW API-ready logic...")
        excel_subtables = excel_table_extractor.extract_subtables_from_buffer(
            excel_buffer, main_sheet_name, main_table_items)
        logger.info(f"NEW API extracted {len(excel_subtables)} Excel subtables")

        # Close Excel buffer
        excel_buffer.close()

        # Create response with both PDF and Excel subtables
        result = SubtableComparisonSummary(
            total_pdf_subtables=len(pdf_subtables),
            total_excel_subtables=len(excel_subtables),
            pdf_subtables=pdf_subtables,
            excel_subtables=excel_subtables
        )

        logger.info("=== SUBTABLE EXTRACTION COMPLETED ===")
        logger.info(f"Summary: {len(pdf_subtables)} PDF subtables extracted")

        # Clean up references
        del pdf_parser, excel_parser, pdf_subtables, excel_buffer
        gc.collect()

        return result

    except HTTPException:
        # Re-raise HTTP exceptions
        raise
    except Exception as e:
        logger.error(
            f"Error during subtable comparison process: {str(e)}", exc_info=True)
        gc.collect()
        raise HTTPException(
            status_code=500, detail=f"Subtable processing error: {str(e)}")

    finally:
        # Clean up PDF temporary file
        try:
            if pdf_fd is not None:
                os.close(pdf_fd)
        except Exception as e:
            logger.error(f"Error closing PDF file descriptor: {e}")

        gc.collect()

        # Remove PDF temporary file
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.unlink(pdf_path)
                logger.info("Temporary PDF file cleaned up")
            except Exception as e:
                logger.warning(f"Could not remove temporary PDF file: {e}")
                try:
                    import time
                    time.sleep(0.5)
                    os.unlink(pdf_path)
                    logger.info("Temporary PDF file cleaned up (retry)")
                except Exception as e2:
                    logger.error(
                        f"Failed to clean up temporary PDF file: {e2}")


@router.post("/debug-matching")
async def debug_matching(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    start_page: Optional[int] = Form(None),
    end_page: Optional[int] = Form(None),
    sheet_name: Optional[str] = Form(None)
):
    """
    Debug endpoint to analyze why specific items are not being matched.
    Shows detailed information about normalization and matching process.
    """
    logger.info("=== STARTING DEBUG MATCHING ANALYSIS ===")

    if not pdf_file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Excel file required")

    pdf_fd = None
    pdf_path = None

    try:
        # Process files
        pdf_fd, pdf_path = tempfile.mkstemp(suffix='.pdf', prefix='debug_')
        pdf_content = await pdf_file.read()

        with os.fdopen(pdf_fd, 'wb') as f:
            f.write(pdf_content)
        pdf_fd = None

        excel_content = await excel_file.read()
        excel_buffer = BytesIO(excel_content)

        # Parse files
        pdf_parser = PDFParser()
        excel_parser = ExcelParser()

        pdf_items = pdf_parser.extract_tables_with_range(
            pdf_path, start_page, end_page)

        if sheet_name:
            excel_table_extractor = ExcelTableExtractorService()
            excel_items = excel_table_extractor.extract_main_table_from_buffer(
                excel_buffer, sheet_name)
        else:
            excel_items = excel_parser.extract_items_from_buffer_with_sheet(
                excel_buffer, sheet_name)

        excel_buffer.close()

        # Create matcher and normalizer for analysis
        matcher = Matcher()
        normalizer = Normalizer()

        # Debug items of interest
        debug_items = [
            "現場溶接 すみ肉溶接6mm換算",
            "現場孔明",
            "芯出し調整"
        ]

        debug_results = []

        for debug_item in debug_items:
            # Find this item in PDF and Excel
            pdf_matches = [
                item for item in pdf_items if debug_item in item.item_key]
            excel_matches = [
                item for item in excel_items if debug_item in item.item_key]

            for pdf_item in pdf_matches:
                # Normalize PDF item
                pdf_normalized_key = normalizer.normalize_item(
                    pdf_item.item_key)

                debug_info = {
                    "search_term": debug_item,
                    "pdf_item": {
                        "original_key": pdf_item.item_key,
                        "normalized_key": pdf_normalized_key,
                        "quantity": pdf_item.quantity,
                        "unit": pdf_item.unit,
                        "raw_fields": pdf_item.raw_fields
                    },
                    "excel_candidates": [],
                    "matching_analysis": {}
                }

                # Check against all Excel items for potential matches
                for excel_item in excel_items:
                    excel_normalized_key = normalizer.normalize_item(
                        excel_item.item_key)

                    # Check if names are similar
                    if (debug_item in excel_item.item_key or
                        any(word in excel_item.item_key for word in debug_item.split()) or
                            any(word in debug_item for word in excel_item.item_key.split())):

                        # Calculate similarity
                        are_different = normalizer.are_items_significantly_different(
                            pdf_normalized_key, excel_normalized_key)

                        candidate = {
                            "original_key": excel_item.item_key,
                            "normalized_key": excel_normalized_key,
                            "quantity": excel_item.quantity,
                            "unit": excel_item.unit,
                            "are_significantly_different": are_different,
                            "exact_match": pdf_normalized_key == excel_normalized_key
                        }

                        debug_info["excel_candidates"].append(candidate)

                # Simulate the matching process
                pdf_normalized = matcher._normalize_items(pdf_items, "PDF")
                excel_normalized = matcher._normalize_items(
                    excel_items, "Excel")

                matched_excel_keys = set()
                comparison_result = matcher._compare_single_pdf_item(
                    pdf_normalized_key, pdf_item, excel_normalized, matched_excel_keys
                )

                debug_info["matching_analysis"] = {
                    "final_status": comparison_result.status,
                    "match_confidence": comparison_result.match_confidence,
                    "matched_excel_item": {
                        "key": comparison_result.excel_item.item_key if comparison_result.excel_item else None,
                        "quantity": comparison_result.excel_item.quantity if comparison_result.excel_item else None,
                        "unit": comparison_result.excel_item.unit if comparison_result.excel_item else None
                    } if comparison_result.excel_item else None
                }

                debug_results.append(debug_info)

        return {
            "debug_results": debug_results,
            "total_pdf_items": len(pdf_items),
            "total_excel_items": len(excel_items),
            "normalization_settings": {
                "min_confidence_threshold": matcher.min_confidence,
                "normalizer_info": "Check logs for detailed normalization process"
            }
        }

    except Exception as e:
        logger.error(f"Error in debug matching: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Debug error: {str(e)}")

    finally:
        # Cleanup
        try:
            if pdf_fd is not None:
                os.close(pdf_fd)
        except Exception:
            pass

        if pdf_path and os.path.exists(pdf_path):
            try:
                os.unlink(pdf_path)
            except Exception:
                pass

        gc.collect()


@router.post("/compare-matching-methods")
async def compare_matching_methods(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    start_page: Optional[int] = Form(None),
    end_page: Optional[int] = Form(None),
    sheet_name: Optional[str] = Form(None)
):
    """
    Compare old vs new matching methods to see the difference in extra items detection.
    This helps debug why items that should be matched are appearing as extra.
    """
    logger.info("=== COMPARING OLD VS NEW MATCHING METHODS ===")

    if not pdf_file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Excel file required")

    pdf_fd = None
    pdf_path = None

    try:
        # Process files
        pdf_fd, pdf_path = tempfile.mkstemp(suffix='.pdf', prefix='compare_')
        pdf_content = await pdf_file.read()

        with os.fdopen(pdf_fd, 'wb') as f:
            f.write(pdf_content)
        pdf_fd = None

        excel_content = await excel_file.read()
        excel_buffer = BytesIO(excel_content)

        # Parse files
        pdf_parser = PDFParser()
        pdf_items = pdf_parser.extract_tables_with_range(
            pdf_path, start_page, end_page)

        if sheet_name:
            excel_table_extractor = ExcelTableExtractorService()
            excel_items = excel_table_extractor.extract_main_table_from_buffer(
                excel_buffer, sheet_name)
        else:
            excel_parser = ExcelParser()
            excel_items = excel_parser.extract_items_from_buffer_with_sheet(
                excel_buffer, sheet_name)

        excel_buffer.close()

        # Test both matching methods
        matcher = Matcher()

        # Old method (complex normalization + strict matching)
        old_extra_items = matcher.get_extra_items_only(pdf_items, excel_items)

        # New method (simplified matching focusing on name, quantity, unit)
        new_extra_items = matcher.get_extra_items_only_simplified(
            pdf_items, excel_items)

        # Find the problematic items mentioned by user
        problem_items = ["現場溶接 すみ肉溶接6mm換算", "現場孔明", "芯出し調整"]

        analysis = {
            "problem_items_analysis": [],
            "old_method_results": {
                "extra_items_count": len(old_extra_items),
                "extra_items": [{"item_key": item.item_key, "quantity": item.quantity, "unit": item.unit} for item in old_extra_items]
            },
            "new_method_results": {
                "extra_items_count": len(new_extra_items),
                "extra_items": [{"item_key": item.item_key, "quantity": item.quantity, "unit": item.unit} for item in new_extra_items]
            },
            "improvement_summary": {}
        }

        # Analyze each problem item
        for problem_item in problem_items:
            # Check if this item appears in old extra items but not in new extra items
            old_has_item = any(
                problem_item in item.item_key for item in old_extra_items)
            new_has_item = any(
                problem_item in item.item_key for item in new_extra_items)

            # Find matching items in PDF and Excel
            pdf_matches = [
                item for item in pdf_items if problem_item in item.item_key]
            excel_matches = [
                item for item in excel_items if problem_item in item.item_key]

            analysis["problem_items_analysis"].append({
                "search_term": problem_item,
                "found_in_pdf": len(pdf_matches) > 0,
                "found_in_excel": len(excel_matches) > 0,
                "pdf_items": [{"key": item.item_key, "qty": item.quantity, "unit": item.unit} for item in pdf_matches],
                "excel_items": [{"key": item.item_key, "qty": item.quantity, "unit": item.unit} for item in excel_matches],
                "old_method_shows_as_extra": old_has_item,
                "new_method_shows_as_extra": new_has_item,
                "improvement": "FIXED" if old_has_item and not new_has_item else "NO_CHANGE" if old_has_item == new_has_item else "REGRESSION"
            })

        # Summary of improvements
        items_fixed = sum(
            1 for item in analysis["problem_items_analysis"] if item["improvement"] == "FIXED")
        items_regressed = sum(
            1 for item in analysis["problem_items_analysis"] if item["improvement"] == "REGRESSION")

        analysis["improvement_summary"] = {
            "total_extra_items_old": len(old_extra_items),
            "total_extra_items_new": len(new_extra_items),
            "reduction_in_extra_items": len(old_extra_items) - len(new_extra_items),
            "problem_items_fixed": items_fixed,
            "problem_items_regressed": items_regressed,
            "recommendation": "USE_NEW_METHOD" if items_fixed > items_regressed else "INVESTIGATE_FURTHER"
        }

        return analysis

    except Exception as e:
        logger.error(f"Error in matching comparison: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Comparison error: {str(e)}")

    finally:
        # Cleanup
        try:
            if pdf_fd is not None:
                os.close(pdf_fd)
        except Exception:
            pass

        if pdf_path and os.path.exists(pdf_path):
            try:
                os.unlink(pdf_path)
            except Exception:
                pass

        gc.collect()
