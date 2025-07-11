from fastapi import APIRouter, UploadFile, File, HTTPException, Form
from typing import List, Optional
import tempfile
import os
import gc
import logging
from io import BytesIO
from ..schemas.tender import ComparisonSummary, SubtableComparisonSummary
from ..services.pdf_parser import PDFParser
from ..services.excel_parser import ExcelParser
from ..services.matcher import Matcher

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

router = APIRouter()


@router.get("/test")
async def test_endpoint():
    """Simple test endpoint to verify the server is working"""
    return {"status": "ok", "message": "Tender API is working"}


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


@router.post("/compare-missing-only")
async def compare_tender_files_missing_only(
    pdf_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    start_page: Optional[int] = Form(None),
    end_page: Optional[int] = Form(None),
    sheet_name: Optional[str] = Form(None)
):
    """
    Compare tender files and return only the missing items (PDF items not in Excel).
    Optimized endpoint for specific use case.

    Args:
        pdf_file: PDF tender document
        excel_file: Excel proposal document
        start_page: Starting page number for PDF extraction (optional)
        end_page: Ending page number for PDF extraction (optional)
        sheet_name: Specific Excel sheet name to extract from (optional)
    """
    logger.info("=== STARTING MISSING ITEMS COMPARISON ===")
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
        excel_items = excel_parser.extract_items_from_buffer_with_sheet(
            excel_buffer, sheet_name)
        excel_buffer.close()

        # Get only missing items
        matcher = Matcher()
        missing_items = matcher.get_missing_items_only(pdf_items, excel_items)

        # Return simplified response
        return {
            "total_pdf_items": len(pdf_items),
            "total_excel_items": len(excel_items),
            "missing_items_count": len(missing_items),
            "missing_items": [
                {
                    "item_key": item.item_key,
                    "raw_fields": item.raw_fields,
                    "quantity": item.quantity,
                    "unit": item.unit,
                    "page_number": item.page_number
                }
                for item in missing_items
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
    sheet_name: Optional[str] = Form(None)
):
    """
    Compare tender files and return only the quantity mismatches.
    Optimized endpoint for specific use case.

    Args:
        pdf_file: PDF tender document
        excel_file: Excel proposal document
        start_page: Starting page number for PDF extraction (optional)
        end_page: Ending page number for PDF extraction (optional)
        sheet_name: Specific Excel sheet name to extract from (optional)
    """
    logger.info("=== STARTING QUANTITY MISMATCHES COMPARISON ===")
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
        excel_items = excel_parser.extract_items_from_buffer_with_sheet(
            excel_buffer, sheet_name)
        excel_buffer.close()

        # Get only mismatched items
        matcher = Matcher()
        mismatched_results = matcher.get_mismatched_items_only(
            pdf_items, excel_items)
        unit_mismatched_results = matcher.get_unit_mismatched_items_only(
            pdf_items, excel_items)

        # Return simplified response
        return {
            "total_pdf_items": len(pdf_items),
            "total_excel_items": len(excel_items),
            "quantity_mismatches_count": len(mismatched_results),
            "quantity_mismatches": [
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
                    "quantity_difference": result.quantity_difference,
                    "match_confidence": result.match_confidence
                }
                for result in mismatched_results
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
        excel_items = excel_parser.extract_items_from_buffer_with_sheet(
            excel_buffer, sheet_name)
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
    logger.info(f"PDF subtable page range: {pdf_subtable_start_page} to {pdf_subtable_end_page}")
    
    # Validate file types
    if not pdf_file.filename.lower().endswith('.pdf'):
        logger.error(f"Invalid PDF file: {pdf_file.filename}")
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        logger.error(f"Invalid Excel file: {excel_file.filename}")
        raise HTTPException(status_code=400, detail="Excel file required")
    
    # Validate page range
    if pdf_subtable_start_page is not None and pdf_subtable_start_page < 1:
        raise HTTPException(status_code=400, detail="PDF subtable start page must be >= 1")
    if pdf_subtable_end_page is not None and pdf_subtable_end_page < 1:
        raise HTTPException(status_code=400, detail="PDF subtable end page must be >= 1")
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
        
        # Extract subtables from PDF with page range
        logger.info("Extracting subtables from PDF with specified parameters...")
        pdf_subtables = pdf_parser.extract_subtables_with_range(
            pdf_path, pdf_subtable_start_page, pdf_subtable_end_page)
        logger.info(f"Total PDF subtables extracted: {len(pdf_subtables)}")
        
        # Extract subtables from Excel
        logger.info("Extracting subtables from Excel…")
        excel_subtables = excel_parser.extract_subtables_from_buffer(
            excel_buffer, main_sheet_name=main_sheet_name)
        logger.info(f"Total Excel subtables extracted: {len(excel_subtables)}")
        
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
        logger.error(f"Error during subtable comparison process: {str(e)}", exc_info=True)
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
                    logger.error(f"Failed to clean up temporary PDF file: {e2}")



