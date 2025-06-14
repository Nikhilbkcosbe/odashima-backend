from fastapi import APIRouter, UploadFile, File, HTTPException
from typing import List
import tempfile
import os
import gc
import logging
from io import BytesIO
from ..schemas.tender import ComparisonSummary
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
    excel_file: UploadFile = File(...)
) -> ComparisonSummary:
    """
    Compare a tender PDF with an Excel proposal iteratively.
    Focus on finding mismatches and items present in PDF but not in Excel.
    """
    logger.info("=== STARTING TENDER COMPARISON ===")
    logger.info(f"PDF file: {pdf_file.filename}")
    logger.info(f"Excel file: {excel_file.filename}")

    # Validate file types
    if not pdf_file.filename.lower().endswith('.pdf'):
        logger.error(f"Invalid PDF file: {pdf_file.filename}")
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        logger.error(f"Invalid Excel file: {excel_file.filename}")
        raise HTTPException(status_code=400, detail="Excel file required")

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
        # Parse files iteratively
        pdf_parser = PDFParser()
        excel_parser = ExcelParser()

        # Extract items page by page from PDF
        logger.info(
            "Extracting items from PDF (page by page, table by table)...")
        pdf_items = pdf_parser.extract_tables(pdf_path)
        logger.info(f"Total PDF items extracted: {len(pdf_items)}")

        # Extract items sheet by sheet from Excel
        logger.info("Extracting items from Excel (sheet by sheet)...")
        excel_items = excel_parser.extract_items_from_buffer(excel_buffer)
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
    excel_file: UploadFile = File(...)
):
    """
    Compare tender files and return only the missing items (PDF items not in Excel).
    Optimized endpoint for specific use case.
    """
    logger.info("=== STARTING MISSING ITEMS COMPARISON ===")

    # Same file processing as main endpoint
    if not pdf_file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Excel file required")

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

        # Parse files
        pdf_parser = PDFParser()
        excel_parser = ExcelParser()

        pdf_items = pdf_parser.extract_tables(pdf_path)
        excel_items = excel_parser.extract_items_from_buffer(excel_buffer)
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
                    "quantity": item.quantity
                }
                for item in missing_items
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
    excel_file: UploadFile = File(...)
):
    """
    Compare tender files and return only the quantity mismatches.
    Optimized endpoint for specific use case.
    """
    logger.info("=== STARTING QUANTITY MISMATCHES COMPARISON ===")

    # Same file processing as main endpoint
    if not pdf_file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Excel file required")

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

        # Parse files
        pdf_parser = PDFParser()
        excel_parser = ExcelParser()

        pdf_items = pdf_parser.extract_tables(pdf_path)
        excel_items = excel_parser.extract_items_from_buffer(excel_buffer)
        excel_buffer.close()

        # Get only mismatched items
        matcher = Matcher()
        mismatched_results = matcher.get_mismatched_items_only(
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
                        "raw_fields": result.pdf_item.raw_fields
                    },
                    "excel_item": {
                        "item_key": result.excel_item.item_key,
                        "quantity": result.excel_item.quantity,
                        "raw_fields": result.excel_item.raw_fields
                    },
                    "quantity_difference": result.quantity_difference,
                    "match_confidence": result.match_confidence
                }
                for result in mismatched_results
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
