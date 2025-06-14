from fastapi import APIRouter, UploadFile, File, HTTPException
from typing import List
import tempfile
import os
import gc
from io import BytesIO
from ..schemas.tender import ComparisonSummary
from ..services.pdf_parser import PDFParser
from ..services.excel_parser import ExcelParser
from ..services.matcher import Matcher

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
    Compare a tender PDF with an Excel proposal.
    """
    # Validate file types
    if not pdf_file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="PDF file required")
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Excel file required")

    # For PDF, we still need a temporary file (pdfplumber requires file path)
    # For Excel, we can use in-memory processing
    pdf_fd = None
    pdf_path = None
    
    try:
        # Handle PDF file (requires temporary file)
        pdf_fd, pdf_path = tempfile.mkstemp(suffix='.pdf', prefix='tender_')

        # Save PDF to temporary file
        pdf_content = await pdf_file.read()
        
        with os.fdopen(pdf_fd, 'wb') as f:
            f.write(pdf_content)
        pdf_fd = None  # File descriptor is closed now

        # Handle Excel file (in-memory processing)
        excel_content = await excel_file.read()
        
        # Create in-memory buffer for Excel
        excel_buffer = BytesIO(excel_content)

        # Parse files
        pdf_parser = PDFParser()
        excel_parser = ExcelParser()

        pdf_items = pdf_parser.extract_tables(pdf_path)
        excel_items = excel_parser.extract_items_from_buffer(excel_buffer)

        # Close Excel buffer
        excel_buffer.close()

        # Compare items
        matcher = Matcher()
        result = matcher.compare_items(pdf_items, excel_items)

        # Clean up references
        del pdf_parser, excel_parser, matcher, pdf_items, excel_items, excel_buffer
        gc.collect()

        return result

    except Exception as e:
        gc.collect()
        raise HTTPException(status_code=500, detail=str(e))
        
    finally:
        # Clean up PDF temporary file
        try:
            if pdf_fd is not None:
                os.close(pdf_fd)
        except Exception:
            pass
        
        gc.collect()
        
        # Remove PDF temporary file
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.unlink(pdf_path)
            except Exception:
                try:
                    import time
                    time.sleep(0.5)
                    os.unlink(pdf_path)
                except Exception:
                    pass
