# Estimate Reference PDF Extraction Implementation

## Overview

This implementation adds functionality to extract specific information from the 見積積算参考資料 (Estimate Reference) PDF file, specifically from the 2nd page, and integrates it with the existing 特記仕様書 (Specification) extraction system.

## New Features

### 1. EstimateReferenceExtractor Service

**File**: `backend/server/services/estimate_extractor.py`

A new service class that extracts the following information from the 2nd page of the estimate reference PDF:

- **省庁** (Government Agency): Extracted from footer position, looks for patterns ending with 県, 都, 府, 市
- **年度** (Fiscal Year): Extracted using Japanese era patterns (令和, 平成, etc.) + number + 年度
- **経費工種** (Expense Type): Extracted from text following "工種区分"
- **施工地域工事場所** (Construction Location): Extracted from text following "工事名"

### 2. Updated API Endpoint

**File**: `backend/server/api/tender.py`

The `/extract-spec` endpoint has been updated to:

- Accept two PDF files: `spec_pdf_file` and `estimate_pdf_file`
- Process both files using their respective extractors
- Return combined results including both spec sections and estimate information

### 3. Enhanced Frontend

**Files**:

- `frontend/src/pages/TenderUpload.jsx`
- `frontend/src/pages/SpecExtractionResults.jsx`

#### TenderUpload.jsx Changes:

- Added validation for both required PDF files
- Enhanced file detection based on filename patterns
- Added visual indicators showing which files are detected
- Updated submit button to be disabled until both files are selected
- Added validation warnings for missing files

#### SpecExtractionResults.jsx Changes:

- Added new section to display estimate information
- Shows both spec sections and estimate info in a unified view
- Enhanced file display to show both uploaded files

## Technical Details

### File Detection Patterns

The system automatically detects files based on these patterns:

**特記仕様書 (Specification)**:

- `特記仕様書`
- `tokki`
- `spec`

**見積積算参考資料 (Estimate Reference)**:

- `見積`
- `積算参考資料`
- `estimate`
- `reference`

### Extraction Logic

#### 省庁 (Government Agency)

```python
shocho_patterns = [
    r'([^\s]+[県都府市])',  # Any text ending with 県, 都, 府, 市
    r'([^\s]+県[^\s]*)',    # Text containing 県
    r'([^\s]+都[^\s]*)',    # Text containing 都
    r'([^\s]+府[^\s]*)',    # Text containing 府
    r'([^\s]+市[^\s]*)',    # Text containing 市
]
```

#### 年度 (Fiscal Year)

```python
nendo_patterns = [
    r'([令和平成昭和大正明治]\s*\d+\s*年度)',  # Full pattern with era
    r'(\d+\s*年度)',                           # Just number + 年度
    r'([令和平成昭和大正明治]\s*\d+)',         # Era + number
]
```

#### 経費工種 (Expense Type)

```python
keihi_patterns = [
    r'工種区分[^\n]*?([^\n]+)',  # After 工種区分, get the next text
    r'工種区分\s*([^\n\s]+)',    # Directly after 工種区分
]
```

#### 施工地域工事場所 (Construction Location)

```python
kouji_patterns = [
    r'工\s*事\s*名[^\n]*?([^\n]+)',  # After 工事名, get the next text
    r'工\s*事\s*名\s*([^\n\s]+)',    # Directly after 工事名
]
```

### API Response Format

The updated API now returns:

```json
{
  "sections": [
    {
      "section": "第２条",
      "data": { ... }
    },
    ...
  ],
  "spec_filename": "特記仕様書.pdf",
  "estimate_filename": "見積積算参考資料.pdf",
  "省庁": "東京都",
  "年度": "令和7年度",
  "経費工種": "土木工事",
  "施工地域工事場所": "水沢橋工事"
}
```

## Usage

### Backend Testing

Run the test script to verify extraction:

```bash
cd backend
python test_estimate_extractor.py
```

### Frontend Usage

1. Navigate to the 特記仕様書 tab
2. Upload both required PDF files:
   - 特記仕様書 PDF
   - 見積積算参考資料 PDF
3. The system will automatically detect and validate the files
4. Click "特記仕様書 抽出" to process both files
5. View results showing both spec sections and estimate information

## Error Handling

- **File Validation**: Ensures both required PDF files are uploaded
- **Pattern Matching**: Gracefully handles cases where information cannot be found
- **Text Cleaning**: Normalizes full-width/half-width spaces and removes extra whitespace
- **Logging**: Comprehensive logging for debugging extraction issues

## Dependencies

- `pdfplumber`: For PDF text extraction
- `re`: For pattern matching
- `logging`: For debugging and monitoring

## Future Enhancements

1. **Improved Pattern Matching**: Add more sophisticated patterns for better extraction accuracy
2. **OCR Support**: Add OCR capabilities for scanned PDFs
3. **Validation Rules**: Add business logic validation for extracted values
4. **Export Features**: Add ability to export extracted data to Excel/CSV
