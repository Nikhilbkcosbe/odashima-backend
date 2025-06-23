# Custom Column Name Feature

## Overview

The Construction Tender vs Proposal Reconciliation System now supports **dynamic column name specification** for item identification. Instead of relying solely on hardcoded column patterns, users can specify which column contains the item names.

## What Changed

### Before

- Item names were identified using hardcoded patterns:
  - `工事区分・工種・種別・細別`
  - `工事区分`, `工種`, `種別`, `細別`, `費目`
  - `規格`, `名称・規格`, `名称`, `項目`, `品名`
  - `摘要`, `備考`

### After

- ✅ **Custom column name can be specified via API parameter**
- ✅ **Custom column takes priority over default patterns**
- ✅ **Falls back to default patterns when custom column is not provided or empty**
- ✅ **All other column names remain the same (数量, 単位, 単価, 金額, etc.)**

## API Changes

### New Parameter

All comparison endpoints now accept an optional `item_name_column` parameter:

```bash
# Basic usage with custom column
curl -X POST "http://localhost:8000/api/v1/tender/compare" \
  -F "pdf_file=@tender.pdf" \
  -F "excel_file=@proposal.xlsx" \
  -F "item_name_column=項目名"

# Usage with all parameters
curl -X POST "http://localhost:8000/api/v1/tender/compare" \
  -F "pdf_file=@tender.pdf" \
  -F "excel_file=@proposal.xlsx" \
  -F "start_page=1" \
  -F "end_page=10" \
  -F "sheet_name=Sheet1" \
  -F "item_name_column=カスタム項目名"
```

### Supported Endpoints

1. **`/api/v1/tender/compare`** - Full comparison with custom column support
2. **`/api/v1/tender/compare-missing-only`** - Missing items only with custom column support
3. **`/api/v1/tender/compare-mismatches-only`** - Quantity mismatches only with custom column support

### Parameters

| Parameter          | Type    | Required | Description                                         |
| ------------------ | ------- | -------- | --------------------------------------------------- |
| `pdf_file`         | File    | ✅       | PDF tender document                                 |
| `excel_file`       | File    | ✅       | Excel proposal document                             |
| `start_page`       | Integer | ❌       | Starting page number for PDF extraction             |
| `end_page`         | Integer | ❌       | Ending page number for PDF extraction               |
| `sheet_name`       | String  | ❌       | Specific Excel sheet name to extract from           |
| `item_name_column` | String  | ❌       | **NEW:** Custom column name for item identification |

## Implementation Details

### Architecture Changes

#### PDF Parser (`pdf_parser.py`)

- Added `set_custom_item_name_column()` method
- Updated `extract_tables_with_range()` to accept `item_name_column` parameter
- Modified `_has_item_identifying_fields()` to prioritize custom column
- Updated `_create_item_key_from_fields()` to use custom column first

#### Excel Parser (`excel_parser.py`)

- Added `set_custom_item_name_column()` method
- Updated `extract_items_from_buffer_with_sheet()` to accept `item_name_column` parameter
- Modified `_has_item_identifying_fields()` to prioritize custom column
- Updated `_create_item_key_from_fields()` to use custom column first

#### API Endpoints (`tender.py`)

- Added `item_name_column: Optional[str] = Form(None)` to all endpoints
- Updated parser method calls to pass the custom column parameter
- Enhanced logging to show when custom columns are being used

### Logic Flow

1. **API receives request** with optional `item_name_column` parameter
2. **Parsers are initialized** with custom column name (if provided)
3. **Column patterns are updated** to prioritize custom column
4. **During extraction:**
   - Custom column is checked first for item identification
   - Falls back to default patterns if custom column is empty/missing
   - All other columns (quantity, unit, price) remain unchanged
5. **Item matching** uses the dynamically identified item names

### Backward Compatibility

- ✅ **100% backward compatible** - existing API calls work unchanged
- ✅ **Default behavior preserved** when `item_name_column` is not specified
- ✅ **No breaking changes** to existing data models or response formats

## Usage Examples

### Example 1: Default Behavior (No Changes)

```bash
# This works exactly as before
curl -X POST "http://localhost:8000/api/v1/tender/compare" \
  -F "pdf_file=@tender.pdf" \
  -F "excel_file=@proposal.xlsx"
```

### Example 2: Custom Item Column

```bash
# Now you can specify which column contains item names
curl -X POST "http://localhost:8000/api/v1/tender/compare" \
  -F "pdf_file=@tender.pdf" \
  -F "excel_file=@proposal.xlsx" \
  -F "item_name_column=項目名称"
```

### Example 3: Custom Column with Page Range

```bash
# Combine custom column with other parameters
curl -X POST "http://localhost:8000/api/v1/tender/compare" \
  -F "pdf_file=@tender.pdf" \
  -F "excel_file=@proposal.xlsx" \
  -F "start_page=5" \
  -F "end_page=15" \
  -F "item_name_column=工事項目"
```

## Benefits

1. **✅ Flexibility** - Works with documents that use non-standard column names
2. **✅ Accuracy** - Users can specify exactly which column contains item names
3. **✅ Compatibility** - Maintains full backward compatibility with existing usage
4. **✅ Simplicity** - Single parameter addition, no complex configuration needed
5. **✅ Robustness** - Falls back gracefully to default patterns when needed

## Testing

A test script `test_custom_column.py` is provided to verify the functionality:

```bash
cd backend
python test_custom_column.py
```

This script demonstrates:

- Custom column setting and retrieval
- Item identification with custom columns
- Fallback behavior to default patterns
- Integration between PDF and Excel parsers

## Column Priority Order

When identifying item names, the system now uses this priority:

1. **Custom column name** (if specified and contains data)
2. **工事区分・工種・種別・細別** (primary default pattern)
3. **摘要** (secondary default pattern)
4. **Other default patterns** (規格, etc.)

## Logging

Enhanced logging shows when custom columns are being used:

```
INFO:pdf_parser:PDF Parser: Set custom item name column to '項目名'
INFO:excel_parser:Excel Parser: Set custom item name column to '項目名'
INFO:tender:Custom item name column: 項目名
```

## Future Enhancements

Potential future improvements could include:

- Multiple custom column support
- Column mapping configuration files
- UI for column selection
- Automatic column detection

---

**This feature maintains complete backward compatibility while adding powerful new functionality for handling diverse document formats.**
