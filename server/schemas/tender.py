from typing import Dict, List, Literal, Optional
from pydantic import BaseModel, Field


class TenderItem(BaseModel):
    item_key: str = Field(...,
                          description="Concatenated and normalized item key")
    raw_fields: Dict[str, str] = Field(...,
                                       description="Original fields like 工種, 種別, etc.")
    quantity: float = Field(..., description="Item quantity")
    unit: Optional[str] = Field(None, description="Unit of measurement (単位)")
    source: Literal["PDF",
                    "Excel"] = Field(..., description="Source document type")
    page_number: Optional[int] = Field(
        None, description="Page number where the item was found (for PDF items)")
    logical_line_number: Optional[int] = Field(
        None, description="Logical line index within the table (for Excel main)")
    table_number: Optional[str] = Field(
        None, description="Detected table number at the end of the table (for Excel main)")


class SubtableItem(BaseModel):
    item_key: str = Field(...,
                          description="Concatenated and normalized item key for subtable")
    raw_fields: Dict[str, str] = Field(...,
                                       description="Original fields like 名称・規格, 単位, 単数, 摘要")
    quantity: float = Field(..., description="Item quantity (単数)")
    unit: Optional[str] = Field(None, description="Unit of measurement (単位)")
    source: Literal["PDF",
                    "Excel"] = Field(..., description="Source document type")
    page_number: Optional[int] = Field(
        None, description="Page number where the item was found (for PDF items)")
    reference_number: Optional[str] = Field(
        None, description="Reference number like '単 3号' associated with this item")
    sheet_name: Optional[str] = Field(
        None, description="Excel sheet name where the item was found (for Excel items)")
    table_title: Optional[Dict[str, str]] = Field(
        None, description="Table title information with item_name, unit, and unit_quantity")
    logical_line_number: Optional[int] = Field(
        None, description="Logical item index per reference subtable (row-spanning aware)")


class ComparisonResult(BaseModel):
    status: Literal["OK", "QUANTITY_MISMATCH", "UNIT_MISMATCH", "NAME_MISMATCH", "MISSING",
                    "EXTRA"] = Field(..., description="Comparison status")
    pdf_item: Optional[TenderItem] = Field(
        None, description="PDF item if present")
    excel_item: Optional[TenderItem] = Field(
        None, description="Excel item if present")
    match_confidence: float = Field(...,
                                    description="Confidence score for the match (0-1)")
    quantity_difference: Optional[float] = Field(
        None, description="Difference in quantities if applicable")
    unit_mismatch: Optional[bool] = Field(
        None, description="Whether units are different")
    type: Literal["Main Table", "Sub Table"] = Field(
        ..., description="Indicates if the result is from the main table or sub table")


class ComparisonSummary(BaseModel):
    total_items: int = Field(..., description="Total number of items compared")
    matched_items: int = Field(...,
                               description="Number of exactly matched items")
    quantity_mismatches: int = Field(...,
                                     description="Number of items with quantity differences")
    unit_mismatches: int = Field(...,
                                 description="Number of items with unit differences")
    missing_items: int = Field(...,
                               description="Number of items missing in Excel")
    extra_items: int = Field(..., description="Number of extra items in Excel")
    results: List[ComparisonResult] = Field(...,
                                            description="Detailed comparison results")


class SubtableComparisonSummary(BaseModel):
    total_pdf_subtables: int = Field(...,
                                     description="Total number of PDF subtable items")
    total_excel_subtables: int = Field(...,
                                       description="Total number of Excel subtable items")
    pdf_subtables: List[SubtableItem] = Field(
        ..., description="PDF subtable items")
    excel_subtables: List[SubtableItem] = Field(
        ..., description="Excel subtable items")


class SubtableComparisonResult(BaseModel):
    status: Literal["OK", "QUANTITY_MISMATCH", "UNIT_MISMATCH", "NAME_MISMATCH", "MISSING",
                    "EXTRA"] = Field(..., description="Comparison status for subtable")
    pdf_item: Optional[SubtableItem] = Field(
        None, description="PDF subtable item if present")
    excel_item: Optional[SubtableItem] = Field(
        None, description="Excel subtable item if present")
    match_confidence: float = Field(...,
                                    description="Confidence score for the match (0-1)")
    quantity_difference: Optional[float] = Field(
        None, description="Difference in quantities if applicable")
    unit_mismatch: Optional[bool] = Field(
        None, description="Whether units are different")
    type: Literal["Sub Table"] = Field(
        "Sub Table", description="Indicates this result is from a sub table")
