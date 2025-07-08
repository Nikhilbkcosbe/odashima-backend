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


class ComparisonResult(BaseModel):
    status: Literal["OK", "QUANTITY_MISMATCH", "UNIT_MISMATCH", "MISSING",
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
