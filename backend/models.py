from typing import Optional, List, Literal
from pydantic import BaseModel, Field
from datetime import date

Region = Literal["CPI Northern", "CPI Southern", "Unassigned"]

class Salesman(BaseModel):
    name: str = Field(..., description="Satışçı ismi")
    region: Region = Field("Unassigned", description="Bölge ataması")

class SalesmanList(BaseModel):
    items: List[Salesman]

class Record(BaseModel):
    # Zorunlu alanlar
    date_of_request: date
    salesman: str
    customer_name: str
    customer_po_no: str
    salesforce_reference: str
    so_no: str
    amount_eur: float
    total_discount_pct: float
    cpi_eur: float
    cps_eur: float

    # Opsiyonel alanlar
    definition: Optional[str] = ""
    date_of_delivery: Optional[date] = None
    date_of_invoice: Optional[date] = None
    note: Optional[str] = ""

    # Sistem alanları
    record_id: Optional[str] = None  # UUID atanacak

class RecordList(BaseModel):
    items: List[Record]

class LookupQuery(BaseModel):
    so_no: Optional[str] = None
    customer_po_no: Optional[str] = None

class LLMParseRequest(BaseModel):
    email_text: str

class LLMParseResponse(BaseModel):
    suggested: Record
    confidence: float = 0.0
    message: str = "LLM parsingi burada devreye alınabilir."
