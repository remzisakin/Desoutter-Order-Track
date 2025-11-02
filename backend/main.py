import os
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from typing import Optional

from .models import (
    Salesman,
    SalesmanList,
    Record,
    RecordList,
    LookupQuery,
    LLMParseRequest,
    LLMParseResponse,
)
from . import excel_store as store

app = FastAPI(title="Desoutter Order Track API", version="1.0.0")


@app.get("/")
def root():
    """Basit bir sağlık kontrolü döndürür."""
    return {"status": "ok"}

# CORS: frontend (Streamlit) için
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------- Salesman/Data ----------

@app.get("/data/salesmen", response_model=SalesmanList)
def get_salesmen():
    df = store.list_salesmen()
    items = [{"name": r["SalesMan"], "region": r["Region"]} for _, r in df.iterrows()]
    return {"items": items}

@app.post("/data/salesmen", response_model=Salesman)
def add_or_update_salesman(s: Salesman):
    store.upsert_salesman(s.name, s.region)
    return s

@app.post("/data/salesmen/bulk", response_model=SalesmanList)
def set_salesmen_bulk(lst: SalesmanList):
    store.bulk_set_salesmen([{"name": i.name, "region": i.region} for i in lst.items])
    return lst

# ---------- Records ----------

@app.get("/records", response_model=RecordList)
def list_records():
    df = store.list_records()
    items = []
    for _, r in df.iterrows():
        items.append(_row_to_record(r))
    return {"items": items}


@app.get("/orders", response_model=RecordList, include_in_schema=False)
def list_orders():
    """Legacy uyumluluğu için /records endpoint'inin takma adı."""
    return list_records()

@app.post("/records", response_model=Record)
def create_record(rec: Record):
    saved = store.create_record(rec.model_dump())
    return _row_to_record(saved)


@app.post("/orders", response_model=Record, include_in_schema=False)
def create_order(rec: Record):
    """Legacy uyumluluğu için /records endpoint'inin takma adı."""
    return create_record(rec)

@app.post("/records/lookup", response_model=Record)
def lookup_record(q: LookupQuery):
    row = store.find_record(so_no=q.so_no, customer_po_no=q.customer_po_no)
    if not row:
        raise HTTPException(status_code=404, detail="Kayıt bulunamadı")
    return _row_to_record(row)


@app.post("/orders/lookup", response_model=Record, include_in_schema=False)
def lookup_order(q: LookupQuery):
    """Legacy uyumluluğu için /records endpoint'inin takma adı."""
    return lookup_record(q)

@app.put("/records/{record_id}", response_model=Record)
def update_record(record_id: str, rec: Record):
    updated = store.update_record(record_id, rec.model_dump())
    if not updated:
        raise HTTPException(status_code=404, detail="Güncellenecek kayıt bulunamadı")
    return _row_to_record(updated)


@app.put("/orders/{record_id}", response_model=Record, include_in_schema=False)
def update_order(record_id: str, rec: Record):
    """Legacy uyumluluğu için /records endpoint'inin takma adı."""
    return update_record(record_id, rec)


@app.get("/records/export")
def export_records():
    path = store.get_excel_path()
    store.ensure_workbook_format()
    if not os.path.exists(path):
        raise HTTPException(status_code=500, detail="Excel dosyası bulunamadı")
    filename = os.path.basename(path)
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )


@app.get("/orders/export", include_in_schema=False)
def export_orders():
    """Legacy uyumluluğu için /records endpoint'inin takma adı."""
    return export_records()

# ---------- Reports ----------

@app.get("/reports/summary")
def get_reports():
    frames = store.report_frames()
    # Pandas DF -> JSON list
    def tolist(df):
        return [] if df is None or len(df) == 0 else df.to_dict(orient="records")
    return {k: tolist(v) for k, v in frames.items()}

# ---------- LLM Parse (stub) ----------

@app.post("/llm/parse", response_model=LLMParseResponse)
def llm_parse(req: LLMParseRequest):
    # Burada ileride OpenAI/LLM entegrasyonunu yapabilirsiniz.
    # Şimdilik basit bir şablon döndürüyoruz:
    from datetime import date
    sample = {
        "date_of_request": date.today(),
        "salesman": "",
        "customer_name": "",
        "customer_po_no": "",
        "salesforce_reference": "",
        "so_no": "",
        "amount_eur": 0.0,
        "total_discount_pct": 0.0,
        "cpi_eur": 0.0,
        "cps_eur": 0.0,
        "definition": "",
        "date_of_delivery": None,
        "date_of_invoice": None,
        "note": "",
        "record_id": None,
    }
    return LLMParseResponse(suggested=sample, confidence=0.0)

# ---------- Helpers ----------


def _row_to_record(r) -> Record:
    # r bir pandas Series ya da dict olabilir
    get = (
        r.get
        if hasattr(r, "get")
        else (lambda k, default=None: r[k] if k in r else default)
    )
    from datetime import date

    def to_float(x):
        try:
            return float(x)
        except Exception:
            return 0.0

    def to_date(x) -> Optional[date]:
        import pandas as pd

        if x in (None, "", "nan", "NaT"):
            return None
        try:
            return pd.to_datetime(x).date()
        except Exception:
            return None

    def value(field: str, default=None):
        column = store.FIELD_TO_COLUMN[field]
        return get(column, default)

    return Record(
        record_id=str(value("record_id", "")) or None,
        date_of_request=to_date(value("date_of_request")),
        salesman=str(value("salesman", "")),
        customer_name=str(value("customer_name", "")),
        customer_po_no=str(value("customer_po_no", "")),
        salesforce_reference=str(value("salesforce_reference", "")),
        so_no=str(value("so_no", "")),
        amount_eur=to_float(value("amount_eur")),
        total_discount_pct=to_float(value("total_discount_pct")),
        cpi_eur=to_float(value("cpi_eur")),
        cps_eur=to_float(value("cps_eur")),
        definition=str(value("definition", "")),
        date_of_delivery=to_date(value("date_of_delivery")),
        date_of_invoice=to_date(value("date_of_invoice")),
        note=str(value("note", "")),
    )
