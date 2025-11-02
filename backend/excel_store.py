import os
import uuid
from datetime import datetime
from typing import List, Optional, Dict, Any

import pandas as pd

EXCEL_PATH = os.path.join("data", "Desoutter Order Track.xlsx")
RECORDS_SHEET = "Records"
DATA_SHEET = "Data"  # Salesman & Region yönetimi

RECORD_COLUMNS = [
    "record_id",
    "Date of Request",
    "SalesMan",
    "Region",
    "Customer Name",
    "Customer PO No",
    "SalesForce Reference",
    "SO No",
    "Amount (EUR)",
    "Total Discount (%)",
    "CPI (EUR)",
    "CPS (EUR)",
    "Defination",
    "Date of Delivery",
    "Date of Invoice",
    "Note",
    "created_at",
    "updated_at",
]

DATA_COLUMNS = ["SalesMan", "Region"]  # Region: CPI Northern / CPI Southern / Unassigned

def _ensure_file_structure():
    os.makedirs(os.path.dirname(EXCEL_PATH), exist_ok=True)
    if not os.path.exists(EXCEL_PATH):
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as w:
            pd.DataFrame(columns=RECORD_COLUMNS).to_excel(w, sheet_name=RECORDS_SHEET, index=False)
            pd.DataFrame(columns=DATA_COLUMNS).to_excel(w, sheet_name=DATA_SHEET, index=False)
        return

    # Sayfalar yoksa ekle
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="overlay") as w:
        wb = w.book
        if RECORDS_SHEET not in wb.sheetnames:
            pd.DataFrame(columns=RECORD_COLUMNS).to_excel(w, sheet_name=RECORDS_SHEET, index=False)
        if DATA_SHEET not in wb.sheetnames:
            pd.DataFrame(columns=DATA_COLUMNS).to_excel(w, sheet_name=DATA_SHEET, index=False)

def _read_df(sheet: str) -> pd.DataFrame:
    _ensure_file_structure()
    try:
        return pd.read_excel(EXCEL_PATH, sheet_name=sheet, dtype=str)
    except Exception:
        return pd.DataFrame()

def _write_df(sheet: str, df: pd.DataFrame):
    _ensure_file_structure()
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        df.to_excel(w, sheet_name=sheet, index=False)

# ------------------ Salesman Data ------------------

def list_salesmen() -> pd.DataFrame:
    df = _read_df(DATA_SHEET)
    if df.empty:
        df = pd.DataFrame(columns=DATA_COLUMNS)
    # Normalize boşluklar
    df["SalesMan"] = df["SalesMan"].fillna("").astype(str)
    df["Region"] = df["Region"].fillna("Unassigned").astype(str)
    return df

def upsert_salesman(name: str, region: str):
    df = list_salesmen()
    mask = df["SalesMan"].str.lower() == name.strip().lower()
    if mask.any():
        df.loc[mask, "Region"] = region
    else:
        df = pd.concat([df, pd.DataFrame([{"SalesMan": name.strip(), "Region": region}])], ignore_index=True)
    _write_df(DATA_SHEET, df)

def bulk_set_salesmen(items: List[Dict[str, str]]):
    rows = []
    for it in items:
        rows.append({"SalesMan": it["name"], "Region": it.get("region", "Unassigned")})
    df = pd.DataFrame(rows, columns=DATA_COLUMNS)
    _write_df(DATA_SHEET, df)

# ------------------ Records ------------------

def list_records() -> pd.DataFrame:
    df = _read_df(RECORDS_SHEET)
    if df.empty:
        df = pd.DataFrame(columns=RECORD_COLUMNS)
    return df

def _infer_region_for_salesman(salesman: str) -> str:
    s = list_salesmen()
    m = s[s["SalesMan"].str.lower() == str(salesman).strip().lower()]
    if not m.empty:
        return m.iloc[0]["Region"] or "Unassigned"
    return "Unassigned"

def create_record(payload: Dict[str, Any]) -> Dict[str, Any]:
    df = list_records()

    record_id = str(uuid.uuid4())
    now = datetime.now().isoformat(timespec="seconds")

    # CPI/CPS kuralı
    amount = float(payload["amount_eur"])
    cps = float(payload.get("cps_eur", 0.0) or 0.0)
    cpi = amount - cps if cps else amount

    region = _infer_region_for_salesman(payload["salesman"])

    row = {
        "record_id": record_id,
        "Date of Request": str(payload["date_of_request"]),
        "SalesMan": payload["salesman"],
        "Region": region,
        "Customer Name": payload["customer_name"],
        "Customer PO No": payload["customer_po_no"],
        "SalesForce Reference": payload["salesforce_reference"],
        "SO No": payload["so_no"],
        "Amount (EUR)": f"{amount:.2f}",
        "Total Discount (%)": f"{float(payload['total_discount_pct']):.2f}",
        "CPI (EUR)": f"{cpi:.2f}",
        "CPS (EUR)": f"{cps:.2f}",
        "Defination": payload.get("definition", ""),
        "Date of Delivery": str(payload.get("date_of_delivery") or ""),
        "Date of Invoice": str(payload.get("date_of_invoice") or ""),
        "Note": payload.get("note", ""),
        "created_at": now,
        "updated_at": now,
    }

    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    _write_df(RECORDS_SHEET, df)
    return row

def find_record(so_no: Optional[str] = None, customer_po_no: Optional[str] = None) -> Optional[Dict[str, Any]]:
    df = list_records()
    if df.empty:
        return None
    result = pd.DataFrame()
    if so_no:
        result = df[df["SO No"].astype(str).str.lower() == so_no.strip().lower()]
    elif customer_po_no:
        result = df[df["Customer PO No"].astype(str).str.lower() == customer_po_no.strip().lower()]
    if result.empty:
        return None
    return result.iloc[0].to_dict()

def update_record(record_id: str, payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    df = list_records()
    if df.empty:
        return None
    mask = df["record_id"] == record_id
    if not mask.any():
        return None

    # CPI/CPS kuralını tekrar uygula
    amount = float(payload["amount_eur"])
    cps = float(payload.get("cps_eur", 0.0) or 0.0)
    cpi = amount - cps if cps else amount
    region = _infer_region_for_salesman(payload["salesman"])

    df.loc[mask, "Date of Request"] = str(payload["date_of_request"])
    df.loc[mask, "SalesMan"] = payload["salesman"]
    df.loc[mask, "Region"] = region
    df.loc[mask, "Customer Name"] = payload["customer_name"]
    df.loc[mask, "Customer PO No"] = payload["customer_po_no"]
    df.loc[mask, "SalesForce Reference"] = payload["salesforce_reference"]
    df.loc[mask, "SO No"] = payload["so_no"]
    df.loc[mask, "Amount (EUR)"] = f"{amount:.2f}"
    df.loc[mask, "Total Discount (%)"] = f"{float(payload['total_discount_pct']):.2f}"
    df.loc[mask, "CPI (EUR)"] = f"{cpi:.2f}"
    df.loc[mask, "CPS (EUR)"] = f"{cps:.2f}"
    df.loc[mask, "Defination"] = payload.get("definition", "")
    df.loc[mask, "Date of Delivery"] = str(payload.get("date_of_delivery") or "")
    df.loc[mask, "Date of Invoice"] = str(payload.get("date_of_invoice") or "")
    df.loc[mask, "Note"] = payload.get("note", "")
    df.loc[mask, "updated_at"] = datetime.now().isoformat(timespec="seconds")

    _write_df(RECORDS_SHEET, df)
    return df.loc[mask].iloc[0].to_dict()

# ------------------ Reports ------------------

def report_frames() -> Dict[str, pd.DataFrame]:
    df = list_records()
    if df.empty:
        return {
            "by_region": pd.DataFrame(columns=["Region", "Amount (EUR)", "CPI (EUR)", "CPS (EUR)"]),
            "or_by_year": pd.DataFrame(columns=["Year", "OR (EUR)"]),
            "oi_by_year": pd.DataFrame(columns=["Year", "OI (EUR)"]),
            "cpi_vs_cps": pd.DataFrame(columns=["Metric", "EUR"]),
        }

    # Numerik çeviriler
    for col in ["Amount (EUR)", "CPI (EUR)", "CPS (EUR)"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Yıl alanları
    df["Year_OR"] = pd.to_datetime(df["Date of Request"], errors="coerce").dt.year
    df["Year_OI"] = pd.to_datetime(df["Date of Invoice"], errors="coerce").dt.year

    by_region = df.groupby("Region")[['Amount (EUR)', 'CPI (EUR)', 'CPS (EUR)']].sum().reset_index()

    or_by_year = (
        df.groupby("Year_OR")[['Amount (EUR)']].sum().reset_index()
        .rename(columns={"Amount (EUR)": "OR (EUR)", "Year_OR": "Year"})
        .dropna(subset=["Year"])
    )

    filtered = df[~df["Date of Invoice"].isna() & (df["Date of Invoice"].astype(str) != "")].copy()
    oi_by_year = (
        filtered.groupby("Year_OI")[['CPI (EUR)', 'CPS (EUR)']].sum().sum(axis=1)
        .reset_index(name="OI (EUR)")
        .rename(columns={"Year_OI": "Year"})
    )

    cpi_vs_cps = pd.DataFrame([
        {"Metric": "CPI (EUR)", "EUR": df["CPI (EUR)"].sum()},
        {"Metric": "CPS (EUR)", "EUR": df["CPS (EUR)"].sum()},
    ])
    return {
        "by_region": by_region,
        "or_by_year": or_by_year,
        "oi_by_year": oi_by_year,
        "cpi_vs_cps": cpi_vs_cps,
    }
