# -*- coding: utf-8 -*-
import os
from datetime import date

import pandas as pd
import requests
import streamlit as st

FALLBACK_API_BASE = "http://localhost:8000"


def _normalize_base_url(base: str | None, fallback: str = FALLBACK_API_BASE) -> str:
    base = (base or "").strip()
    if not base:
        return fallback
    return base.rstrip("/") or fallback


def _determine_default_api_base() -> str:
    env_base = os.getenv("API_BASE")
    if env_base and env_base.strip():
        return _normalize_base_url(env_base)

    try:
        secret_base = st.secrets.get("API_BASE")  # type: ignore[attr-defined]
        if secret_base and str(secret_base).strip():
            return _normalize_base_url(str(secret_base))
    except Exception:
        pass

    return FALLBACK_API_BASE


DEFAULT_API_BASE = _determine_default_api_base()


if "api_base_override" not in st.session_state:
    st.session_state["api_base_override"] = DEFAULT_API_BASE


def get_api_base() -> str:
    return _normalize_base_url(
        st.session_state.get("api_base_override", DEFAULT_API_BASE),
        DEFAULT_API_BASE,
    )


def build_api_url(path: str, base: str | None = None) -> str:
    fallback = get_api_base()
    base_url = _normalize_base_url(base if base is not None else fallback, fallback)
    if not path.startswith("/"):
        path = f"/{path}"
    return f"{base_url}{path}"

st.set_page_config(page_title="Desoutter Order Track", page_icon="🧭", layout="wide")

# --- Stil ufak dokunuşlar ---
st.markdown("""
<style>
.small { font-size: 0.85rem; color:#666; }
.green-row { background: #16a34a22 !important; }
</style>
""", unsafe_allow_html=True)

st.title("🧭 Desoutter Order Track")
st.caption("Tek Excel dosyası ile kayıt, düzeltme ve zengin raporlama")

# -------------- Yardımcılar --------------
def api_get(path: str, base: str | None = None):
    r = requests.get(build_api_url(path, base), timeout=30)
    r.raise_for_status()
    return r.json()

def api_post(path: str, json: dict, base: str | None = None):
    r = requests.post(build_api_url(path, base), json=json, timeout=30)
    r.raise_for_status()
    return r.json()

def api_put(path: str, json: dict, base: str | None = None):
    r = requests.put(build_api_url(path, base), json=json, timeout=30)
    r.raise_for_status()
    return r.json()


def api_get_bytes(path: str, base: str | None = None) -> bytes:
    r = requests.get(build_api_url(path, base), timeout=30)
    r.raise_for_status()
    return r.content

@st.cache_data(ttl=30)
def load_salesmen(api_base: str):
    data = api_get("/data/salesmen", base=api_base)
    return data["items"]

def refresh_salesmen():
    load_salesmen.clear()

def style_invoice_green(df):
    def row_style(r):
        return ["green-row" if str(r.get("Date of Invoice") or "").strip() not in ("", "NaT") else "" for _ in r]
    return df.style.apply(row_style, axis=1)

def compute_cpi(amount, cps):
    amount = float(amount or 0)
    cps = float(cps or 0)
    return amount - cps if cps else amount

# -------------- Giriş Ekranı --------------
with st.sidebar:
    st.header("⚙️ Ayarlar")
    current_api_base = get_api_base()
    api_url = st.text_input(
        "API Base URL",
        value=current_api_base,
        help="FastAPI sunucusu adresi",
    )
    normalized_url = _normalize_base_url(api_url, DEFAULT_API_BASE)
    if normalized_url != current_api_base:
        st.session_state["api_base_override"] = normalized_url
        refresh_salesmen()
        st.session_state.pop("excel_bytes", None)
        current_api_base = normalized_url

    st.markdown("---")
    st.subheader("📥 Excel Çıktısı")
    excel_bytes = st.session_state.get("excel_bytes")
    if st.button("Excel dosyasını hazırla", key="prepare_excel"):
        try:
            data = api_get_bytes("/records/export", base=current_api_base)
            st.session_state["excel_bytes"] = data
            st.success("Excel dosyası indirilmeye hazır.")
        except Exception as e:
            st.error(f"Excel alınamadı: {e}")
    if excel_bytes:
        st.download_button(
            "Excel dosyasını indir",
            data=excel_bytes,
            file_name="Desoutter Order Track.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel",
        )

    st.markdown("---")
    st.subheader("👥 SalesMan Data")
    # Listele
    sms = load_salesmen(current_api_base)
    st.write(pd.DataFrame(sms))
    # Ekle/Güncelle
    with st.form("salesman_form", clear_on_submit=True):
        name = st.text_input("SalesMan adı", "")
        region = st.selectbox("Bölge", ["CPI Northern", "CPI Southern", "Unassigned"])
        submitted = st.form_submit_button("Kaydet / Güncelle")
        if submitted and name.strip():
            _ = api_post(
                "/data/salesmen",
                {"name": name.strip(), "region": region},
                base=current_api_base,
            )
            st.success("SalesMan kaydedildi.")
            refresh_salesmen()

# -------------- Sayfa Sekmeleri --------------
tab1, tab2 = st.tabs(["📋 Kayıt", "📊 Raporlar"])

api_base = get_api_base()

with tab1:
    st.subheader("Giriş modu")
    mode = st.radio("İşlem seçin", ["Yeni Kayıt", "Mevcut Kaydı Düzelt"], horizontal=True)

    # LLM kutusu (aç/kapa)
    with st.expander("📧 LLM ile e-posta metninden alanları doldur (opsiyonel)"):
        email_text = st.text_area("E-posta metnini yapıştırın", height=150)
        if st.button("Ön Doldur (LLM Stub)"):
            try:
                parsed = api_post("/llm/parse", {"email_text": email_text}, base=api_base)
                st.session_state["prefill"] = parsed["suggested"]
                st.success("Ön dolum önerileri yüklendi.")
            except Exception as e:
                st.error(f"LLM parse hatası: {e}")

    pre = st.session_state.get("prefill", {})

    if mode == "Yeni Kayıt":
        with st.form("create_form"):
            col1, col2, col3 = st.columns(3)

            with col1:
                date_of_request = st.date_input("Date of Request", value=pre.get("date_of_request") or date.today())
                customer_name = st.text_input("Customer Name", value=pre.get("customer_name", ""))
                salesforce_reference = st.text_input("SalesForce Reference", value=pre.get("salesforce_reference", ""))

            with col2:
                salesmen = [s["name"] for s in load_salesmen(api_base)]
                salesman = st.selectbox("SalesMan", options=salesmen, index=0 if salesmen else None)
                customer_po_no = st.text_input("Customer PO No", value=pre.get("customer_po_no", ""))
                so_no = st.text_input("SO No", value=pre.get("so_no", ""))

            with col3:
                amount = st.number_input("Amount (€)", min_value=0.0, step=100.0, value=float(pre.get("amount_eur", 0.0) or 0.0))
                cps = st.number_input("CPS (€)", min_value=0.0, step=10.0, value=float(pre.get("cps_eur", 0.0) or 0.0))
                total_discount = st.number_input("Total Discount (%)", min_value=0.0, max_value=100.0, step=0.5, value=float(pre.get("total_discount_pct", 0.0) or 0.0))

            cpi = compute_cpi(amount, cps)
            st.metric("CPI (€)", f"{cpi:,.2f}")

            definition = st.text_input("Definition", value=pre.get("definition", ""))
            col4, col5, col6 = st.columns(3)
            with col4:
                date_of_delivery = st.date_input("Date of Delivery", value=pre.get("date_of_delivery"))
            with col5:
                date_of_invoice = st.date_input("Date of Invoice", value=pre.get("date_of_invoice"))
            with col6:
                note = st.text_input("Note", value=pre.get("note", ""))

            submitted = st.form_submit_button("➕ Kaydı Ekle")
            if submitted:
                payload = {
                    "date_of_request": str(date_of_request),
                    "salesman": salesman or "",
                    "customer_name": customer_name,
                    "customer_po_no": customer_po_no,
                    "salesforce_reference": salesforce_reference,
                    "so_no": so_no,
                    "amount_eur": float(amount),
                    "total_discount_pct": float(total_discount),
                    "cpi_eur": float(cpi),
                    "cps_eur": float(cps),
                    "definition": definition,
                    "date_of_delivery": str(date_of_delivery) if date_of_delivery else None,
                    "date_of_invoice": str(date_of_invoice) if date_of_invoice else None,
                    "note": note,
                }
                try:
                    rec = api_post("/records", payload, base=api_base)
                    st.success(f"Kayıt eklendi. Record ID: {rec['record_id']}")
                    st.session_state.pop("prefill", None)
                except Exception as e:
                    st.error(f"Hata: {e}")

    else:
        st.info("SO No veya Customer PO No girerek kaydı bulun. Sonra form üzerinde güncelleyin.")
        colL, colR = st.columns([1,2])
        with colL:
            lookup_type = st.selectbox("Arama türü", ["SO No", "Customer PO No"])
            lookup_value = st.text_input("Arama değeri")
            if st.button("Bul"):
                try:
                    q = {"so_no": lookup_value} if lookup_type == "SO No" else {"customer_po_no": lookup_value}
                    rec = api_post("/records/lookup", q, base=api_base)
                    st.session_state["editing"] = rec
                    st.success("Kayıt yüklendi.")
                except Exception as e:
                    st.error(f"Bulunamadı: {e}")

        rec = st.session_state.get("editing")
        if rec:
            with st.form("edit_form"):
                rid = rec["record_id"]
                col1, col2, col3 = st.columns(3)

                with col1:
                    date_of_request = st.date_input("Date of Request", value=pd.to_datetime(rec["date_of_request"]).date())
                    customer_name = st.text_input("Customer Name", value=rec["customer_name"])
                    salesforce_reference = st.text_input("SalesForce Reference", value=rec["salesforce_reference"])

                with col2:
                    salesmen = [s["name"] for s in load_salesmen(api_base)]
                    salesman = st.selectbox("SalesMan", options=salesmen, index=(salesmen.index(rec["salesman"]) if rec["salesman"] in salesmen else 0))
                    customer_po_no = st.text_input("Customer PO No", value=rec["customer_po_no"])
                    so_no = st.text_input("SO No", value=rec["so_no"])

                with col3:
                    amount = st.number_input("Amount (€)", min_value=0.0, step=100.0, value=float(rec["amount_eur"]))
                    cps = st.number_input("CPS (€)", min_value=0.0, step=10.0, value=float(rec["cps_eur"]))
                    total_discount = st.number_input("Total Discount (%)", min_value=0.0, max_value=100.0, step=0.5, value=float(rec["total_discount_pct"]))

                cpi = compute_cpi(amount, cps)
                st.metric("CPI (€)", f"{cpi:,.2f}")

                definition = st.text_input("Definition", value=rec.get("definition") or "")
                col4, col5, col6 = st.columns(3)
                with col4:
                    date_of_delivery = st.date_input("Date of Delivery", value=pd.to_datetime(rec["date_of_delivery"]).date() if rec.get("date_of_delivery") else None)
                with col5:
                    date_of_invoice = st.date_input("Date of Invoice", value=pd.to_datetime(rec["date_of_invoice"]).date() if rec.get("date_of_invoice") else None)
                with col6:
                    note = st.text_input("Note", value=rec.get("note") or "")

                submitted = st.form_submit_button("💾 Kaydı Güncelle")
                if submitted:
                    payload = {
                        "record_id": rid,
                        "date_of_request": str(date_of_request),
                        "salesman": salesman or "",
                        "customer_name": customer_name,
                        "customer_po_no": customer_po_no,
                        "salesforce_reference": salesforce_reference,
                        "so_no": so_no,
                        "amount_eur": float(amount),
                        "total_discount_pct": float(total_discount),
                        "cpi_eur": float(cpi),
                        "cps_eur": float(cps),
                        "definition": definition,
                        "date_of_delivery": str(date_of_delivery) if date_of_delivery else None,
                        "date_of_invoice": str(date_of_invoice) if date_of_invoice else None,
                        "note": note,
                    }
                    try:
                        updated = api_put(f"/records/{rid}", payload, base=api_base)
                        st.success("Kayıt güncellendi.")
                        st.session_state["editing"] = updated
                    except Exception as e:
                        st.error(f"Hata: {e}")

    # Kayıtlar önizleme
    st.markdown("---")
    st.subheader("Son Kayıtlar")
    try:
        items = api_get("/records", base=api_base)["items"]
        dfv = pd.DataFrame(items)
        if not dfv.empty:
            # Tarih & görsel vurgu
            dfv = dfv.rename(
                columns={
                    "date_of_request": "Date of Request",
                    "salesman": "Sales Person",
                    "customer_name": "Customer Name",
                    "customer_po_no": "Customer PO No",
                    "salesforce_reference": "Salesforce Reference",
                    "so_no": "SO No",
                    "amount_eur": "Amount (EUR)",
                    "total_discount_pct": "Total Discount (%)",
                    "cpi_eur": "CPI (EUR)",
                    "cps_eur": "CPS (EUR)",
                    "definition": "Definition",
                    "date_of_delivery": "Date of Delivery",
                    "date_of_invoice": "Date of Invoice",
                    "note": "Note",
                    "record_id": "Record ID",
                }
            )
            ordered = [
                "Date of Request",
                "Sales Person",
                "Customer Name",
                "Customer PO No",
                "Salesforce Reference",
                "SO No",
                "Amount (EUR)",
                "Total Discount (%)",
                "CPI (EUR)",
                "CPS (EUR)",
                "Definition",
                "Date of Delivery",
                "Date of Invoice",
                "Note",
                "Record ID",
            ]
            available = [col for col in ordered if col in dfv.columns]
            dfv = dfv[available + [c for c in dfv.columns if c not in available]]
            st.dataframe(dfv.style.apply(
                lambda r: ["background-color: #d1fae5" if str(r["Date of Invoice"]).strip() not in ("", "NaT", "None") else "" for _ in r],
                axis=1
            ), use_container_width=True, hide_index=True)
        else:
            st.info("Henüz kayıt yok.")
    except Exception as e:
        st.error(f"Önizleme hatası: {e}")

with tab2:
    st.subheader("Raporlar")
    try:
        rep = api_get("/reports/summary", base=api_base)
        # Bölge bazında
        st.markdown("### Bölge Bazında Toplamlar")
        df_region = pd.DataFrame(rep["by_region"])
        if not df_region.empty:
            c1, c2 = st.columns([2,1])
            with c1:
                st.bar_chart(df_region.set_index("Region")[["Amount (EUR)", "CPI (EUR)", "CPS (EUR)"]])
            with c2:
                st.dataframe(df_region, use_container_width=True, hide_index=True)
        else:
            st.info("Bölge raporu için veri yok.")

        st.markdown("### CPI vs CPS")
        df_c = pd.DataFrame(rep["cpi_vs_cps"])
        if not df_c.empty:
            st.bar_chart(df_c.set_index("Metric")["EUR"])
            st.dataframe(df_c, use_container_width=True, hide_index=True)

        st.markdown("### OR (Order Received) – Yıllara Göre")
        df_or = pd.DataFrame(rep["or_by_year"])
        if not df_or.empty:
            st.line_chart(df_or.set_index("Year")["OR (EUR)"])
            st.dataframe(df_or, use_container_width=True, hide_index=True)
        else:
            st.info("OR için veri yok.")

        st.markdown("### OI (Order Invoiced) – Yıllara Göre")
        df_oi = pd.DataFrame(rep["oi_by_year"])
        if not df_oi.empty:
            st.line_chart(df_oi.set_index("Year")["OI (EUR)"])
            st.dataframe(df_oi, use_container_width=True, hide_index=True)
        else:
            st.info("OI için veri yok.")
    except Exception as e:
        st.error(f"Raporlar yüklenemedi: {e}")
