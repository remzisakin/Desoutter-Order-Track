# Desoutter Order Track

Tek bir Excel dosyasına (“**data/Desoutter Order Track.xlsx**”) sürekli veri ekleyen, mevcut kayıtları düzenlemeye izin veren; SalesMan & Bölge (Data) yönetimi ve otomatik raporlama sekmeleri olan **Streamlit (frontend)** + **FastAPI (backend)** uygulaması.

## Kurulum

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Çalıştırma

1) Backend (FastAPI)
```bash
uvicorn backend.main:app --reload --port 8000
```

2) Frontend (Streamlit)
```bash
streamlit run frontend/app.py
```

Varsayılan olarak Streamlit, http://localhost:8501 üzerinde çalışır ve http://localhost:8000’daki API’ye bağlanır.

İlk çalıştırmada data/Desoutter Order Track.xlsx dosya ve sayfaları otomatik oluşturulur:

- **Records**: tüm kayıtlar
- **Data**: SalesMan & Region eşleştirmeleri

## Özellikler

- **Giriş modu**: Uygulama açılışında “Yeni Kayıt” veya “Mevcut Kaydı Düzelt”.
- **Zorunlu alanlar**: Date of Request (takvim), SalesMan (select), Customer Name, Customer PO No, SalesForce Reference, SO No, Amount (€), Total Discount (%), CPI (€), CPS (€).
- **CPI kuralı**: CPS > 0 ise CPI = Amount - CPS, değilse CPI = Amount.
- **Opsiyonel alanlar**: Defination, Date of Delivery, Date of Invoice, Note.
- **Görsel vurgu**: Date of Invoice doluysa liste görünümünde satır yeşil renkte gösterilir.
- **SalesMan & Bölge Yönetimi (Data)**: Sol kenardaki panelden SalesMan ekle/güncelle; bölge olarak CPI Northern / CPI Southern atanabilir.
- **Kayıt düzeltme**: SO No veya Customer PO No ile arayıp ilgili satırı bul, formu düzenle, kaydet.
- **LLM kutusu (opsiyonel)**: E-posta metnini yapıştır → ileride devreye alınacak parsere gönderir (şimdilik stub).
- **Raporlar**: Bölge bazında toplamlar, CPI vs CPS, OR (Order Received) yıllara göre, OI (Order Invoiced) yıllara göre. Veri arttıkça otomatik güncellenir.

## LLM Entegrasyonu (Opsiyonel)

`backend/main.py` içinde `/llm/parse` endpointi stub’dır. OpenAI vb. ile bağlamak isterseniz:

1. `requirements.txt` içine `openai` ekleyin.
2. Ortama `OPENAI_API_KEY` koyun.
3. `/llm/parse` içinde `email_text`’i prompt’a verip `Record` şemasına uygun alanları çıkarın.

## GitHub’da Çalıştırma

Bu klasörü GitHub’a push edin.

Sunucuda/PC’de:

```bash
git clone <repo-url>
cd desoutter-order-track
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn backend.main:app --host 0.0.0.0 --port 8000
# başka bir terminal
streamlit run frontend/app.py
```

Streamlit tarafında farklı API adresi gerekiyorsa `frontend/app.py` içinde `API_BASE`’i değiştirebilir ya da `~/.streamlit/secrets.toml` dosyasına:

```
API_BASE = "http://sunucu-adresiniz:8000"
```

yazarak yapılandırabilirsiniz.

## Notlar

- Excel dosyası başka bir programda açıkken yazma hatası alabilirsiniz; kapatıp tekrar deneyin.
- Records sayfasında `record_id` alanı backend tarafından üretilen benzersiz kimliktir; güncellemelerde kullanılır.
- “OR” toplamları Date of Request’e göre, “OI” toplamları Date of Invoice’ı dolu kayıtlara göre hesaplanır.

---

### Hepsi bu kadar 🎯

İsterseniz **SalesMan isimlerini ve bölgelerini** bana şimdi liste olarak verin; backend’e uygun **toplu yükleme JSON**’unu da hazırlayıp paylaşayım. Ayrıca LLM tarafını da (OpenAI ile) bağlamak isterseniz, `/llm/parse` için örnek bir prompt & kod parçası da ekleyebilirim.
