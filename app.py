import streamlit as st
import os, io, datetime, zipfile, time
import fitz  # PyMuPDF
from docx import Document
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import PatternFill
from PIL import Image

# --- YENI AÇIK KAYNAK YAPAY ZEKA MİMARİSİ ---
import instructor
from groq import Groq
from pydantic import BaseModel, Field
from typing import List, Optional, Dict

APP_VERSION = "v2.0 (Groq Llama-3 & Instructor Core)"
TEMPLATE_PATH = "SEAFARER APPLICATION FORM.xlsx"
RED_FILL = PatternFill(start_color="FF4D4D", end_color="FF4D4D", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

st.set_page_config(page_title="Bluefin Crewing - Enterprise CV Parser", layout="wide")

# --- PYDANTIC VERİ ŞEMALARI (SIFIR HATA GARANTİSİ) ---
class SeaService(BaseModel):
    vsl_type: str = Field(default="")
    vsl_name: str = Field(default="")
    company: str = Field(default="")
    rank: str = Field(default="")
    dwt: str = Field(default="")
    engine_type: str = Field(default="")
    sign_on: str = Field(default="", description="DD.MM.YYYY formatinda")
    sign_off: str = Field(default="", description="DD.MM.YYYY formatinda")
    reason: str = Field(default="")

class STCWDate(BaseModel):
    issue: str = Field(default="")
    expiry: str = Field(default="")

class SeafarerCV(BaseModel):
    name: str = Field(default="")
    surname: str = Field(default="")
    rank_applied: str = Field(default="")
    qualification: str = Field(default="")
    id_no: str = Field(default="")
    dob: str = Field(default="")
    english_level: str = Field(default="")
    phone: str = Field(default="")
    phone2: str = Field(default="")
    email: str = Field(default="")
    address: str = Field(default="")
    marital_status: str = Field(default="")
    children: str = Field(default="")
    height: str = Field(default="")
    weight: str = Field(default="")
    shoes: str = Field(default="")
    overall: str = Field(default="")
    next_of_kin: str = Field(default="Isim - Telefon - Adres")
    seaman_book_no: str = Field(default="")
    seaman_book_issue: str = Field(default="")
    seaman_book_expiry: str = Field(default="")
    passport_no: str = Field(default="")
    passport_issue: str = Field(default="")
    passport_expiry: str = Field(default="")
    medical_issue: str = Field(default="")
    medical_expiry: str = Field(default="")
    yellow_fever_issue: str = Field(default="")
    yellow_fever_expiry: str = Field(default="")
    usa_visa: str = Field(default="")
    schengen_visa: str = Field(default="")
    sea_service: List[SeaService] = Field(default_factory=list, description="En guncel gemi en ustte olacak sekilde sirala.")
    stcw_certificates: Dict[str, STCWDate] = Field(default_factory=dict, description="Sertifika kodlarini (Orn: '52', '53') key olarak kullan")
    uncertain_fields: List[str] = Field(default_factory=list, description="Okunamayan veya emin olunmayan alanlarin listesi")
    warnings: List[str] = Field(default_factory=list, description="Zabitler icin 4, ratingler icin 6 aydan kisa hizmet suresi uyari mesajlari")

@st.cache_data(show_spinner=False)
def load_excel_template():
    with open(TEMPLATE_PATH, "rb") as f:
        return f.read()

def extract_content(file_bytes, file_name):
    ext = os.path.splitext(file_name)[1].lower()
    text = ""
    photo = None
    
    if ext == ".docx":
        doc = Document(io.BytesIO(file_bytes))
        for p in doc.paragraphs:
            text += p.text + "\n"
        for t in doc.tables:
            for r in t.rows:
                text += " | ".join([c.text.strip() for c in r.cells if c.text.strip()]) + "\n"
    elif ext == ".pdf":
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page in doc:
            text += page.get_text()
            # Hizli fotograf cikarimi
            if not photo:
                images = page.get_images(full=True)
                if images:
                    xref = images[0][0]
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    try:
                        p_candidate = Image.open(io.BytesIO(image_bytes)).convert("RGB")
                        if p_candidate.size[0] > 80 and p_candidate.size[1] > 80:
                            photo = p_candidate
                    except:
                        pass
    return text, photo

def analyze_cv_with_groq(api_key, text_content):
    # Google yerine Groq + Instructor mimarisi
    client = instructor.from_groq(Groq(api_key=api_key.strip()), mode=instructor.Mode.JSON)
    
    prompt = f"""
    Asagidaki seafarer (gemi adami) CV'sini analiz et ve BİREBİR Pydantic semasina uygun doldur.
    Tarihleri DD.MM.YYYY yap. Eksiklere "" koy.
    Zabitler icin 4 ay, Ratingler icin 6 aydan kisa hizmetleri 'warnings' dizisine "[Gemi Adi] - Erken Ayrilma" olarak ekle.
    Emin olmadigin alanlari 'uncertain_fields' icine yaz.
    
    CV METNİ:
    {text_content}
    """
    
    try:
        # LLM'e veriyi Pydantic nesnesi olarak uretmesini zorunlu kiliyoruz
        cv_data = client.chat.completions.create(
            model="llama3-70b-8192",
            response_model=SeafarerCV,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0
        )
        return cv_data.model_dump() # Excel fonksiyonu icin Dictionary'ye cevir
    except Exception as e:
        raise Exception(f"Sistem Hatasi (Groq/Parser): {str(e)}")

def safe_write(ws, row, column, value, fill=None):
    cell = ws.cell(row=row, column=column)
    if type(cell).__name__ == "MergedCell":
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                cell = ws.cell(row=rng.min_row, column=rng.min_col)
                break
    cell.value = value
    if fill:
        cell.fill = fill
    return cell

def fill_excel_template(data, photo=None):
    template_bytes = load_excel_template()
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb.active
    today = datetime.datetime.now()
    
    uncertain_keys = data.get('uncertain_fields', [])
    def get_fill(field_name): return YELLOW_FILL if field_name in uncertain_keys else None

    safe_write(ws, 7, 3, data.get('name', ''), fill=get_fill('name'))
    safe_write(ws, 8, 3, data.get('surname', ''), fill=get_fill('surname'))
    safe_write(ws, 9, 3, data.get('rank_applied', ''), fill=get_fill('rank_applied'))
    safe_write(ws, 9, 8, data.get('qualification', ''), fill=get_fill('qualification'))
    safe_write(ws, 7, 8, data.get('id_no', ''), fill=get_fill('id_no'))
    safe_write(ws, 8, 8, data.get('dob', ''), fill=get_fill('dob'))
    safe_write(ws, 11, 3, data.get('english_level', ''), fill=get_fill('english_level'))
    safe_write(ws, 12, 3, data.get('phone', ''), fill=get_fill('phone'))
    safe_write(ws, 12, 8, data.get('phone2', ''), fill=get_fill('phone2'))
    safe_write(ws, 13, 3, data.get('email', ''), fill=get_fill('email'))
    safe_write(ws, 14, 3, data.get('address', ''), fill=get_fill('address'))
    safe_write(ws, 15, 3, data.get('marital_status', ''), fill=get_fill('marital_status'))
    safe_write(ws, 15, 8, data.get('children', ''), fill=get_fill('children'))
    safe_write(ws, 16, 3, data.get('height', ''), fill=get_fill('height'))
    safe_write(ws, 16, 8, data.get('weight', ''), fill=get_fill('weight'))
    safe_write(ws, 17, 3, data.get('shoes', ''), fill=get_fill('shoes'))
    safe_write(ws, 17, 5, data.get('overall', ''), fill=get_fill('overall'))
    safe_write(ws, 18, 4, data.get('next_of_kin', ''), fill=get_fill('next_of_kin'))

    safe_write(ws, 20, 5, data.get('seaman_book_no', ''), fill=get_fill('seaman_book_no'))
    safe_write(ws, 20, 6, data.get('seaman_book_issue', ''), fill=get_fill('seaman_book_issue'))
    safe_write(ws, 20, 8, data.get('seaman_book_expiry', ''), fill=get_fill('seaman_book_expiry'))
    safe_write(ws, 21, 5, data.get('passport_no', ''), fill=get_fill('passport_no'))
    safe_write(ws, 21, 6, data.get('passport_issue', ''), fill=get_fill('passport_issue'))
    safe_write(ws, 21, 8, data.get('passport_expiry', ''), fill=get_fill('passport_expiry'))
    safe_write(ws, 22, 6, data.get('medical_issue', ''), fill=get_fill('medical_issue'))
    safe_write(ws, 22, 8, data.get('medical_expiry', ''), fill=get_fill('medical_expiry'))

    if photo:
        photo_buf = io.BytesIO()
        photo_resized = photo.copy()
        photo_resized.thumbnail((115, 140))
        photo_resized.save(photo_buf, format="PNG")
        photo_buf.seek(0)
        img = OpenpyxlImage(photo_buf)
        ws.add_image(img, "H1")

    for row_str, dates in data.get('stcw_certificates', {}).items():
        try:
            r = int(row_str)
            if dates.get('issue'): safe_write(ws, r, 8, dates['issue'], fill=get_fill('stcw_certificates'))
            if dates.get('expiry'): safe_write(ws, r, 9, dates['expiry'], fill=get_fill('stcw_certificates'))
        except:
            pass

    r_idx = 30
    for exp in data.get('sea_service', []):
        row_fill = get_fill('sea_service')
        safe_write(ws, r_idx, 1, exp.get('vsl_type', ''), fill=row_fill)
        safe_write(ws, r_idx, 2, exp.get('vsl_name', ''), fill=row_fill)
        safe_write(ws, r_idx, 3, exp.get('company', ''), fill=row_fill)
        safe_write(ws, r_idx, 4, exp.get('rank', ''), fill=row_fill)
        safe_write(ws, r_idx, 5, exp.get('dwt', ''), fill=row_fill)
        safe_write(ws, r_idx, 6, exp.get('engine_type', ''), fill=row_fill)
        safe_write(ws, r_idx, 7, exp.get('sign_on', ''), fill=row_fill)
        safe_write(ws, r_idx, 8, exp.get('sign_off', ''), fill=row_fill)
        safe_write(ws, r_idx, 9, exp.get('reason', ''), fill=row_fill)
        r_idx += 1

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    return out_stream

with st.sidebar:
    st.markdown(f"**⚓ Mimari:** `{APP_VERSION}`")
    st.divider()
    
    # Artik Google degil, Groq API Anahtari kullanacagiz
    saved_key = st.secrets.get("GROQ_API_KEY", "")
    api_key = st.text_input("Groq API Anahtari:", type="password", value=saved_key)
    st.caption("Google Gemini devreden cikarildi. Lutfen console.groq.com adresinden aldiginiz anahtari girin.")
    
if not api_key:
    st.error("Sistemi kullanmak icin Groq API anahtari gereklidir.")
    st.stop()

uploaded_file = st.file_uploader("Donusturulecek CV yi Secin", type=["pdf", "docx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    
    if st.button("🚀 CV'yi Çözümle (Sıfır Hata Modu)", type="primary"):
        with st.spinner("Yeni Llama-3 Motoru Devrede..."):
            try:
                text_content, photo = extract_content(file_bytes, uploaded_file.name)
                parsed_data = analyze_cv_with_groq(api_key, text_content)
                st.session_state['parsed_data'] = parsed_data
                st.session_state['photo'] = photo
                st.success("Çözümleme saniyeler içinde başarıyla tamamlandı!")
            except Exception as e:
                st.error(str(e))

    if 'parsed_data' in st.session_state:
        d = st.session_state['parsed_data']
        photo = st.session_state.get('photo')

        excel_data = fill_excel_template(d, photo)
        safe_name = f"{d.get('rank_applied','')}_{d.get('name','')}_{d.get('surname','')}".replace(' ', '_').upper()
        
        st.download_button(
            label=f"📥 İNDİR: {safe_name}.xlsx",
            data=excel_data,
            file_name=f"{safe_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        st.divider()

        if d.get('warnings'):
            for warning in d['warnings']: st.error(f"🚨 DİKKAT: {warning}")
        if d.get('uncertain_fields'):
            st.warning(f"⚠️ SARI İŞARETLİ ALANLAR: {', '.join(d['uncertain_fields'])}")
