import streamlit as st
from groq import Groq
import json
import re
import pandas as pd
from openpyxl import load_workbook
import fitz  # PyMuPDF
from docx import Document
from pydantic import BaseModel, Field, ValidationError
from typing import List, Dict, Optional

# ==========================================
# 1. MODELLER (VERİ YAPISI)
# ==========================================
class SeaServiceEntry(BaseModel):
    vsl_type: str = ""
    vsl_name: str = ""
    company: str = ""
    rank: str = ""
    dwt: str = ""
    engine_type: str = ""
    sign_on: str = ""
    sign_off: str = ""
    reason: str = ""

class STCWCertificate(BaseModel):
    issue: str = ""
    expiry: str = ""

class CVData(BaseModel):
    name: str = ""
    surname: str = ""
    rank_applied: str = ""
    qualification: str = ""
    id_no: str = ""
    dob: str = ""
    english_level: str = ""
    phone: str = ""
    phone2: str = ""
    email: str = ""
    address: str = ""
    marital_status: str = ""
    children: str = ""
    height: str = ""
    weight: str = ""
    shoes: str = ""
    overall: str = ""
    next_of_kin: str = ""
    seaman_book_no: str = ""
    seaman_book_issue: str = ""
    seaman_book_expiry: str = ""
    passport_no: str = ""
    passport_issue: str = ""
    passport_expiry: str = ""
    medical_issue: str = ""
    medical_expiry: str = ""
    yellow_fever_issue: str = ""
    yellow_fever_expiry: str = ""
    usa_visa: str = ""
    schengen_visa: str = ""
    sea_service: List[SeaServiceEntry] = []
    stcw_certificates: Dict[str, STCWCertificate] = {}
    uncertain_fields: List[str] = []
    warnings: List[str] = []

    def normalize_dates(self):
        date_fields = [
            'dob', 'seaman_book_issue', 'seaman_book_expiry', 
            'passport_issue', 'passport_expiry', 'medical_issue', 
            'medical_expiry', 'yellow_fever_issue', 'yellow_fever_expiry'
        ]
        for field in date_fields:
            val = getattr(self, field)
            if val: setattr(self, field, self._clean_date(val))
        for entry in self.sea_service:
            entry.sign_on = self._clean_date(entry.sign_on)
            entry.sign_off = self._clean_date(entry.sign_off)

    def _clean_date(self, date_str):
        if not date_str: return ""
        match = re.search(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})', date_str)
        if match:
            d, m, y = match.groups()
            return f"{int(d):02d}.{int(m):02d}.{y}"
        return date_str

# ==========================================
# 2. BELGE OKUMA ARAÇLARI
# ==========================================
def extract_text_from_pdf(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    return text

def extract_text_from_docx(file):
    doc = Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

def safe_extract(uploaded_file):
    extension = uploaded_file.name.split('.')[-1].lower()
    if extension == 'pdf':
        return extract_text_from_pdf(uploaded_file.read())
    elif extension == 'docx':
        return extract_text_from_docx(uploaded_file)
    else:
        raise Exception("Desteklenmeyen dosya formatı. Lütfen PDF veya DOCX yükleyin.")

# ==========================================
# 3. ANA UYGULAMA VE LOGIC
# ==========================================
st.set_page_config(page_title="BlueFin Crewing Parser", page_icon="🚢", layout="wide")

# --- API KEY BURAYA YERLEŞTİRİLDİ ---
GROQ_API_KEY = "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"
TEMPLATE_PATH = "template.xlsx"

client = Groq(api_key=GROQ_API_KEY)

SYSTEM_PROMPT = """Sen uzman bir gemi insan kaynaklari (crewing) analistisin.
Verilen CV metnini analiz et ve kesinlikle SADECE JSON formatında cevap ver. 

KRİTİK KURALLAR:
1. Tarihleri GG.AA.YYYY formatında çıkart. Bilinmiyorsa boş bırak "".
2. 'sea_service' kısmındaki gemi tecrübelerini en yeni tarihliden en eski tarihliye sırala.
3. 'stcw_certificates' kısmında key olarak yalnızca şu rakamları kullan: 52, 53, 54, 55, 57, 61, 62, 63.
4. Kesinlikle JSON dışında açıklama, giriş cümlesi veya kapanış notu EKLEME.

BEKLENEN MODEL:
{
  "name": "", "surname": "", "rank_applied": "", "qualification": "", "id_no": "", "dob": "",
  "english_level": "", "phone": "", "phone2": "", "email": "", "address": "", "marital_status": "",
  "children": "", "height": "", "weight": "", "shoes": "", "overall": "", "next_of_kin": "",
  "seaman_book_no": "", "seaman_book_issue": "", "seaman_book_expiry": "",
  "passport_no": "", "passport_issue": "", "passport_expiry": "",
  "medical_issue": "", "medical_expiry": "", "yellow_fever_issue": "", "yellow_fever_expiry": "",
  "usa_visa": "", "schengen_visa": "",
  "sea_service": [{"vsl_type": "", "vsl_name": "", "company": "", "rank": "", "dwt": "", "engine_type": "", "sign_on": "", "sign_off": "", "reason": ""}],
  "stcw_certificates": {"52": {"issue": "", "expiry": ""}},
  "uncertain_fields": [], "warnings": []
}"""

def process_cv_with_llm(text):
    try:
        completion = client.chat.completions.create(
            model="llama3-70b-8192",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": f"CV TEXT:\n{text}"}
            ],
            temperature=0,
        )
        raw_output = completion.choices[0].message.content
        json_match = re.search(r'\{.*\}', raw_output, re.DOTALL)
        if not json_match:
            raise ValueError("AI çıktısında JSON formatı bulunamadı.")
        
        data = CVData.parse_raw(json_match.group(0))
        data.normalize_dates()
        return data
    except Exception as e:
        st.error(f"🤖 AI İşleme Hatası: {str(e)}")
        return None

def fill_excel_template(data: CVData):
    try:
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active
        
        # HUCRE MAPINGI (Buraları kendi template.xlsx dosyanıza göre güncelleyin)
        mapping = {
            "B5": data.name, 
            "D5": data.surname, 
            "F5": data.rank_applied,
            "B6": data.dob, 
            "D6": data.id_no, 
            "F6": data.qualification,
            "B7": data.phone, 
            "D7": data.email, 
            "F7": data.english_level,
        }
        
        for cell, value in mapping.items():
            ws[cell] = value
            
        start_row = 15
        for i, exp in enumerate(data.sea_service):
            row = start_row + i
            ws.cell(row=row, column=2, value=exp.vsl_name)
            ws.cell(row=row, column=3, value=exp.vsl_type)
            ws.cell(row=row, column=4, value=exp.rank)
            ws.cell(row=row, column=5, value=exp.sign_on)
            ws.cell(row=row, column=6, value=exp.sign_off)

        out_filename = f"Processed_{data.name}_{data.surname}.xlsx"
        wb.save(out_filename)
        return out_filename
    except Exception as e:
        st.error(f"📊 Excel Yazma Hatası: {str(e)}")
        return None

# --- STREAMLIT INTERFACE ---
st.title("🚢 BlueFin CV Converter")
st.info("PDF veya DOCX formatındaki CV'leri şirketin standart formatına dönüştürür.")

uploaded_file = st.file_uploader("Lütfen Aday CV'sini Yükleyin", type=['pdf', 'docx'])

if uploaded_file:
    with st.spinner("📝 CV inceleniyor ve veriler ayıklanıyor..."):
        try:
            cv_text = safe_extract(uploaded_file)
            parsed_data = process_cv_with_llm(cv_text)
            
            if parsed_data:
                st.success("✨ Veriler Başarıyla Yakalandı!")
                
                st.subheader("📋 Veri Kontrol Paneli")
                col1, col2, col3 = st.columns(3)
                with col1:
                    edited_name = st.text_input("Ad", parsed_data.name)
                with col2:
                    edited_surname = st.text_input("Soyad", parsed_data.surname)
                with col3:
                    edited_rank = st.text_input("Rütbe", parsed_data.rank_applied)
                
                parsed_data.name = edited_name
                parsed_data.surname = edited_surname
                parsed_data.rank_applied = edited_rank
                
                if parsed_data.warnings:
                    st.warning(f"💡 Notlar: {', '.join(parsed_data.warnings)}")

                if st.button("🚀 Formatı Excel'e Aktar ve İndir"):
                    result_file = fill_excel_template(parsed_data)
                    if result_file:
                        with open(result_file, "rb") as f:
                            st.download_button("📥 EXCEL'İ İNDİR", f, file_name=result_file)

        except Exception as e:
            st.error(f"❌ Sistem Hatası: {str(e)}")
