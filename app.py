import streamlit as st
import google.generativeai as genai
import json
import re
import pandas as pd
from openpyxl import load_workbook
import fitz  # PyMuPDF
from docx import Document
from pydantic import BaseModel, ValidationError
from typing import List, Dict

# ==========================================
# 1. VERİ MODELLERİ (DEĞİŞMEDİ)
# ==========================================
class SeaServiceEntry(BaseModel):
    vsl_type: str = ""; vsl_name: str = ""; company: str = ""; rank: str = ""; dwt: str = ""; engine_type: str = ""; sign_on: str = ""; sign_off: str = ""; reason: str = ""

class STCWCertificate(BaseModel):
    issue: str = ""; expiry: str = ""

class CVData(BaseModel):
    name: str = ""; surname: str = ""; rank_applied: str = ""; qualification: str = ""; id_no: str = ""; dob: str = ""; english_level: str = ""; phone: str = ""; phone2: str = ""; email: str = ""; address: str = ""; marital_status: str = ""; children: str = ""; height: str = ""; weight: str = ""; shoes: str = ""; overall: str = ""; next_of_kin: str = ""; seaman_book_no: str = ""; seaman_book_issue: str = ""; seaman_book_expiry: str = ""; passport_no: str = ""; passport_issue: str = ""; passport_expiry: str = ""; medical_issue: str = ""; medical_expiry: str = ""; yellow_fever_issue: str = ""; yellow_fever_expiry: str = ""; usa_visa: str = ""; schengen_visa: str = ""; sea_service: List[SeaServiceEntry] = []; stcw_certificates: Dict[str, STCWCertificate] = {}; uncertain_fields: List[str] = []; warnings: List[str] = []

    def normalize_dates(self):
        date_fields = ['dob', 'seaman_book_issue', 'seaman_book_expiry', 'passport_issue', 'passport_expiry', 'medical_issue', 'medical_expiry', 'yellow_fever_issue', 'yellow_fever_expiry']
        for field in date_fields:
            val = getattr(self, field)
            if val: setattr(self, field, self._clean_date(val))
        for entry in self.sea_service:
            entry.sign_on = self._clean_date(entry.sign_on); entry.sign_off = self._clean_date(entry.sign_off)

    def _clean_date(self, date_str):
        if not date_str: return ""
        match = re.search(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})', date_str)
        if match:
            d, m, y = match.groups()
            return f"{int(d):02d}.{int(m):02d}.{y}"
        return date_str

# ==========================================
# 2. FONKSİYONLAR
# ==========================================
def safe_extract(uploaded_file):
    ext = uploaded_file.name.split('.')[-1].lower()
    if ext == 'pdf':
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        return "".join([page.get_text() for page in doc])
    elif ext == 'docx':
        doc = Document(uploaded_file)
        return "\n".join([p.text for p in doc.paragraphs])
    raise Exception("Format desteklenmiyor")

# ==========================================
# 3. ANA SAYFA
# ==========================================
st.set_page_config(page_title="BlueFin Gemini Parser", page_icon="⚡")

st.title("🚢 BlueFin CV Analyzer (Flash Speed Edition)")

user_api_key = st.sidebar.text_input("🔑 Gemini API Key Giriniz:", type="password")

if not user_api_key:
    st.warning("Lütfen sol menüdeki kutucuğa Gemini API anahtarınızı girin.")
    st.stop()

# Gemini Konfigürasyonu
genai.configure(api_key=user_api_key)

# --- MODEL DEĞİŞİKLİĞİ BURADA ---
# 'gemini-1.5-pro' yerine 'gemini-1.5-flash' kullandık. Erişim oranı çok daha yüksektir.
try:
    model = genai.GenerativeModel(
        model_name="gemini-1.5-flash", 
        generation_config={"response_mime_type": "application/json"}
    )
except Exception as e:
    st.error(f"Model başlatma hatası: {str(e)}")
    st.stop()

SYSTEM_PROMPT = """Sen uzman bir gemi insan kaynaklari (crewing) analistisin. 
Sana verilen CV metnini analiz et ve SADECE JSON formatında çıktı ver. 

KURALLAR:
1. Tarihleri GG.AA.YYYY formatında çıkart.
2. Sea service tecrübelerini en yeni tarihten eskiye sırala.
3. Sertifika anahtarları sadece şunlar olsun: 52, 53, 54, 55, 57, 61, 62, 63.
4. Dil Türkçe olsun.
"""

uploaded_file = st.file_uploader("CV Yükle", type=['pdf', 'docx'])

if uploaded_file:
    if st.button("🔄 Gemini Flash ile Analiz Et"):
        try:
            with st.spinner("⚡ Gemini Flash ışık hızında analiz ediyor..."):
                text = safe_extract(uploaded_file)
                
                full_prompt = f"{SYSTEM_PROMPT}\n\nANALİZ EDİLECEK CV METNİ:\n{text}"
                response = model.generate_content(full_prompt)
                
                raw_json = response.text
                raw_json = re.sub(r'^```json\s*|\s*```$', '', raw_json.strip(), flags=re.MULTILINE)
                
                data = CVData.parse_raw(raw_json)
                data.normalize_dates()
                
                st.success("✅ Gemini Flash Başarıyla Analiz Etti!")
                st.write(data.dict())
                
        except Exception as e:
            st.error(f"❌ ANALİZ HATASI: {str(e)}")
