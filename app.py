import streamlit as st
from groq import Groq
import json
import re
import pandas as pd
from openpyxl import load_workbook
from utils import safe_extract
from models import CVData
from pydantic import ValidationError

# --- YAPILANDIRMA ---
st.set_page_config(page_title="BlueFin Crewing Parser", page_icon="🚢", layout="wide")

# Secret'tan API Key alımı (Localde ise direkt yazabilirsiniz: GROQ_API_KEY = "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")
GROQ_API_KEY = st.secrets.get("GROQ_API_KEY", "")
TEMPLATE_PATH = "template.xlsx" # Boş Excel şablonunuzun adı

if not GROQ_API_KEY:
    st.error("API Key eksik! Lütfen Secrets kısmına GROQ_API_KEY ekleyin.")
    st.stop()

client = Groq(api_key=GROQ_API_KEY)

# --- PROMPT TASARIMI ---
SYSTEM_PROMPT = """Sen uzman bir gemi insan kaynaklari (crewing) analistisin.
Verilen CV metnini analiz et ve kesinlikle SADECE JSON formatında cevap ver. 

KRİTİK KURALLAR:
1. Tarihleri GG.AA.YYYY formatında çıkart. Bilinmiyorsa boş bırak "".
2. 'sea_service' kısmındaki gemi tecrübelerini en yeni tarihliden en eski tarihliye sırala.
3. 'stcw_certificates' kısmında key olarak yalnızca şu rakamları kullan: 52, 53, 54, 55, 57, 61, 62, 63.
4. Kesinlikle JSON dışında açıklama, giriş cümlesi veya kapanış notu EKLEME.
5. Yanıtın başında ```json ... ``` blokları olsa bile içerik geçerli bir JSON olmalı.

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
        
        # Markdown JSON bloklarını temizle
        json_match = re.search(r'\{.*\}', raw_output, re.DOTALL)
        if not json_match:
            raise ValueError("AI çıktısında JSON formatı bulunamadı.")
        
        # Pydantic ile valide et ve tarihlerini normalize et
        data = CVData.parse_raw(json_match.group(0))
        data.normalize_dates()
        return data
    except Exception as e:
        st.error(f"LLM Hatası: {str(e)}")
        return None

def fill_excel_template(data: CVData):
    try:
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active
        
        # NOT: Buradaki hücre koordinatlarını (A1, B2 gibi) kendi template.xlsx dosyanıza göre ayarlamalısınız!
        mapping = {
            "B5": data.name, "D5": data.surname, "F5": data.rank_applied,
            "B6": data.dob, "D6": data.id_no, "F6": data.qualification,
            "B7": data.phone, "D7": data.email, "F7": data.english_level,
            # Diğer alanları buraya eklemeye devam edin...
        }
        
        for cell, value in mapping.items():
            ws[cell] = value
            
        # Deneyim Tablosunu Doldur (Örn: Satır 15'ten başladığını varsayalım)
        start_row = 15
        for i, exp in enumerate(data.sea_service):
            row = start_row + i
            ws.cell(row=row, column=2, value=exp.vsl_name)
            ws.cell(row=row, column=3, value=exp.vsl_type)
            ws.cell(row=row, column=4, value=exp.rank)
            ws.cell(row=row, column=5, value=exp.sign_on)
            ws.cell(row=row, column=6, value=exp.sign_off)

        out_filename = f"Parsed_{data.name}_{data.surname}.xlsx"
        wb.save(out_filename)
        return out_filename
    except Exception as e:
        st.error(f"Excel Yazma Hatası: {str(e)}")
        return None

# --- STREAMLIT UI ---
st.title("🚢 BlueFin CV Dönüştürücü")
st.markdown("Sisteme PDF veya DOCX formatında CV yükleyin, kurum içi formatınıza dönüştürelim.")

uploaded_file = st.file_uploader("CV Seçin", type=['pdf', 'docx'])

if uploaded_file:
    with st.spinner("🔍 CV Analiz Ediliyor..."):
        # 1. Text Extract
        try:
            cv_text = safe_extract(uploaded_file)
            
            # 2. LLM Process
            parsed_data = process_cv_with_llm(cv_text)
            
            if parsed_data:
                st.success("✅ Veriler başarıyla ayıklandı!")
                
                # Düzenlenebilir Alanlar
                col1, col2 = st.columns(2)
                with col1:
                    edited_name = st.text_input("Ad", parsed_data.name)
                    edited_surname = st.text_input("Soyad", parsed_data.surname)
                    edited_rank = st.text_input("Rank", parsed_data.rank_applied)
                
                # Veriyi güncelle
                parsed_data.name = edited_name
                parsed_data.surname = edited_surname
                parsed_data.rank_applied = edited_rank
                
                # Warningleri Göster
                if parsed_data.warnings:
                    st.warning(f"Uyarılar: {', '.join(parsed_data.warnings)}")

                # 3. Excel Export
                if st.button("💾 Excel Olarak İndir"):
                    result_file = fill_excel_template(parsed_data)
                    if result_file:
                        with open(result_file, "rb") as f:
                            st.download_button("📥 Dosyayı Hemen İndir", f, file_name=result_file)

        except Exception as e:
            st.error(f"Sistem Hatası: {str(e)}")
