import streamlit as st
import os, io, datetime, json, re
import fitz  # PyMuPDF
from docx import Document
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import PatternFill
from PIL import Image
from groq import Groq

APP_VERSION = "v2.2 (Groq Regex Parser & UTF-8 Safe)"
TEMPLATE_PATH = "SEAFARER APPLICATION FORM.xlsx"
RED_FILL = PatternFill(start_color="FF4D4D", end_color="FF4D4D", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# API Anahtarinizi dogrudan koda gomduk
GROQ_API_KEY = "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"

st.set_page_config(page_title="Bluefin Crewing - Enterprise CV Parser", layout="wide")

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

def analyze_cv_with_groq(text_content):
    client = Groq(api_key=GROQ_API_KEY)
    
    system_prompt = """
    Sen uzman bir gemi insan kaynaklari (crewing) analistisin. 
    Sana verilen CV'yi analiz et ve SADECE asagidaki JSON formatinda cikti ver. Baska hicbir aciklama veya metin yazma.
    
    KURALLAR:
    1. Tarihleri DD.MM.YYYY formatinda cikar. Bilinmeyen alanlara "" koy.
    2. Zabitler icin 4 ay, Ratingler icin 6 aydan kisa hizmetleri "warnings" dizisine "[Gemi Adi] - Erken Ayrilma" olarak ekle.
    3. Okuyamadigin veya emin olmadigin alanlarin isimlerini "uncertain_fields" dizisine ekle.
    4. "sea_service" icindeki tecrubeleri en guncel en ustte olacak sekilde sirala.
    5. "stcw_certificates" icinde sertifika satir numaralarini (52, 53, 54, 55, 57, 61, 62, 63) key olarak kullan.
    
    BEKLENEN JSON FORMATI:
    {
      "name": "", "surname": "", "rank_applied": "", "qualification": "", "id_no": "", "dob": "",
      "english_level": "", "phone": "", "phone2": "", "email": "", "address": "", "marital_status": "",
      "children": "", "height": "", "weight": "", "shoes": "", "overall": "", "next_of_kin": "",
      "seaman_book_no": "", "seaman_book_issue": "", "seaman_book_expiry": "",
      "passport_no": "", "passport_issue": "", "passport_expiry": "",
      "medical_issue": "", "medical_expiry": "", "yellow_fever_issue": "", "yellow_fever_expiry": "",
      "usa_visa": "", "schengen_visa": "",
      "sea_service": [
        {"vsl_type": "", "vsl_name": "", "company": "", "rank": "", "dwt": "", "engine_type": "", "sign_on": "", "sign_off": "", "reason": ""}
      ],
      "stcw_certificates": {
        "52": {"issue": "", "expiry": ""}
      },
      "uncertain_fields": [],
      "warnings": []
    }
    """
    
    # Turkce karakter cokmelerini engellemek icin metni UTF-8'e zorluyoruz
    safe_text = str(text_content).encode('utf-8', 'ignore').decode('utf-8')
    
    try:
        response = client.chat.completions.create(
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"CV METNİ:\n{safe_text}"}
            ],
            model="llama-3.3-70b-versatile",
            temperature=0.0
            # 400 hatasini engellemek icin strict JSON format zorlamasini kaldirdik
        )
        
        raw_text = response.choices[0].message.content
        
        # Yapay zeka fazladan metin uretse bile icindeki JSON blogunu (veriyi) zorla söküp aliyoruz
        match = re
