import streamlit as st
import os, io, datetime, json, re
import fitz  # PyMuPDF
from docx import Document
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import PatternFill
from PIL import Image
from groq import Groq

APP_VERSION = "v2.6 (Groq Smart Model Selector & Regex Parser)"
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
    
    # --- YENİ: AKILLI MODEL DEDEKTÖRÜ (Güvenlik modellerini eler) ---
    try:
        available_models = client.models.list().data
        model_ids = [m.id for m in available_models]
        
        # En iyi ve acik kaynakli metin modelleri oncelik sirasi:
        priority_models = [
            "llama-3.3-70b-versatile",
            "llama-3.1-70b-versatile",
            "llama-3.1-8b-instant",
            "gemma2-9b-it",
            "mixtral-8x7b-32768"
        ]
        
        active_model = None
        for pm in priority_models:
            if pm in model_ids:
                active_model = pm
                break
                
        # Eger hicbiri yoksa, isminde "guard" (guvenlik) ve "whisper" (ses) gecmeyen ilk modeli sec
        if not active_model:
            for mid in model_ids:
                if "guard" not in mid.lower() and "whisper" not in mid.lower():
                    active_model = mid
                    break
                    
        if not active_model:
            active_model = "llama-3.1-8b-instant" # Son guvenlik agi
    except Exception:
        active_model = "llama-3.1-8b-instant" 
        
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
      "children": "", "height": "", "weight": "", "shoes": "", "overall":
