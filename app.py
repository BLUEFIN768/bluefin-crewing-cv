import streamlit as st
import os, io, datetime, json, re
import fitz  # PyMuPDF
from docx import Document
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import PatternFill
from PIL import Image
from groq import Groq

APP_VERSION = "v2.7 (Groq Smart Model & Foolproof Syntax)"
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
    
    try:
        available_models = client.models.list().data
        model_ids = [m.id for m in available_models]
        
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
                
        if not active_model:
            for mid in model_ids:
                if "guard" not in mid.lower() and "whisper" not in mid.lower():
                    active_model = mid
                    break
                    
        if not active_model:
            active_model = "llama-3.1-8b-instant"
    except Exception:
        active_model = "llama-3.1-8b-instant" 
        
    # --- YENI: KOPYALAMA HATASI VERMEYEN PROMPT YAPISI ---
    system_prompt = (
        "Sen uzman bir gemi insan kaynaklari (crewing) analistisin. "
        "Sana verilen CV'yi analiz et ve SADECE asagidaki JSON formatinda cikti ver. Baska hicbir aciklama veya metin yazma. "
        "KURALLAR: "
        "1. Tarihleri DD.MM.YYYY formatinda cikar. Bilinmeyen alanlara '' koy. "
        "2. Zabitler icin 4 ay, Ratingler icin 6 aydan kisa hizmetleri 'warnings' dizisine '[Gemi Adi] - Erken Ayrilma' olarak ekle. "
        "3. Okuyamadigin veya emin olmadigin alanlarin isimlerini 'uncertain_fields' dizisine ekle. "
        "4. 'sea_service' icindeki tecrubeleri en guncel en ustte olacak sekilde sirala. "
        "5. 'stcw_certificates' icinde sertifika satir numaralarini (52, 53, 54, 55, 57, 61, 62, 63) key olarak kullan. "
        "BEKLENEN JSON FORMATI: "
        "{"
        "  'name': '', 'surname': '', 'rank_applied': '', 'qualification': '', 'id_no': '', 'dob': '', "
        "  'english_level': '', 'phone': '', 'phone2': '', 'email': '', 'address': '', 'marital_status': '', "
        "  'children': '', 'height': '', 'weight': '', 'shoes': '', 'overall': '', 'next_of_kin': '', "
        "  'seaman_book_no': '', 'seaman_book_issue': '', 'seaman_book_expiry': '', "
        "  'passport_no': '', 'passport_issue': '', 'passport_expiry': '', "
        "  'medical_issue': '', 'medical_expiry': '', 'yellow_fever_issue': '', 'yellow_fever_expiry': '', "
        "  'usa_visa': '', 'schengen_visa': '', "
        "  'sea_service': [ "
        "    {'vsl_type': '', 'vsl_name': '', 'company': '', 'rank': '', 'dwt': '', 'engine_type': '', 'sign_on': '', 'sign_off': '', 'reason': ''} "
        "  ], "
        "  'stcw_certificates': { "
        "    '52': {'issue': '', 'expiry': ''} "
        "  }, "
        "  'uncertain_fields': [], "
        "  'warnings': [] "
        "}"
    ).replace("'", '"')
    
    safe_text = str(text_content).encode('utf-8', 'ignore').decode('utf-8')
    
    try:
        response = client.chat.completions.create(
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"CV METNİ:\n{safe_text}"}
            ],
            model=active_model,
            temperature=0.0
        )
        
        raw_text = response.choices[0].message.content
        
        match = re.search(r'\{.*\}', raw_text, re.DOTALL)
        if match:
            json_str = match.group(0)
            return json.loads(json_str), active_model
        else:
            raise Exception("Yapay zeka geçerli bir format üretmedi.")
            
    except Exception as e:
        raise Exception(f"Okuma Hatasi (Kullanılan Model: {active_model}): {str(e)}")

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
