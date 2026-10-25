import streamlit as st
import os, io, json, re, datetime, zipfile
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import PatternFill
from PIL import Image
import pypdf
from docx import Document
import google.generativeai as genai

st.set_page_config(
    page_title="Bluefin Crewing - Seafarer CV Converter",
    page_icon="⚓",
    layout="wide"
)

# 1. ŞİRKET ŞİFRE KONTROLÜ
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets.get("APP_PASSWORD", "bluefin2026!"):
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct", False):
        return True

    st.markdown("### ⚓ Bluefin Crewing Giriş Paneli")
    st.text_input("Şirket Giriş Şifresi:", type="password", on_change=password_entered, key="password")
    if "password_correct" in st.session_state and not st.session_state["password_correct"]:
        st.error("😕 Hatalı şifre. Lütfen şirket yetkilisine danışın.")
    return False

if not check_password():
    st.stop()

# 2. AYARLAR & SABİTLER
TEMPLATE_PATH = "SEAFARER APPLICATION FORM.xlsx"
RED_FILL = PatternFill(start_color="FF4D4D", end_color="FF4D4D", fill_type="solid")

SYSTEM_PROMPT = """
Sen uzman bir crewing (gemi insan kaynakları) veri analistisin. 
Sana verilen seafarer CV metnini/belgelerini analiz et ve şirketin 'SEAFARER APPLICATION FORM' şablonuna birebir uyacak şekilde aşağıdaki JSON formatında döndür.

KURALLAR:
1. Rütbeleri (Rank) kesinlikle uluslararası standartlara çevir:
   - "Güverte Lostroması / Boatswain" -> "Bosun"
   - "Usta Gemici / Able Seaman" -> "AB"
   - "Gemici / Ordinary Seaman" -> "OS"
   - "Başmühendis / Çarkçıbaşı" -> "Chief Engineer"
   - "1. Zabit / 2. Kaptan" -> "Chief Officer"
   - "2. Mühendis" -> "2nd Engineer"
   - "3. Mühendis" -> "3rd Engineer"
   - "Vardiya Zabiti" -> "2nd Officer" veya "3rd Officer"
   - "Süvari / Kaptan" -> "Master"
   - "Fitter / Kaynakçı" -> "Fitter"
   - "Yağcı" -> "Oiler"
   - "Elektro-Teknik Tayfası" -> "ETR"
2. Pasaport ve Liman Cüzdanı numaralarını dikkatli ayıkla. "DETAILS" kelimesini numara sanma; gerçek seri numarasını al. Veriliş (issue) ve Bitiş (expiry) tarihlerini DD.MM.YYYY olarak çıkar.
3. Next of Kin (Acil Yakını): "İsim (Yakınlık) - Telefon - Şehir/Adres" formatında birleştir.
4. STCW sertifikalarını şablon satır numaralarıyla eşleştir:
   - 52: Personal Survival Techniques
   - 53: Fire Prevention and Fire Fighting
   - 54: Elementary First Aid
   - 55: Personal Safety and Social Responsibilities (PSSR)
   - 56: Advanced Fire Fighting
   - 57: Proficiency in Survival Craft and Rescue Boats
   - 58: Fast Rescue Boats
   - 59: Medical First Aid
   - 60: Medical Care
   - 61: Security Awareness
   - 62: Designated Security Duties
   - 63: Ship Security Officer (SSO)
   - 67: Leadership and Teamwork
   - 68: ECDIS
   - 69: ARPA
   - 70: Radar
   - 71: Basic Training Oil and Chemical Tanker
   - 72: Advanced Oil Tanker
   - 73: Advanced Chemical Tanker
   - 82: High Voltage
   - 86: Security-Related Familiarization
5. Tarihleri mutlaka DD.MM.YYYY formatına çevir. Bilinmeyen alanları boş string ("") bırak.

Çıktıyı SADECE geçerli bir JSON olarak ver:
{
  "name": "",
  "surname": "",
  "rank_applied": "",
  "qualification": "",
  "id_no": "",
  "dob": "",
  "english_level": "",
  "phone": "",
  "phone2": "",
  "email": "",
  "address": "",
  "marital_status": "",
  "children": "",
  "height": "",
  "weight": "",
  "shoes": "",
  "overall": "",
  "next_of_kin": "",
  "seaman_book_no": "",
  "seaman_book_issue": "",
  "seaman_book_expiry": "",
  "passport_no": "",
  "passport_issue": "",
  "passport_expiry": "",
  "medical_issue": "",
  "medical_expiry": "",
  "yellow_fever_issue": "",
  "yellow_fever_expiry": "",
  "usa_visa": "",
  "schengen_visa": "",
  "sea_service": [
    {
      "vsl_type": "",
      "vsl_name": "",
      "company": "",
      "rank": "",
      "dwt": "",
      "engine_type": "",
      "sign_on": "",
      "sign_off": "",
      "reason": ""
    }
  ],
  "stcw_certificates": {
    "52": {"issue": "", "expiry": ""},
    "53": {"issue": "", "expiry": ""},
    "54": {"issue": "", "expiry": ""},
    "55": {"issue": "", "expiry": ""},
    "57": {"issue": "", "expiry": ""},
    "61": {"issue": "", "expiry": ""},
    "62": {"issue": "", "expiry": ""},
    "63": {"issue": "", "expiry": ""}
  }
}
"""

def extract_content(file_bytes, file_name):
    ext = os.path.splitext(file_name)[1].lower()
    text = ""
    photo = None
    
    if ext == ".docx":
        doc = Document(io.BytesIO(file_bytes))
        for p in doc.paragraphs:
            if p.text.strip():
                text += p.text + "\n"
        for t in doc.tables:
            for r in t.rows:
                row_t = " | ".join([c.text.strip().replace('\n', ' ') for c in r.cells if c.text.strip()])
                if row_t:
                    text += row_t + "\n"
        try:
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
                imgs = [n for n in z.namelist() if n.startswith('word/media/')]
                if imgs:
                    img_data = z.read(imgs[0])
                    photo = Image.open(io.BytesIO(img_data)).convert("RGB")
        except:
            pass

    elif ext == ".pdf":
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
            if not photo and len(page.images) > 0:
                try:
                    img_obj = page.images[0]
                    p_candidate = Image.open(io.BytesIO(img_obj.data)).convert("RGB")
                    if p_candidate.size[0] > 70 and p_candidate.size[1] > 70:
                        photo = p_candidate
                except:
                    pass

    elif ext in [".jpg", ".jpeg", ".png"]:
        photo = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        text = "IMAGE_FILE"

    return text, photo

def analyze_cv_with_gemini(api_key, text_content, image=None):
    genai.configure(api_key=api_key)
    selected_model = "gemini-3.8-flash"
    model = genai.GenerativeModel(selected_model)
    content = [SYSTEM_PROMPT]
    
    if image and text_content == "IMAGE_FILE":
        content.append(image)
        content.append("Bu görseldeki tüm CV verilerini oku ve JSON formatında döndür.")
    else:
        content.append(f"Aşağıdaki CV içeriğini analiz et:\n\n{text_content}")
        if image:
            content.append(image)

    response = model.generate_content(content)
    raw_text = response.text.strip()
    match = re.search(r"\{.*\}", raw_text, re.DOTALL)
    if match:
        return json.loads(match.group(0))
    return json.loads(raw_text)

def fill_excel_template(data, photo=None):
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active
    today = datetime.datetime.now()

    ws.cell(row=7, column=3, value=data.get('name', ''))
    ws.cell(row=8, column=3, value=data.get('surname', ''))
    ws.cell(row=9, column=3, value=data.get('rank_applied', ''))
    ws.cell(row=9, column=8, value=data.get('qualification', ''))
    ws.cell(row=7, column=8, value=data.get('id_no', ''))
    ws.cell(row=8, column=8, value=data.get('dob', ''))
    ws.cell(row=11, column=3, value=data.get('english_level', ''))
    ws.cell(row=12, column=3, value=data.get('phone', ''))
    ws.cell(row=12, column=8, value=data.get('phone2', ''))
    ws.cell(row=13, column=3, value=data.get('email', ''))
    ws.cell(row=14, column=3, value=data.get('address', ''))
    ws.cell(row=15, column=3, value=data.get('marital_status', ''))
    ws.cell(row=15, column=8, value=data.get('children', ''))
    ws.cell(row=16, column=3, value=data.get('height', ''))
    ws.cell(row=16, column=8, value=data.get('weight', ''))
    ws.cell(row=17, column=3, value=data.get('shoes', ''))
    ws.cell(row=17, column=5, value=data.get('overall', ''))
    ws.cell(row=18, column=4, value=data.get('next_of_kin', ''))

    ws.cell(row=20, column=5, value=data.get('seaman_book_no', ''))
    ws.cell(row=20, column=6, value=data.get('seaman_book_issue', ''))
    ws.cell(row=20, column=8, value=data.get('seaman_book_
