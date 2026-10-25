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
        if st.session_state["password"] == st.secrets.get("APP_PASSWORD", "bluefin2026"):
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
            if p.text.strip(): text += p.text + "\n"
        for t in doc.tables:
            for r in t.rows:
                row_t = " | ".join([c.text.strip().replace('\n', ' ') for c in r.cells if c.text.strip()])
                if row_t: text += row_t + "\n"
        try:
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
                imgs = [n for n in z.namelist() if n.startswith('word/media/')]
                if imgs:
                    img_data = z.read(imgs[0])
                    photo = Image.open(io.BytesIO(img_data)).convert("RGB")
        except: pass

    elif ext == ".pdf":
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        for page in reader.pages:
            t = page.extract_text()
            if t: text += t + "\n"
            if not photo and len(page.images) > 0:
                try:
                    img_obj = page.images[0]
                    p_candidate = Image.open(io.BytesIO(img_obj.data)).convert("RGB")
                    if p_candidate.size[0] > 70 and p_candidate.size[1] > 70:
                        photo = p_candidate
                except: pass

    elif ext in [".jpg", ".jpeg", ".png"]:
        photo = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        text = "IMAGE_FILE"

    return text, photo

def analyze_cv_with_gemini(api_key, text_content, image=None):
    genai.configure(api_key=api_key)
    
    selected_model = "gemini-2.5-flash"
    try:
        models = [m.name for m in genai.list_models() if "generateContent" in m.supported_generation_methods]
        for candidate in ["models/gemini-2.5-flash", "models/gemini-2.0-flash", "models/gemini-1.5-flash"]:
            if candidate in models:
                selected_model = candidate
                break
    except: pass

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
    ws.cell(row=20, column=8, value=data.get('seaman_book_expiry', ''))

    ws.cell(row=21, column=5, value=data.get('passport_no', ''))
    ws.cell(row=21, column=6, value=data.get('passport_issue', ''))
    ws.cell(row=21, column=8, value=data.get('passport_expiry', ''))

    ws.cell(row=22, column=6, value=data.get('medical_issue', ''))
    ws.cell(row=22, column=8, value=data.get('medical_expiry', ''))

    ws.cell(row=23, column=6, value=data.get('yellow_fever_issue', ''))
    ws.cell(row=23, column=8, value=data.get('yellow_fever_expiry', ''))

    ws.cell(row=24, column=5, value=data.get('usa_visa', ''))
    ws.cell(row=25, column=5, value=data.get('schengen_visa', ''))

    if photo:
        photo_buf = io.BytesIO()
        photo_resized = photo.copy()
        photo_resized.thumbnail((115, 140))
        photo_resized.save(photo_buf, format="PNG")
        photo_buf.seek(0)
        img = OpenpyxlImage(photo_buf)
        ws.add_image(img, "H1")

    stcw_data = data.get('stcw_certificates', {})
    for row_str, dates in stcw_data.items():
        try:
            r = int(row_str)
            if dates.get('issue'): ws.cell(row=r, column=8, value=dates['issue'])
            if dates.get('expiry'):
                exp_cell = ws.cell(row=r, column=9, value=dates['expiry'])
                try:
                    exp_dt = datetime.datetime.strptime(dates['expiry'].replace('/', '.').replace('-', '.'), "%d.%m.%Y")
                    if exp_dt < today:
                        exp_cell.fill = RED_FILL
                        ws.cell(row=r, column=8).fill = RED_FILL
                except: pass
        except: pass

    r_idx = 30
    for exp in data.get('sea_service', []):
        ws.cell(row=r_idx, column=1, value=exp.get('vsl_type', ''))
        ws.cell(row=r_idx, column=2, value=exp.get('vsl_name', ''))
        ws.cell(row=r_idx, column=3, value=exp.get('company', ''))
        ws.cell(row=r_idx, column=4, value=exp.get('rank', ''))
        ws.cell(row=r_idx, column=5, value=exp.get('dwt', ''))
        ws.cell(row=r_idx, column=6, value=exp.get('engine_type', ''))
        ws.cell(row=r_idx, column=7, value=exp.get('sign_on', ''))
        ws.cell(row=r_idx, column=8, value=exp.get('sign_off', ''))
        ws.cell(row=r_idx, column=9, value=exp.get('reason', ''))
        r_idx += 1

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    return out_stream

# 3. KULLANICI ARAYÜZÜ
st.title("⚓ Bluefin Crewing - Otomatik CV Dönüştürücü")
st.markdown("Personel CV'lerini standart şirket başvuru formuna (**SEAFARER APPLICATION FORM.xlsx**) aktarın.")

api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    st.error("Sistem API anahtarı yapılandırılmamış. Lütfen yöneticiyle iletişime geçin.")
    st.stop()

uploaded_file = st.file_uploader("Dönüştürülecek CV'yi Seçin (PDF, Word, JPG, PNG)", type=["pdf", "docx", "jpg", "jpeg", "png"])

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    file_name = uploaded_file.name

    st.write(f"📁 **Yüklenen Belge:** {file_name}")
    if st.button("🚀 CV'yi Çözümle ve Formu Doldur", type="primary"):
        with st.spinner("Yapay zeka CV'yi inceliyor, unvan ve sertifikaları eşleştiriyor..."):
            try:
                text_content, photo = extract_content(file_bytes, file_name)
                parsed_data = analyze_cv_with_gemini(api_key, text_content, photo)
                st.session_state['parsed_data'] = parsed_data
                st.session_state['photo'] = photo
                st.success("Çözümleme tamamlandı! Bilgileri kontrol edip Excel dosyanızı indirebilirsiniz.")
            except Exception as e:
                st.error(f"Hata oluştu: {str(e)}")

    if 'parsed_data' in st.session_state:
        d = st.session_state['parsed_data']
        photo = st.session_state.get('photo', None)

        st.divider()
        st.subheader("📋 Çıkarılan Bilgiler ve Önizleme")

        c1, c2, c3 = st.columns([1, 2, 2])
        with c1:
            if photo:
                st.image(photo, caption="Aday Fotoğrafı (Excel H1'e eklenecek)", width=120)
            else:
                st.info("Vesikalık fotoğraf bulunamadı.")
        with c2:
            st.text_input("Adı:", value=d.get('name', ''), key="name")
            st.text_input("Soyadı:", value=d.get('surname', ''), key="surname")
            st.text_input("Görev (Rank):", value=d.get('rank_applied', ''), key="rank")
            st.text_input("Pasaport No:", value=d.get('passport_no', ''), key="pass_no")
        with c3:
            st.text_input("T.C. / Uyruk:", value=d.get('id_no', ''), key="id_no")
            st.text_input("Doğum Tarihi:", value=d.get('dob', ''), key="dob")
            st.text_input("Liman Cüzdanı No:", value=d.get('seaman_book_no', ''), key="sb_no")
            st.text_input("Acil Durum Yakını:", value=d.get('next_of_kin', ''), key="kin")

        excel_data = fill_excel_template(d, photo)
        out_filename = f"Doldurulan_{d.get('name', 'Seafarer')}_{d.get('surname', 'CV')}.xlsx"

        st.download_button(
            label="📥 Doldurulmuş SEAFARER APPLICATION FORM.xlsx İndir",
            data=excel_data,
            file_name=out_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
