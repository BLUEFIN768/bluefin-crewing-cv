import streamlit as st
import os, io, json, re, datetime, zipfile, time
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import PatternFill
from PIL import Image
import pypdf
from docx import Document
from google import genai

st.set_page_config(
    page_title="Bluefin Crewing - Seafarer CV Converter",
    page_icon="⚓",
    layout="wide"
)

TEMPLATE_PATH = "SEAFARER APPLICATION FORM.xlsx"
RED_FILL = PatternFill(start_color="FF4D4D", end_color="FF4D4D", fill_type="solid")

SYSTEM_PROMPT = """
Sen uzman bir crewing (gemi insan kaynakları) veri analistisin. 
Sana verilen seafarer CV metnini analiz et ve şirketin 'SEAFARER APPLICATION FORM' şablonuna birebir uyacak şekilde aşağıdaki JSON formatında döndür.

KURALLAR:
1. Rütbeleri (Rank) uluslararası standartlara çevir (Master, Chief Officer, 2nd Officer, Chief Engineer, 2nd Engineer, Bosun, AB, OS, Oiler, Fitter).
2. Pasaport ve Liman Cüzdanı numaralarını dikkatli ayıkla. Veriliş (issue) ve Bitiş (expiry) tarihlerini DD.MM.YYYY olarak çıkar.
3. Next of Kin (Acil Yakını): "İsim (Yakınlık) - Telefon - Şehir/Adres" formatında birleştir.
4. STCW sertifikalarını şablon satır numaralarıyla eşleştir:
   52: Personal Survival Techniques
   53: Fire Prevention and Fire Fighting
   54: Elementary First Aid
   55: Personal Safety and Social Responsibilities (PSSR)
   56: Advanced Fire Fighting
   57: Proficiency in Survival Craft
   61: Security Awareness
   62: Designated Security Duties (DSD)
   63: Ship Security Officer (SSO)
   68: ECDIS
5. Tarihleri DD.MM.YYYY yap. Bilinmeyen alanları "" bırak.

SADECE geçerli bir JSON döndür:
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

    return text, photo

def analyze_cv(api_key, text_content):
    clean_key = api_key.strip()
    client = genai.Client(api_key=clean_key)
    
    prompt = f"{SYSTEM_PROMPT}\n\nAşağıdaki CV içeriğini analiz et:\n\n{text_content}"
    prompt_safe = prompt.encode("utf-8", "ignore").decode("utf-8")
    
    # 503 yoğunluk hatası gelirse otomatik olarak 3 saniye bekleyip 3 kez dener
    max_retries = 3
    last_error = None
    
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model="gemini-3.8-flash",
                contents=prompt_safe,
            )
            raw_text = response.text.strip()
            match = re.search(r"\{.*\}", raw_text, re.DOTALL)
            if match:
                return json.loads(match.group(0))
            return json.loads(raw_text)
        except Exception as e:
            last_error = e
            err_msg = str(e)
            if "503" in err_msg or "UNAVAILABLE" in err_msg:
                time.sleep(3)
                continue
            raise e
            
    raise last_error

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
            if dates.get('issue'):
                ws.cell(row=r, column=8, value=dates['issue'])
            if dates.get('expiry'):
                exp_cell = ws.cell(row=r, column=9, value=dates['expiry'])
                try:
                    exp_dt = datetime.datetime.strptime(dates['expiry'].replace('/', '.').replace('-', '.'), "%d.%m.%Y")
                    if exp_dt < today:
                        exp_cell.fill = RED_FILL
                        ws.cell(row=r, column=8).fill = RED_FILL
                except:
                    pass
        except:
            pass

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

# Arayüz
st.title("⚓ Bluefin Crewing - Otomatik CV Dönüştürücü")
st.markdown("Personel CV'lerini standart şirket başvuru formuna (**SEAFARER APPLICATION FORM.xlsx**) aktarın.")

with st.sidebar:
    st.header("🔑 Güvenlik & Ayarlar")
    api_key = st.text_input("Gemini API Anahtarı:", type="password")

if not api_key:
    st.warning("Lütfen sol menüden Gemini API anahtarınızı girin.")
    st.stop()

uploaded_file = st.file_uploader("Dönüştürülecek CV'yi Seçin (PDF, Word)", type=["pdf", "docx"])

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    file_name = uploaded_file.name

    st.write(f"📁 **Yüklenen Belge:** {file_name}")
    if st.button("🚀 CV'yi Çözümle ve Formu Doldur", type="primary"):
        with st.spinner("Yapay zeka CV'yi inceliyor ve formu dolduruyor..."):
            try:
                text_content, photo = extract_content(file_bytes, file_name)
                parsed_data = analyze_cv(api_key, text_content)
                st.session_state['parsed_data'] = parsed_data
                st.session_state['photo'] = photo
                st.success("Çözümleme tamamlandı!")
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
                st.image(photo, caption="Aday Fotoğrafı", width=120)
            else:
                st.info("Fotoğraf bulunamadı.")
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
