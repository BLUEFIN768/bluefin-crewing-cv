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
    layout="wide"
)

TEMPLATE_PATH = "SEAFARER APPLICATION FORM.xlsx"
RED_FILL = PatternFill(start_color="FF4D4D", end_color="FF4D4D", fill_type="solid")

SYSTEM_PROMPT = """
Sen uzman bir crewing (gemi insan kaynaklari) veri analistisin. 
Sana verilen seafarer CV metnini analiz et ve sirketin 'SEAFARER APPLICATION FORM' sablonuna birebir uyacak sekilde asagidaki JSON formatinda dondur.

KURALLAR:
1. Rutbeleri (Rank) uluslararasi standartlara cevir (Master, Chief Officer, 2nd Officer, Chief Engineer, 2nd Engineer, Bosun, AB, OS, Oiler, Fitter).
2. Pasaport ve Liman Cuzdani numaralarini dikkatli ayikla. Verilis (issue) ve Bitis (expiry) tarihlerini DD.MM.YYYY olarak cikar.
3. Next of Kin (Acil Yakini): "Isim (Yakinlik) - Telefon - Sehir/Adres" formatinda birlestir.
4. STCW sertifikalarini sablon satir numaralariyla eslestir:
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
5. Tarihleri DD.MM.YYYY yap. Bilinmeyen alanlari "" birak.

SADECE gecerli bir JSON dondur:
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
            try:
                t = page.extract_text()
                if t:
                    text += t + "\n"
            except:
                pass
            
            # PDF kilitlenmelerini onleyen guvenli resim tarama
            if not photo:
                try:
                    for img_obj in page.images:
                        p_candidate = Image.open(io.BytesIO(img_obj.data)).convert("RGB")
                        if p_candidate.size[0] > 80 and p_candidate.size[1] > 80:
                            photo = p_candidate
                            break
                except:
                    pass

    return text, photo

def analyze_cv(api_key, text_content):
    clean_key = api_key.strip()
    client = genai.Client(api_key=clean_key)
    
    prompt = f"{SYSTEM_PROMPT}\n\nAsagidaki CV icerigini analiz et:\n\n{text_content}"
    prompt_safe = prompt.encode("ascii", "ignore").decode("ascii")
    
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

def safe_write(ws, row, column, value, fill=None):
    """MergedCell (birlestirilmis hucre) hatasini onleyen guvenli hucre yazici"""
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
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active
    today = datetime.datetime.now()

    safe_write(ws, 7, 3, data.get('name', ''))
    safe_write(ws, 8, 3, data.get('surname', ''))
    safe_write(ws, 9, 3, data.get('rank_applied', ''))
    safe_write(ws, 9, 8, data.get('qualification', ''))
    safe_write(ws, 7, 8, data.get('id_no', ''))
    safe_write(ws, 8, 8, data.get('dob', ''))
    safe_write(ws, 11, 3, data.get('english_level', ''))
    safe_write(ws, 12, 3, data.get('phone', ''))
    safe_write(ws, 12, 8, data.get('phone2', ''))
    safe_write(ws, 13, 3, data.get('email', ''))
    safe_write(ws, 14, 3, data.get('address', ''))
    safe_write(ws, 15, 3, data.get('marital_status', ''))
    safe_write(ws, 15, 8, data.get('children', ''))
    safe_write(ws, 16, 3, data.get('height', ''))
    safe_write(ws, 16, 8, data.get('weight', ''))
    safe_write(ws, 17, 3, data.get('shoes', ''))
    safe_write(ws, 17, 5, data.get('overall', ''))
    safe_write(ws, 18, 4, data.get('next_of_kin', ''))

    safe_write(ws, 20, 5, data.get('seaman_book_no', ''))
    safe_write(ws, 20, 6, data.get('seaman_book_issue', ''))
    safe_write(ws, 20, 8, data.get('seaman_book_expiry', ''))

    safe_write(ws, 21, 5, data.get('passport_no', ''))
    safe_write(ws, 21, 6, data.get('passport_issue', ''))
    safe_write(ws, 21, 8, data.get('passport_expiry', ''))

    safe_write(ws, 22, 6, data.get('medical_issue', ''))
    safe_write(ws, 22, 8, data.get('medical_expiry', ''))

    safe_write(ws, 23, 6, data.get('yellow_fever_issue', ''))
    safe_write(ws, 23, 8, data.get('yellow_fever_expiry', ''))

    safe_write(ws, 24, 5, data.get('usa_visa', ''))
    safe_write(ws, 25, 5, data.get('schengen_visa', ''))

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
                safe_write(ws, r, 8, dates['issue'])
            if dates.get('expiry'):
                is_expired = False
                try:
                    exp_dt = datetime.datetime.strptime(dates['expiry'].replace('/', '.').replace('-', '.'), "%d.%m.%Y")
                    if exp_dt < today:
                        is_expired = True
                except:
                    pass
                fill_to_apply = RED_FILL if is_expired else None
                safe_write(ws, r, 9, dates['expiry'], fill=fill_to_apply)
                if is_expired:
                    safe_write(ws, r, 8, dates.get('issue', ''), fill=RED_FILL)
        except:
            pass

    r_idx = 30
    for exp in data.get('sea_service', []):
        safe_write(ws, r_idx, 1, exp.get('vsl_type', ''))
        safe_write(ws, r_idx, 2, exp.get('vsl_name', ''))
        safe_write(ws, r_idx, 3, exp.get('company', ''))
        safe_write(ws, r_idx, 4, exp.get('rank', ''))
        safe_write(ws, r_idx, 5, exp.get('dwt', ''))
        safe_write(ws, r_idx, 6, exp.get('engine_type', ''))
        safe_write(ws, r_idx, 7, exp.get('sign_on', ''))
        safe_write(ws, r_idx, 8, exp.get('sign_off', ''))
        safe_write(ws, r_idx, 9, exp.get('reason', ''))
        r_idx += 1

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    return out_stream

st.title("Bluefin Crewing - Otomatik CV Donusturucu")
st.markdown("Personel CV lerini standart sirket basvuru formuna (SEAFARER APPLICATION FORM.xlsx) aktarin.")

with st.sidebar:
    st.header("Guvenlik ve Ayarlar")
    api_key = st.text_input("Gemini API Anahtari:", type="password")

if not api_key:
    st.warning("Lutfen sol menuden Gemini API anahtarinizi girin.")
    st.stop()

uploaded_file = st.file_uploader("Donusturulecek CV yi Secin (PDF, Word)", type=["pdf", "docx"])

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    file_name = uploaded_file.name

    st.write(f"Yuklenen Belge: {file_name}")
    if st.button("CV yi Cozumle ve Formu Doldur", type="primary"):
        with st.spinner("Yapay zeka CV yi inceliyor ve formu dolduruyor..."):
            try:
                text_content, photo = extract_content(file_bytes, file_name)
                parsed_data = analyze_cv(api_key, text_content)
                st.session_state['parsed_data'] = parsed_data
                st.session_state['photo'] = photo
                st.success("Cozumleme tamamlandi!")
            except Exception as e:
                st.error(f"Hata olustu: {str(e)}")

    if 'parsed_data' in st.session_state:
        d = st.session_state['parsed_data']
        photo = st.session_state.get('photo', None)

        st.divider()
        st.subheader("Cikarilan Bilgiler ve Onizleme")

        c1, c2, c3 = st.columns([1, 2, 2])
        with c1:
            if photo:
                st.image(photo, caption="Aday Fotografi", width=120)
            else:
                st.info("Fotograf bulunamadi.")
        with c2:
            st.text_input("Adi:", value=d.get('name', ''), key="name")
            st.text_input("Soyadi:", value=d.get('surname', ''), key="surname")
            st.text_input("Gorev (Rank):", value=d.get('rank_applied', ''), key="rank")
            st.text_input("Pasaport No:", value=d.get('passport_no', ''), key="pass_no")
        with c3:
            st.text_input("T.C. / Uyruk:", value=d.get('id_no', ''), key="id_no")
            st.text_input("Dogum Tarihi:", value=d.get('dob', ''), key="dob")
            st.text_input("Liman Cuzdani No:", value=d.get('seaman_book_no', ''), key="sb_no")
            st.text_input("Acil Durum Yakini:", value=d.get('next_of_kin', ''), key="kin")

        excel_data = fill_excel_template(d, photo)
        out_filename = f"Doldurulan_{d.get('name', 'Seafarer')}_{d.get('surname', 'CV')}.xlsx"

        st.download_button(
            label="Doldurulmus SEAFARER APPLICATION FORM.xlsx Indir",
            data=excel_data,
            file_name=out_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
