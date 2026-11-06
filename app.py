import streamlit as st
import os, io, datetime, json, re
import fitz  # PyMuPDF
from docx import Document
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import PatternFill
from PIL import Image
from groq import Groq

APP_VERSION = "v2.5 (Groq Auto-Model & Regex Parser)"
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
    
    # --- YENİ: OTOMATİK MODEL DEDEKTÖRÜ ---
    # Groq'un isim degistirme/silme krizlerine karsi API'ye o an hangi modellerin acik oldugu sorulur.
    try:
        available_models = client.models.list().data
        model_ids = [m.id for m in available_models]
        
        active_model = None
        # 1. Oncelik: En zeki Llama (70B)
        for m in model_ids:
            if "llama" in m.lower() and "70b" in m.lower():
                active_model = m
                break
        # 2. Oncelik: Herhangi bir Llama
        if not active_model:
            for m in model_ids:
                if "llama" in m.lower():
                    active_model = m
                    break
        # 3. Oncelik: Hesaba tanimli rastgele ilk calisan model
        if not active_model and model_ids:
            active_model = model_ids[0]
    except Exception:
        active_model = "llama-3.1-8b-instant" # Son care guvenli liman
        
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
    safe_write(ws, 23, 6, data.get('yellow_fever_issue', ''), fill=get_fill('yellow_fever_issue'))
    safe_write(ws, 23, 8, data.get('yellow_fever_expiry', ''), fill=get_fill('yellow_fever_expiry'))
    safe_write(ws, 24, 5, data.get('usa_visa', ''), fill=get_fill('usa_visa'))
    safe_write(ws, 25, 5, data.get('schengen_visa', ''), fill=get_fill('schengen_visa'))

    if photo:
        photo_buf = io.BytesIO()
        photo_resized = photo.copy()
        photo_resized.thumbnail((115, 140))
        photo_resized.save(photo_buf, format="PNG")
        photo_buf.seek(0)
        img = OpenpyxlImage(photo_buf)
        ws.add_image(img, "H1")

    for row_str, dates in data.get('stcw_certificates', {}).items():
        try:
            r = int(row_str)
            if dates.get('issue'): safe_write(ws, r, 8, dates['issue'], fill=get_fill('stcw_certificates'))
            if dates.get('expiry'): safe_write(ws, r, 9, dates['expiry'], fill=get_fill('stcw_certificates'))
        except:
            pass

    r_idx = 30
    for exp in data.get('sea_service', []):
        row_fill = get_fill('sea_service')
        safe_write(ws, r_idx, 1, exp.get('vsl_type', ''), fill=row_fill)
        safe_write(ws, r_idx, 2, exp.get('vsl_name', ''), fill=row_fill)
        safe_write(ws, r_idx, 3, exp.get('company', ''), fill=row_fill)
        safe_write(ws, r_idx, 4, exp.get('rank', ''), fill=row_fill)
        safe_write(ws, r_idx, 5, exp.get('dwt', ''), fill=row_fill)
        safe_write(ws, r_idx, 6, exp.get('engine_type', ''), fill=row_fill)
        safe_write(ws, r_idx, 7, exp.get('sign_on', ''), fill=row_fill)
        safe_write(ws, r_idx, 8, exp.get('sign_off', ''), fill=row_fill)
        safe_write(ws, r_idx, 9, exp.get('reason', ''), fill=row_fill)
        r_idx += 1

    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    return out_stream

with st.sidebar:
    st.markdown(f"**⚓ Mimari:** `{APP_VERSION}`")
    st.divider()
    st.success("✅ **Sistem Durumu:** Çevrimiçi")
    st.caption("Otomatik Model Bulucu aktif: Sistem artık Groq'taki çalışan modelleri kendi tarar ve 404 hatalarını engeller.")

uploaded_file = st.file_uploader("Donusturulecek CV yi Secin", type=["pdf", "docx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    
    if st.button("🚀 CV'yi Çözümle (Sıfır Hata Modu)", type="primary"):
        with st.spinner("Motor Devrede (Otomatik Model Seçiliyor ve CV Okunuyor)..."):
            try:
                text_content, photo = extract_content(file_bytes, uploaded_file.name)
                parsed_data, used_model = analyze_cv_with_groq(text_content)
                st.session_state['parsed_data'] = parsed_data
                st.session_state['photo'] = photo
                st.session_state['used_model'] = used_model
                st.success(f"Çözümleme {used_model} motoruyla başarıyla tamamlandı!")
            except Exception as e:
                st.error(str(e))

    if 'parsed_data' in st.session_state:
        d = st.session_state['parsed_data']
        photo = st.session_state.get('photo')

        excel_data = fill_excel_template(d, photo)
        safe_name = f"{d.get('rank_applied','')}_{d.get('name','')}_{d.get('surname','')}".replace(' ', '_').upper()
        
        st.download_button(
            label=f"📥 İNDİR: {safe_name}.xlsx",
            data=excel_data,
            file_name=f"{safe_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        st.divider()

        if d.get('warnings'):
            for warning in d['warnings']: st.error(f"🚨 DİKKAT (Hizmet Süresi): {warning}")
        if d.get('uncertain_fields'):
            st.warning(f"⚠️ DİKKAT: Yapay Zeka bazı alanları net okuyamadı. Bu alanlar Excel'de SARI ile işaretlendi.")
            
        st.subheader("📋 Çıkarılan Bilgiler ve Önizleme")

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
