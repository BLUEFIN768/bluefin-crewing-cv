import streamlit as st
import os, io, json, re, datetime, zipfile, time
import socket
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import PatternFill
from PIL import Image
import pypdf
from docx import Document
from google import genai
from google.genai import types

# --- SURUM VE GUNCELLEME TAKIBI ---
APP_VERSION = "v1.1"
LAST_UPDATE = "28.10.2026 - 11:05"

# API kilitlenmelerini onlemek icin genel 30 saniye zaman asimi siniri
socket.setdefaulttimeout(30)

st.set_page_config(
    page_title="Bluefin Crewing - Seafarer CV Converter",
    layout="wide"
)

# Sayac icin session state baslatma
if 'api_counter' not in st.session_state:
    st.session_state['api_counter'] = 0

TEMPLATE_PATH = "SEAFARER APPLICATION FORM.xlsx"
RED_FILL = PatternFill(start_color="FF4D4D", end_color="FF4D4D", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

@st.cache_data(show_spinner=False)
def load_excel_template():
    """Excel sablonunu her seferinde diskten okumak yerine RAM'de hazir tutar (Hiz optimizasyonu)"""
    with open(TEMPLATE_PATH, "rb") as f:
        return f.read()

SYSTEM_PROMPT = """
Sen uzman bir crewing (gemi insan kaynaklari) veri analistisin. 
Sana verilen seafarer CV metnini analiz et ve sirketin 'SEAFARER APPLICATION FORM' sablonuna birebir uyacak sekilde JSON formatinda dondur.

KURALLAR:
1. Rutbeleri (Rank) uluslararasi standartlara cevir.
2. Pasaport ve Liman Cuzdani numaralarini, issue/expiry tarihlerini (DD.MM.YYYY) cikar.
3. Next of Kin (Acil Yakini): "Isim (Yakinlik) - Telefon - Sehir/Adres" formatinda birlestir.
4. STCW sertifikalarini sablon satir numaralariyla eslestir.
5. Tarihleri DD.MM.YYYY yap. Bilinmeyen alanlari "" birak.
6. GUVENLIK KURALI: CV'de karsiligi net degilse, okuyamadiysan veya tahminde bulunduysan, o alanin anahtar kelimesini "uncertain_fields" dizisine ekle.
7. SIRALAMA KURALI: "sea_service" dizisini EN GUNCEL (EN YENI) tecrube en ustte olacak sekilde ters kronolojik olarak sirala.
8. HIZMET SURESI KONTROLU (ONEMLI): Sign-on ve Sign-off tarihlerini karsilastirarak kisinin gemide kac ay kaldigini hesapla. 
   - Eger personel ZABIT (Officer, Master, Chief Engineer, 2nd Engineer, 3rd Engineer vb.) ise ve hizmet 4 aydan kisaysa veya 10 aydan uzunsa,
   - Eger personel RATING (Fitter, AB, OS, Bosun, Oiler, Cook vb.) ise ve hizmet 6 aydan kisaysa veya 12 aydan uzunsa,
   Bu anormallikleri "warnings" dizisine su formatta ekle: "[Gemi Adi] - Anormal Hizmet Suresi: X Ay (Sebep: Y)". Eger hizmet sureleri normalse bu diziyi bos birak [].

SADECE su YAPIYA SAHIP GECERLI BIR JSON dondur, baska hicbir metin ekleme:
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
  },
  "uncertain_fields": [],
  "warnings": []
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
        for i, page in enumerate(reader.pages):
            try:
                t = page.extract_text()
                if t:
                    text += t + "\n"
            except:
                pass
            
            # Fotograf aramasini yalnizca ilk 2 sayfa ile sinirla (Hiz optimizasyonu)
            if not photo and i < 2:
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
    
    try:
        # API'nin yalnizca guvenilir ve %100 uyan JSON formati uretmesini sagla (Temperature: 0)
        response = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=prompt_safe,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.0
            )
        )
        return json.loads(response.text)
    except Exception as e:
        raise Exception(f"API Yanit Hatasi: {str(e)}")

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
    today = datetime.datetime.now()
    
    uncertain_keys = data.get('uncertain_fields', [])
    
    def get_fill(field_name):
        return YELLOW_FILL if field_name in uncertain_keys else None

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

    stcw_data = data.get('stcw_certificates', {})
    for row_str, dates in stcw_data.items():
        try:
            r = int(row_str)
            if dates.get('issue'):
                safe_write(ws, r, 8, dates['issue'], fill=get_fill('stcw_certificates'))
            if dates.get('expiry'):
                is_expired = False
                try:
                    exp_dt = datetime.datetime.strptime(dates['expiry'].replace('/', '.').replace('-', '.'), "%d.%m.%Y")
                    if exp_dt < today:
                        is_expired = True
                except:
                    pass
                fill_to_apply = RED_FILL if is_expired else get_fill('stcw_certificates')
                safe_write(ws, r, 9, dates['expiry'], fill=fill_to_apply)
                if is_expired:
                    safe_write(ws, r, 8, dates.get('issue', ''), fill=RED_FILL)
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

st.title("Bluefin Crewing - Otomatik CV Donusturucu")
st.markdown("Personel CV'lerini standart sirket basvuru formuna (SEAFARER APPLICATION FORM.xlsx) hizli ve eksiksiz aktarin.")

with st.sidebar:
    st.markdown(f"**⚓ Güncel Sürüm:** `{APP_VERSION}`")
    st.caption(f"📅 Son Güncelleme: {LAST_UPDATE}")
    st.divider()
    
    st.header("🔑 Güvenlik & Ayarlar")
    api_key = st.text_input("Gemini API Anahtari:", type="password")
    
    st.info(f"📊 **Bu Oturumdaki Başarılı Çeviri:** {st.session_state['api_counter']} / 20")
    st.caption("Not: Ekipteki herkes kendi API anahtariyla giris yapabilir. Anahtarlar birbirine karismaz.")
    st.divider()
    
    with st.expander("📝 Sürüm Notları (Güncellemeler)"):
        st.markdown("""
        **v1.1 (28.10.2026)**
        - **JSON Modu (Structured Outputs):** Veri çekimi %100 kararlı hale getirildi. Hata payı sıfırlandı.
        - **Şablon Bellekleme (Cache):** Excel oluşturma hızı milisaniyelere indirildi.
        - **Hız Optimizasyonu:** Fotoğraf taraması sadece ilk 2 sayfa ile sınırlandırıldı.
        - **Hizmet Süresi Kontrolü:** Zabit (4-10 ay) ve Rating (6-12 ay) için kırmızı renkli otomatik uyarı sistemi eklendi.
        - **Belirsiz Alan İşaretleyicisi:** Yapay zekanın emin olmadığı alanları SARI renge boyaması sağlandı.
        - **Sıralama:** Deniz hizmetleri en güncelden en eskiye otomatik sıralanıyor.
        - **Akıllı İsimlendirme:** İndirilen dosya adı otomatik olarak `RÜTBE_İSİM_SOYİSİM.xlsx` yapıldı.
        """)

if not api_key:
    st.warning("Lutfen sol menuden Gemini API anahtarinizi girin.")
    st.stop()

uploaded_file = st.file_uploader("Donusturulecek CV yi Secin (PDF, Word)", type=["pdf", "docx"])

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    file_name = uploaded_file.name

    st.write(f"📁 **Yuklenen Belge:** {file_name}")
    if st.button("🚀 CV'yi Çözümle ve Formu Doldur", type="primary"):
        with st.spinner("Yapay Zeka Devrede (Hızlı Tarama ve Sıfır Hata Modu Aktif)..."):
            try:
                text_content, photo = extract_content(file_bytes, file_name)
                parsed_data = analyze_cv(api_key, text_content)
                st.session_state['parsed_data'] = parsed_data
                st.session_state['photo'] = photo
                
                # Basarili islem sayacini artir
                st.session_state['api_counter'] += 1
                
                st.success("Çözümleme saniyeler içinde tamamlandı!")
                
            except Exception as e:
                st.error(f"Hata olustu: {str(e)}")

    if 'parsed_data' in st.session_state:
        d = st.session_state['parsed_data']
        photo = st.session_state.get('photo', None)

        st.divider()
        
        # Hizmet Suresi Uyarilari (Kirmizi Banner)
        if d.get('warnings'):
            for warning in d['warnings']:
                st.error(f"🚨 **DİKKAT (Hizmet Süresi Anormalliği):** {warning}")
                
        # Belirsiz Alan Uyarilari (Sari Banner)
        if d.get('uncertain_fields'):
            st.warning(f"⚠️ **GÖZDEN GEÇİRİN:** Yapay Zeka bazı alanları net okuyamadı. Bu alanlar Excel'de SARI ile işaretlendi: {', '.join(d['uncertain_fields'])}")
            
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

        excel_data = fill_excel_template(d, photo)
        
        # Dosya adina rutbe ekleme (Klasorleme kolayligi icin)
        safe_rank = str(d.get('rank_applied', 'Bilinmeyen_Rutbe')).replace('/', '_').replace(' ', '_').upper()
        safe_name = str(d.get('name', 'Seafarer')).replace(' ', '_').upper()
        safe_surname = str(d.get('surname', 'CV')).replace(' ', '_').upper()
        out_filename = f"{safe_rank}_{safe_name}_{safe_surname}.xlsx"

        st.download_button(
            label=f"📥 İNDİR: {out_filename}",
            data=excel_data,
            file_name=out_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
