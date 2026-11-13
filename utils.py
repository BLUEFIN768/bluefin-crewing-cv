import fitz  # PyMuPDF
from docx import Document

def extract_text_from_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    return text

def extract_text_from_docx(file):
    doc = Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

def safe_extract(uploaded_file):
    extension = uploaded_file.name.split('.')[-1].lower()
    if extension == 'pdf':
        return extract_text_from_pdf(uploaded_file)
    elif extension == 'docx':
        return extract_text_from_docx(uploaded_file)
    else:
        raise Exception("Desteklenmeyen dosya formatı. Lütfen PDF veya DOCX yükleyin.")
