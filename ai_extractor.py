import os
import io
import json
from pypdf import PdfReader
import docx
from google import genai
from google.genai import types

def read_file_content(uploaded_file):
    """Mengekstrak teks murni dari PDF, DOCX, atau TXT."""
    file_bytes = uploaded_file.getvalue()
    filename = uploaded_file.name.lower()
    text = ""
    
    try:
        if filename.endswith(".pdf"):
            pdf_reader = PdfReader(io.BytesIO(file_bytes))
            for page in pdf_reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        elif filename.endswith(".docx"):
            doc = docx.Document(io.BytesIO(file_bytes))
            text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        else:
            text = file_bytes.decode("utf-8", errors="ignore")
    except Exception as e:
        print(f"Error membaca {uploaded_file.name}: {e}")
        
    return text

def extract_kak_narrative_from_docs(uploaded_files, api_key=None):
    active_key = api_key or os.getenv("GEMINI_API_KEY")
    if not active_key:
        raise ValueError("API Key Gemini belum disetel! Masukkan API Key di sidebar terlebih dahulu.")

    client = genai.Client(api_key=active_key)
    
    combined_text = ""
    for file in uploaded_files:
        content = read_file_content(file)
        combined_text += f"\n--- ISI DOKUMEN: {file.name} ---\n{content}\n"

    if len(combined_text.strip()) < 50:
        raise ValueError("Dokumen yang diunggah kosong atau teksnya tidak terbaca.")

    prompt = f"""
    Kamu adalah perencana kegiatan profesional di Badan Pusat Statistik (BPS).
    Berdasarkan dokumen pedoman teknis/Juknis berikut, rangkum dan susunlah 4 narasi untuk Kerangka Acuan Kerja (KAK):
    1. Latar Belakang (Jelaskan urgensi, landasan hukum/UU statistik, dan konteks kegiatan)
    2. Maksud dan Tujuan Kegiatan
    3. Metodologi dan Alur Pelaksanaan Kegiatan
    4. Jadwal/Tahapan Kegiatan

    Kembalikan HANYA format JSON valid seperti ini:
    {{
        "latar_belakang": "...",
        "tujuan": "...",
        "metodologi": "...",
        "jadwal": "..."
    }}

    Dokumen Sumber:
    {combined_text[:30000]}
    """

    # Daftar model prioritas sesuai anjuran Google API
    model_candidates = ["gemini-3.8-flash", "gemini-2.5-flash", "gemini-2.0-flash"]
    last_error = None

    for m in model_candidates:
        try:
            response = client.models.generate_content(
                model=m,
                contents=prompt,
                config=types.GenerateContentConfig(response_mime_type="application/json")
            )
            return json.loads(response.text)
        except Exception as e:
            last_error = e
            continue

    raise last_error