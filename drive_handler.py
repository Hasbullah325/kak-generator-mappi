import io
import os
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2 import service_account

# =====================================================================
# MASUKKAN ID FOLDER GOOGLE DRIVE ANDA DI SINI
# =====================================================================
FOLDER_ID_HASIL_KAK = "1CcnGThVTGWnvXudCvcspq-P0D9JzBEkq"
FOLDER_ID_POK = "1sAWdSwUHMkUchENehjlXL3s6fPhi45BP"
FOLDER_ID_MATRIKS_PERAN = "1DOtQYygFyr4ecBjYl2K-PAAoAF_-GbnO"
# =====================================================================

def get_drive_service(credentials_path="service_account.json"):
    if not os.path.exists(credentials_path):
        return None, "File service_account.json tidak ditemukan di folder aplikasi."
    try:
        scopes = ['https://www.googleapis.com/auth/drive']
        creds = service_account.Credentials.from_service_account_file(
            credentials_path, scopes=scopes
        )
        service = build('drive', 'v3', credentials=creds)
        return service, "OK"
    except Exception as e:
        return None, f"Gagal otentikasi Google Drive: {e}"

def download_latest_spreadsheet(folder_id):
    """Membaca file Excel (.xlsx) ATAU Google Sheets dari folder Google Drive."""
    service, status = get_drive_service()
    if not service:
        return None, status

    if "MASUKKAN_FOLDER_ID" in folder_id:
        return None, "ID Folder belum diganti dengan ID asli dari Google Drive."

    try:
        query = f"'{folder_id}' in parents and trashed=false"
        results = service.files().list(
            q=query, 
            orderBy="modifiedTime desc", 
            pageSize=5, 
            fields="files(id, name, mimeType)"
        ).execute()
        files = results.get('files', [])

        if not files:
            return None, "Folder di Google Drive kosong (tidak ada file ditemukan)."

        target_file = files[0]
        file_id = target_file['id']
        mime_type = target_file['mimeType']

        fh = io.BytesIO()

        if mime_type == 'application/vnd.google-apps.spreadsheet':
            request = service.files().export_media(
                fileId=file_id, 
                mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            request = service.files().get_media(fileId=file_id)

        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        
        fh.seek(0)
        df = pd.read_excel(fh)
        return df, f"Berhasil membaca file: {target_file['name']}"

    except Exception as e:
        return None, f"Error saat mengunduh file dari Drive: {e}"

def get_matriks_peran_data():
    """Mengambil data Ketua Tim dan menormalkan nama kolom."""
    df, status = download_latest_spreadsheet(FOLDER_ID_MATRIKS_PERAN)
    
    if df is None and os.path.exists("matriks_peran.xlsx"):
        try:
            df = pd.read_excel("matriks_peran.xlsx")
            status = "Menggunakan file cadangan lokal: matriks_peran.xlsx"
        except Exception:
            pass

    if df is None:
        status_msg = f"Gagal membaca dari Drive ({status}). Menggunakan data contoh default."
        df = pd.DataFrame([
            {"Jabatan": "Ketua Tim Sensus Ekonomi 2026", "Nama": "Muhammad Faza, S.Tr.Stat.", "NIP": "199802162021041001"},
            {"Jabatan": "Ketua Tim Statistik Sosial", "Nama": "John Tri Merjaya, SST., M.Ec.Dev.", "NIP": "198406192007011002"}
        ])
        return df, status_msg

    cols_map = {}
    for col in df.columns:
        c_clean = str(col).strip().lower()
        if "jabatan" in c_clean or "peran" in c_clean:
            cols_map[col] = "Jabatan"
        elif "nama" in c_clean:
            cols_map[col] = "Nama"
        elif "nip" in c_clean:
            cols_map[col] = "NIP"
    df = df.rename(columns=cols_map)

    for required in ["Jabatan", "Nama", "NIP"]:
        if required not in df.columns:
            df[required] = "-"

    return df, status

import openpyxl

def get_cell_color_code(cell):
    """Mendeteksi warna background sel secara akurat."""
    if cell.fill and cell.fill.fill_type:
        fg = cell.fill.fgColor
        if fg:
            if fg.type == 'rgb' and fg.rgb:
                rgb_str = str(fg.rgb).upper()
                # Abaikan warna putih atau hitam default/transparan
                if rgb_str not in ['00000000', 'FFFFFFFF', '00FFFFFF']:
                    return rgb_str
            elif fg.type == 'theme':
                return f"THEME_{fg.theme}_{fg.tint}"
            elif fg.type == 'indexed' and fg.indexed not in [64, 65]:
                return f"INDEXED_{fg.indexed}"
            elif fg.value:
                val_str = str(fg.value).upper()
                if val_str not in ['00000000', 'FFFFFFFF']:
                    return val_str
    return "NO_COLOR"

def parse_pok_by_color(file_stream_or_path):
    """
    Membaca POK SAKTI berdasarkan urutan 7 warna:
    Warna 1: Program
    Warna 2: Kegiatan
    Warna 3: KRO
    Warna 4: RO
    Warna 5: Komponen
    Warna 6: Sub Komponen
    Warna 7: Akun
    Sisa/Tanpa Warna: Detil
    """
    wb = openpyxl.load_workbook(file_stream_or_path, data_only=True)
    sheet = wb.active

    # 1. Deteksi urutan 7 warna unik yang pertama kali muncul dari atas ke bawah
    detected_colors = []
    for row in sheet.iter_rows():
        # Ambil teks sel
        cell_texts = [str(c.value).strip() for c in row if c.value is not None and str(c.value).strip()]
        full_row_text = " ".join(cell_texts).upper()
        
        # Lewati baris judul kolom (Header table)
        if "KODE" in full_row_text and ("URAIAN" in full_row_text or "PAGU" in full_row_text):
            continue

        # Cek warna pada sel pertama yang berisi data
        row_color = "NO_COLOR"
        for cell in row:
            if cell.value is not None and str(cell.value).strip():
                c_color = get_cell_color_code(cell)
                if c_color != "NO_COLOR":
                    row_color = c_color
                    break
        
        if row_color != "NO_COLOR" and row_color not in detected_colors:
            detected_colors.append(row_color)
            if len(detected_colors) == 7:
                break  # Sudah menemukan 7 tingkat warna lengkap

    # 2. Baca seluruh baris dan bentuk pohon hierarki
    cur_prog = ""
    cur_keg = ""
    cur_kro = ""
    cur_ro = ""
    cur_komp = ""
    cur_subkomp = ""
    cur_akun = ""

    records = []

    for row in sheet.iter_rows():
        non_empty = [c for c in row if c.value is not None and str(c.value).strip()]
        if not non_empty:
            continue

        row_texts = [str(c.value).strip() for c in non_empty]
        row_text_joined = " - ".join(row_texts[:2]) if len(row_texts) >= 2 else row_texts[0]

        # Lewati header tabel
        if "KODE" in row_text_joined.upper() and ("URAIAN" in row_text_joined.upper() or "JUMLAH" in row_text_joined.upper()):
            continue

        # Ambil warna baris
        row_color = "NO_COLOR"
        for cell in non_empty:
            c_color = get_cell_color_code(cell)
            if c_color != "NO_COLOR":
                row_color = c_color
                break

        # Klasifikasikan berdasarkan urutan warna yang terdaftar
        if len(detected_colors) > 0 and row_color == detected_colors[0]:
            cur_prog = row_text_joined
            cur_keg, cur_kro, cur_ro, cur_komp, cur_subkomp, cur_akun = "", "", "", "", "", ""
        elif len(detected_colors) > 1 and row_color == detected_colors[1]:
            cur_keg = row_text_joined
            cur_kro, cur_ro, cur_komp, cur_subkomp, cur_akun = "", "", "", "", ""
        elif len(detected_colors) > 2 and row_color == detected_colors[2]:
            cur_kro = row_text_joined
            cur_ro, cur_komp, cur_subkomp, cur_akun = "", "", "", ""
        elif len(detected_colors) > 3 and row_color == detected_colors[3]:
            cur_ro = row_text_joined
            cur_komp, cur_subkomp, cur_akun = "", "", ""
        elif len(detected_colors) > 4 and row_color == detected_colors[4]:
            cur_komp = row_text_joined
            cur_subkomp, cur_akun = "", ""
        elif len(detected_colors) > 5 and row_color == detected_colors[5]:
            cur_subkomp = row_text_joined
            cur_akun = ""
        elif len(detected_colors) > 6 and row_color == detected_colors[6]:
            cur_akun = row_text_joined
        else:
            # Baris Detil (Warna ke-8 atau tanpa warna / putih)
            detil_text = row_text_joined
            if cur_subkomp:  # Hanya rekam jika sudah berada di bawah subkomponen
                records.append({
                    "program": cur_prog or "-",
                    "kegiatan": cur_keg or "-",
                    "kro": cur_kro or "-",
                    "ro": cur_ro or "-",
                    "komponen": cur_komp or "-",
                    "sub_komponen": cur_subkomp or "-",
                    "akun": cur_akun or "-",
                    "detil": detil_text
                })

    if not records:
        return None

    return pd.DataFrame(records)

def get_pok_data():
    """Mengambil POK dari Google Drive dan memproses hierarki warnanya."""
    service, _ = get_drive_service()
    
    # 1. Coba unduh dari Google Drive
    if service and "MASUKKAN_FOLDER_ID" not in FOLDER_ID_POK:
        try:
            query = f"'{FOLDER_ID_POK}' in parents and trashed=false"
            results = service.files().list(q=query, orderBy="modifiedTime desc", pageSize=1, fields="files(id, name)").execute()
            files = results.get('files', [])
            if files:
                file_id = files[0]['id']
                request = service.files().get_media(fileId=file_id)
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
                fh.seek(0)
                df = parse_pok_by_color(fh)
                if df is not None:
                    return df, f"Berhasil membaca POK berdasarkan warna ({files[0]['name']})"
        except Exception as e:
            print(f"Gagal baca Drive POK: {e}")

    # 2. Coba file lokal pok.xlsx jika ada
    if os.path.exists("pok.xlsx"):
        try:
            df = parse_pok_by_color("pok.xlsx")
            if df is not None:
                return df, "Berhasil membaca file lokal pok.xlsx berdasarkan kelompok warna."
        except Exception as e:
            print(f"Gagal baca lokal POK: {e}")

    # 3. Data Cadangan jika file belum diunggah
    df_fallback = pd.DataFrame([
        {
            "program": "PROGRAM PENYEDIAAN DAN PELAYANAN INFORMASI STATISTIK (GG)",
            "kegiatan": "PENYEDIAAN DAN PENGEMBANGAN STATISTIK DISTRIBUSI (2902)",
            "kro": "DATA DAN INFORMASI PUBLIK (2902.BMA)",
            "ro": "PUBLIKASI/LAPORAN SENSUS EKONOMI (2902.BMA.006)",
            "komponen": "PELAKSANAAN SE2026 (530)",
            "sub_komponen": "Pelatihan Petugas Lapangan (A)",
            "akun": "524114. BELANJA PERJALANAN DINAS PAKET MEETING DALAM KOTA",
            "detil": "000488. Paket Meeting Fullday pelatihan petugas pendataan"
        },
        {
            "program": "PROGRAM PENYEDIAAN DAN PELAYANAN INFORMASI STATISTIK (GG)",
            "kegiatan": "PENYEDIAAN DAN PENGEMBANGAN STATISTIK DISTRIBUSI (2902)",
            "kro": "DATA DAN INFORMASI PUBLIK (2902.BMA)",
            "ro": "PUBLIKASI/LAPORAN SENSUS EKONOMI (2902.BMA.006)",
            "komponen": "PELAKSANAAN SE2026 (530)",
            "sub_komponen": "Pelatihan Petugas Lapangan (A)",
            "akun": "524114. BELANJA PERJALANAN DINAS PAKET MEETING DALAM KOTA",
            "detil": "000489. Perjalanan pelatihan petugas pendataan"
        }
    ])
    return df_fallback, "Menggunakan data default (Unggah file POK SAKTI untuk membaca warna otomatis)."

def upload_kak_to_drive(doc_bytes_io, file_name, folder_id=FOLDER_ID_HASIL_KAK):
    service, _ = get_drive_service()
    if not service:
        return None
    file_metadata = {'name': file_name, 'parents': [folder_id]}
    doc_bytes_io.seek(0)
    media = MediaIoBaseUpload(
        doc_bytes_io,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        resumable=True
    )
    uploaded = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
    return uploaded.get('webViewLink')