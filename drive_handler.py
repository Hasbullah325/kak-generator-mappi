import io
import os
import re
import requests
import openpyxl
import pandas as pd

# ==============================================================================
# MASUKKAN ID FILE GOOGLE DRIVE ANDA DI SINI
# (Didapat dari link 'Bagikan -> Siapa saja yang memiliki link')
# ==============================================================================
FILE_ID_POK = "1Y4-erCGAjUVfo4zI0dAt1JG-HYTHJgn9"
FILE_ID_MATRIKS = "1Tr0Fcaq-QwtZFSMG0ZUOjA_ddO7ciPIu"
# ==============================================================================

def download_file_from_drive(file_id):
    """Mengunduh file Excel dari Google Drive via direct download link."""
    if not file_id or "MASUKKAN_ID" in file_id:
        return None
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    session = requests.Session()
    response = session.get(url, timeout=30)
    
    # Tangani konfirmasi file besar Google Drive jika ada
    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            url = f"https://drive.google.com/uc?export=download&confirm={value}&id={file_id}"
            response = session.get(url, timeout=30)
            break
            
    if response.status_code == 200:
        return io.BytesIO(response.content)
    return None

def get_cell_color_code(cell):
    """Mendeteksi warna background sel secara akurat."""
    if cell.fill and cell.fill.fill_type:
        fg = cell.fill.fgColor
        if fg:
            if fg.type == 'rgb' and fg.rgb:
                rgb_str = str(fg.rgb).upper()
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
    """Membaca hierarki POK berdasarkan 7 urutan warna."""
    wb = openpyxl.load_workbook(file_stream_or_path, data_only=True)
    sheet = wb.active

    # 1. Kumpulkan urutan warna unik dari atas ke bawah
    detected_colors = []
    for row in sheet.iter_rows():
        cell_texts = [str(c.value).strip() for c in row if c.value is not None and str(c.value).strip()]
        full_row_text = " ".join(cell_texts).upper()
        if "KODE" in full_row_text and ("URAIAN" in full_row_text or "PAGU" in full_row_text):
            continue

        row_color = "NO_COLOR"
        for cell in row:
            if cell.value is not None and str(cell.value).strip():
                c_color = get_cell_color_code(cell)
                if c_color != "NO_COLOR":
                    row_color = c_color
                    break
        
        if row_color != "NO_COLOR" and row_color not in detected_colors:
            detected_colors.append(row_color)
            if len(detected_colors) >= 8:
                break

    # 2. Baca seluruh baris dan bentuk pohon hierarki
    num_colors = len(detected_colors)
    records = []

    cur_prog = ""
    cur_keg = ""
    cur_kro = ""
    cur_ro = ""
    cur_komp = ""
    cur_subkomp = ""
    cur_akun = ""

    for row in sheet.iter_rows():
        non_empty = [c for c in row if c.value is not None and str(c.value).strip()]
        if not non_empty:
            continue

        row_texts = [str(c.value).strip() for c in non_empty]
        row_text_joined = " - ".join(row_texts[:2]) if len(row_texts) >= 2 else row_texts[0]

        if "KODE" in row_text_joined.upper() and ("URAIAN" in row_text_joined.upper() or "JUMLAH" in row_text_joined.upper()):
            continue

        row_color = "NO_COLOR"
        for cell in non_empty:
            c_color = get_cell_color_code(cell)
            if c_color != "NO_COLOR":
                row_color = c_color
                break

        if num_colors > 0 and row_color == detected_colors[0]:
            cur_prog = row_text_joined
            cur_keg, cur_kro, cur_ro, cur_komp, cur_subkomp, cur_akun = "", "", "", "", "", ""
        elif num_colors > 1 and row_color == detected_colors[1]:
            cur_keg = row_text_joined
            cur_kro, cur_ro, cur_komp, cur_subkomp, cur_akun = "", "", "", "", ""
        elif num_colors > 2 and row_color == detected_colors[2]:
            cur_kro = row_text_joined
            cur_ro, cur_komp, cur_subkomp, cur_akun = "", "", "", ""
        elif num_colors > 3 and row_color == detected_colors[3]:
            cur_ro = row_text_joined
            cur_komp, cur_subkomp, cur_akun = "", "", ""
        elif num_colors > 4 and row_color == detected_colors[4]:
            cur_komp = row_text_joined
            cur_subkomp, cur_akun = "", ""
        elif num_colors > 5 and row_color == detected_colors[5]:
            cur_subkomp = row_text_joined
            cur_akun = ""
        elif num_colors > 6 and row_color == detected_colors[6]:
            cur_akun = row_text_joined
        else:
            detil_text = row_text_joined
            if cur_subkomp or cur_prog:
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
    """Mengambil POK otomatis dari Google Drive."""
    stream = download_file_from_drive(FILE_ID_POK)
    if stream:
        try:
            df = parse_pok_by_color(stream)
            if df is not None and not df.empty:
                return df, f"Berhasil sinkronisasi {len(df)} baris data POK dari Google Drive."
        except Exception as e:
            return None, f"Gagal membaca format warna POK dari Drive: {e}"

    # Jika file lokal ada di repo
    for fname in ["pok.xlsx", "POK.xlsx"]:
        if os.path.exists(fname):
            try:
                df = parse_pok_by_color(fname)
                if df is not None:
                    return df, f"Membaca file lokal {fname}."
            except Exception:
                pass

    # Data Default jika ID belum diisi
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
        }
    ])
    return df_fallback, "Google Drive belum tersambung (masukkan FILE_ID_POK)."

def get_matriks_peran_data():
    """Mengambil Matriks Peran otomatis dari Google Drive."""
    stream = download_file_from_drive(FILE_ID_MATRIKS)
    if stream:
        try:
            df = pd.read_excel(stream)
            cols_map = {}
            for col in df.columns:
                c = str(col).strip().lower()
                if "jabatan" in c or "peran" in c: cols_map[col] = "Jabatan"
                elif "nama" in c: cols_map[col] = "Nama"
                elif "nip" in c: cols_map[col] = "NIP"
            df = df.rename(columns=cols_map)
            return df, f"Berhasil sinkronisasi Matriks Peran ({len(df)} pejabat) dari Google Drive."
        except Exception as e:
            return None, f"Error Matriks: {e}"

    for fname in ["matriks_peran.xlsx", "Matriks_Peran.xlsx"]:
        if os.path.exists(fname):
            df = pd.read_excel(fname)
            return df, f"Membaca file lokal {fname}."

    df_default = pd.DataFrame([
        {"Jabatan": "Ketua Tim Sensus Ekonomi 2026", "Nama": "Muhammad Faza, S.Tr.Stat.", "NIP": "199802162021041001"},
        {"Jabatan": "Ketua Tim Statistik Sosial", "Nama": "John Tri Merjaya, SST., M.Ec.Dev.", "NIP": "198406192007011002"}
    ])
    return df_default, "Google Drive belum tersambung (masukkan FILE_ID_MATRIKS)."

def upload_kak_to_drive(doc_bytes_io, file_name):
    # Fitur download Word lokal tetap aktif 100%
    return None
