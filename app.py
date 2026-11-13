import streamlit as st
import io
import os
import pandas as pd
from docx_builder import generate_kak_docx
from ai_extractor import extract_kak_narrative_from_docs
from drive_handler import get_pok_data, get_matriks_peran_data, upload_kak_to_drive

st.set_page_config(page_title="Generator KAK - BPS Mappi", layout="wide")
st.title("🏛️ Aplikasi Generator KAK BPS Kabupaten Mappi")

# --- KONFIGURASI SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Konfigurasi Sistem")
    api_key_input = st.text_input("Gemini API Key:", value=os.getenv("GEMINI_API_KEY", ""), type="password")
    if api_key_input:
        os.environ["GEMINI_API_KEY"] = api_key_input
    st.caption("Kunci API diperlukan untuk menjalankan fitur AI perangkum dokumen Juknis.")
    st.markdown("---")
    st.caption("BPS Kabupaten Mappi © 2026")

# --- AMBIL MASTER DATA (POK & MATRIKS PERAN) ---
df_pok, status_pok = get_pok_data()
df_matriks, status_matriks = get_matriks_peran_data()

# --- INISIALISASI VARIABEL DEFAULT (MENCEGAH NameError) ---
program = "-"
kegiatan = "-"
kro = "-"
ro = "-"
komponen = "-"
sub_komponen = "-"
akun = "-"
detil = "-"
teks_akun_lengkap = ""
teks_detil_lengkap = ""

# Inisialisasi State Tabel Anggaran
if "tabel_rab_df" not in st.session_state:
    st.session_state.tabel_rab_df = pd.DataFrame([
        {
            "Nama Pembiayaan": "Paket Meeting Fullday pelatihan petugas pendataan",
            "Volume": 1.0,
            "Ukuran": "PAKET",
            "Harga Satuan (Rp)": 246000000.0,
            "Total Biaya (Rp)": 246000000.0
        },
        {
            "Nama Pembiayaan": "Perjalanan pelatihan petugas pendataan",
            "Volume": 1.0,
            "Ukuran": "PAKET",
            "Harga Satuan (Rp)": 273950000.0,
            "Total Biaya (Rp)": 273950000.0
        }
    ])

# Inisialisasi State Narasi AI
if "latar" not in st.session_state:
    st.session_state.latar = ""
if "tujuan" not in st.session_state:
    st.session_state.tujuan = ""
if "metodologi" not in st.session_state:
    st.session_state.metodologi = ""

# --- PEMBUATAN TAB NAVIGASI ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1. Header & BAS", 
    "2. Narasi AI & Juknis", 
    "3. Teknis, Waktu & Output", 
    "4. Anggaran (RAB)", 
    "5. Penandatangan & Cetak"
])

# ==============================================================================
# TAB 1: HEADER & BAS BERJENJANG (CASCADING DROPDOWN)
# ==============================================================================
with tab1:
    st.subheader("Judul Dokumen")
    nama_kegiatan = st.text_input("Nama Kegiatan (Otomatis Huruf Kapital):", value="Pelatihan Petugas Pendataan Lapangan Sensus Ekonomi 2026")
    
    st.markdown("---")
    st.subheader("Klasifikasi Bagan Akun Standar (BAS)")
    if "Berhasil" in status_pok:
        st.success(f"✅ {status_pok}")
    else:
        st.info(f"ℹ️ Status POK: {status_pok}")

    col_bas1, col_bas2 = st.columns(2)
    
    with col_bas1:
        # Level 1: Program
        list_prog = df_pok["program"].dropna().astype(str).unique().tolist() if "program" in df_pok.columns else ["-"]
        program = st.selectbox("1. Program:", list_prog)
        df_lvl1 = df_pok[df_pok["program"].astype(str) == program] if "program" in df_pok.columns else df_pok

        # Level 2: Kegiatan
        list_keg = df_lvl1["kegiatan"].dropna().astype(str).unique().tolist() if "kegiatan" in df_lvl1.columns else ["-"]
        kegiatan = st.selectbox("2. Kegiatan:", list_keg)
        df_lvl2 = df_lvl1[df_lvl1["kegiatan"].astype(str) == kegiatan] if "kegiatan" in df_lvl1.columns else df_lvl1

        # Level 3: Klasifikasi RO (KRO)
        list_kro = df_lvl2["kro"].dropna().astype(str).unique().tolist() if "kro" in df_lvl2.columns else ["-"]
        kro = st.selectbox("3. Klasifikasi RO (KRO):", list_kro)
        df_lvl3 = df_lvl2[df_lvl2["kro"].astype(str) == kro] if "kro" in df_lvl2.columns else df_lvl2

    with col_bas2:
        # Level 4: Rincian Output (RO)
        list_ro = df_lvl3["ro"].dropna().astype(str).unique().tolist() if "ro" in df_lvl3.columns else ["-"]
        ro = st.selectbox("4. Rincian Output (RO):", list_ro)
        df_lvl4 = df_lvl3[df_lvl3["ro"].astype(str) == ro] if "ro" in df_lvl3.columns else df_lvl3

        # Level 5: Komponen
        list_komp = df_lvl4["komponen"].dropna().astype(str).unique().tolist() if "komponen" in df_lvl4.columns else ["-"]
        komponen = st.selectbox("5. Komponen:", list_komp)
        df_lvl5 = df_lvl4[df_lvl4["komponen"].astype(str) == komponen] if "komponen" in df_lvl4.columns else df_lvl4

        # Level 6: Sub Komponen
        list_subkomp = df_lvl5["sub_komponen"].dropna().astype(str).unique().tolist() if "sub_komponen" in df_lvl5.columns else ["-"]
        sub_komponen = st.selectbox("6. Sub Komponen:", list_subkomp)

    # Filter akhir data POK sesuai Sub Komponen terpilih
    if "sub_komponen" in df_lvl5.columns:
        pok_final = df_lvl5[df_lvl5["sub_komponen"].astype(str) == sub_komponen]
    else:
        pok_final = df_lvl5

    # Ambil SEMUA Akun dan SEMUA Detil di bawah Sub Komponen ini
    semua_akun = sorted([str(x) for x in pok_final["akun"].dropna().unique().tolist() if str(x) != "-"]) if "akun" in pok_final.columns else []
    semua_detil = sorted([str(x) for x in pok_final["detil"].dropna().unique().tolist() if str(x) != "-"]) if "detil" in pok_final.columns else []
    
    teks_akun_lengkap = "\n".join(semua_akun) if semua_akun else "-"
    teks_detil_lengkap = "\n".join(semua_detil) if semua_detil else "-"

    st.markdown("##### Rincian Beban Anggaran di Bawah Sub Komponen Terpilih:")
    col_ak1, col_ak2 = st.columns(2)
    with col_ak1:
        akun = st.text_area("Semua Akun Belanja Terkait:", value=teks_akun_lengkap, height=100)
    with col_ak2:
        detil = st.text_area("Semua Detil / Uraian Kegiatan Terkait:", value=teks_detil_lengkap, height=100)

    st.markdown("---")
    st.subheader("Organisasi Pengadaan (PPK)")
    ppk_option = st.selectbox("Pilih Profil PPK BPS Mappi:", [
        "John Tri Merjaya, SST., M.Ec.Dev. (198406192007011002)",
        "Input Manual..."
    ])
    if "Manual" in ppk_option:
        ppk_nama = st.text_input("Nama PPK:")
        ppk_nip = st.text_input("NIP PPK:")
    else:
        ppk_nama = "John Tri Merjaya, SST., M.Ec.Dev."
        ppk_nip = "198406192007011002"

# ==============================================================================
# TAB 2: NARASI AI (JUKNIS)
# ==============================================================================
with tab2:
    st.subheader("Ekstraksi Dokumen Referensi (Juknis/TOR)")
    uploaded_files = st.file_uploader("Unggah Juknis Kegiatan (PDF / DOCX / TXT):", accept_multiple_files=True, type=["pdf", "docx", "txt"])
    
    if st.button("✨ Ekstrak & Rangkum dengan AI"):
        if not uploaded_files:
            st.warning("Silakan unggah minimal 1 file Juknis/TOR terlebih dahulu.")
        elif not os.getenv("GEMINI_API_KEY"):
            st.error("Gemini API Key belum dimasukkan. Silakan isi di sidebar kiri terlebih dahulu.")
        else:
            with st.spinner("AI sedang membaca dan mengekstrak narasi KAK..."):
                try:
                    res = extract_kak_narrative_from_docs(uploaded_files)
                    st.session_state.latar = res.get("latar_belakang", "")
                    st.session_state.tujuan = res.get("tujuan", "")
                    st.session_state.metodologi = res.get("metodologi", "")
                    st.success("✅ Ekstraksi dokumen berhasil! Anda dapat membaca dan mengedit teks di bawah.")
                except Exception as e:
                    st.error(f"Gagal melakukan ekstraksi AI: {str(e)}")

    latar_belakang = st.text_area("A. Latar Belakang Kegiatan:", value=st.session_state.latar, height=150)
    tujuan = st.text_area("B. Maksud dan Tujuan Kegiatan:", value=st.session_state.tujuan, height=100)
    metodologi = st.text_area("C. Metodologi dan Pelaksanaan:", value=st.session_state.metodologi, height=120)

# ==============================================================================
# TAB 3: WAKTU, LOKASI, OUTPUT & SPESIFIKASI
# ==============================================================================
with tab3:
    c1, c2 = st.columns(2)
    with c1:
        waktu_pelaksanaan = st.text_input("Waktu/Tanggal Kegiatan:", value="4 s.d 10 Juni 2026")
        lokasi_pelaksanaan = st.text_input("Tempat/Lokasi Kegiatan:", value="Hotel Grand Avista, Jl. Irian Kepi KM 2")
    with c2:
        keluaran = st.text_area("D. Keluaran / Output Kegiatan:", value="Petugas pendataan lapangan yang memahami SOP, konsep metodologi, dan tata cara pengisian kuesioner SE2026.")

    st.markdown("---")
    st.subheader("E. Spesifikasi Teknis")
    spesifikasi_tambahan = st.text_area("Tambah Poin Spesifikasi Teknis Khusus (1 poin per baris):", value="")

    st.subheader("H. Penutup")
    penutup_tambahan = st.text_area("Tambah Poin Penutup Khusus (1 poin per baris):", value="")

# ==============================================================================
# TAB 4: SUMBER DANA & BIAYA (RAB)
# ==============================================================================
with tab4:
    st.subheader("G. Sumber Dana dan Perkiraan Biaya")
    sumber_dipa = st.text_input("Poin a (Sumber Dana DIPA):", value="DIPA Satker BPS Kabupaten Mappi Tahun Anggaran 2026 Nomor SP DIPA-054.01.2.668760/2026 Tanggal 2 Desember 2025")
    
    beban_raw = st.text_area("Poin b (Akun yang dibebankan pada kegiatan ini):", value=teks_akun_lengkap)
    ket_biaya = st.text_area("Poin c (Keterangan Pembebanan Lain - Kosongkan jika tidak ada):", value="")

    st.markdown("---")
    st.subheader("Tabel Rincian Anggaran Biaya")
    st.caption("💡 Kolom **Total Biaya (Rp)** otomatis terkunci dan terhitung dari perkalian (Volume × Harga Satuan).")

    # Kalkulasi nilai awal
    df_calc = st.session_state.tabel_rab_df.copy()
    df_calc["Volume"] = pd.to_numeric(df_calc["Volume"], errors='coerce').fillna(0)
    df_calc["Harga Satuan (Rp)"] = pd.to_numeric(df_calc["Harga Satuan (Rp)"], errors='coerce').fillna(0)
    df_calc["Total Biaya (Rp)"] = df_calc["Volume"] * df_calc["Harga Satuan (Rp)"]

    edited_df = st.data_editor(
        df_calc,
        num_rows="dynamic",
        use_container_width=True,
        disabled=["Total Biaya (Rp)"],
        column_config={
            "Nama Pembiayaan": st.column_config.TextColumn("Nama Pembiayaan", required=True),
            "Volume": st.column_config.NumberColumn("Volume", min_value=0.0, step=1.0, format="%.2f"),
            "Ukuran": st.column_config.TextColumn("Ukuran / Satuan", help="Contoh: PAKET, ORANG, HARI"),
            "Harga Satuan (Rp)": st.column_config.NumberColumn("Harga Satuan (Rp)", min_value=0.0, format="Rp %d"),
            "Total Biaya (Rp)": st.column_config.NumberColumn("Total Biaya (Rp)", format="Rp %d", help="Otomatis: Volume x Harga Satuan")
        }
    )

    # Sinkronkan kembali perubahan user
    edited_df["Total Biaya (Rp)"] = (
        pd.to_numeric(edited_df["Volume"], errors='coerce').fillna(0) *
        pd.to_numeric(edited_df["Harga Satuan (Rp)"], errors='coerce').fillna(0)
    )
    st.session_state.tabel_rab_df = edited_df

    total_anggaran = edited_df["Total Biaya (Rp)"].sum()
    st.metric("Total Perkiraan Biaya Kegiatan:", f"Rp {total_anggaran:,.0f}")

# ==============================================================================
# TAB 5: PENANDATANGAN & GENERATE DOKUMEN
# ==============================================================================
with tab5:
    st.subheader("Penandatanganan Dokumen")
    if "Berhasil" in status_matriks:
        st.success(f"✅ {status_matriks}")
    else:
        st.info(f"ℹ️ Status Matriks Peran: {status_matriks}")

    col_tgl, col_tim = st.columns(2)
    with col_tgl:
        tanggal_penandatangan = st.text_input("Tanggal Penandatanganan KAK (Input Manual):", value="18 Mei 2026")
    with col_tim:
        daftar_ketua = []
        for _, r in df_matriks.iterrows():
            daftar_ketua.append(f"{r['Jabatan']} - {r['Nama']} ({r['NIP']})")
        if not daftar_ketua:
            daftar_ketua = ["Ketua Tim - Nama Ketua Tim (19xxxxxxxxxxxxxxx)"]
            
        ketua_option = st.selectbox("Pilih Ketua Tim (Berdasarkan Matriks Peran):", daftar_ketua)
        
        try:
            jabatan_kt = ketua_option.split(" - ")[0]
            nama_kt = ketua_option.split(" - ")[1].split(" (")[0]
            nip_kt = ketua_option.split("(")[1].replace(")", "")
        except Exception:
            jabatan_kt, nama_kt, nip_kt = "Ketua Tim", "-", "-"

    st.markdown("---")
    st.subheader("Lembar Lampiran")
    lampiran_input = st.text_area("Daftar Judul Lampiran (Otomatis dibuatkan halaman kosong di bagian akhir, 1 per baris):", 
                                  value="Susunan Acara Pelatihan\nDaftar Calon Peserta dan Panitia")

    st.markdown("---")
    if st.button("🚀 Generate Dokumen KAK (.docx)", type="primary"):
        data_payload = {
            "nama_kegiatan": nama_kegiatan,
            "program": program,
            "kegiatan": kegiatan,
            "kro": kro,
            "ro": ro,
            "komponen": komponen,
            "sub_komponen": sub_komponen,
            "akun_belanja": akun if 'akun' in locals() and akun else teks_akun_lengkap,
            "detil": detil if 'detil' in locals() and detil else teks_detil_lengkap,
            "ppk_nama": ppk_nama,
            "ppk_nip": ppk_nip,
            "latar_belakang": latar_belakang,
            "tujuan": tujuan,
            "metodologi": metodologi,
            "waktu_pelaksanaan": waktu_pelaksanaan,
            "lokasi_pelaksanaan": lokasi_pelaksanaan,
            "keluaran": keluaran,
            "spesifikasi_tambahan": [x for x in spesifikasi_tambahan.split("\n") if x.strip()],
            "penutup_tambahan": [x for x in penutup_tambahan.split("\n") if x.strip()],
            "sumber_dana_dipa": sumber_dipa,
            "pembiayaan_beban": [x for x in beban_raw.split("\n") if x.strip()],
            "keterangan_tambahan_biaya": ket_biaya,
            "tabel_biaya": edited_df.to_dict("records"),
            "tanggal_penandatangan": tanggal_penandatangan,
            "jabatan_ketua_tim": jabatan_kt,
            "nama_ketua_tim": nama_kt,
            "nip_ketua_tim": nip_kt,
            "daftar_lampiran": [x for x in lampiran_input.split("\n") if x.strip()]
        }

        with st.spinner("Sedang menyusun file Word (.docx)..."):
            doc = generate_kak_docx(data_payload)
            
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            
            file_name = f"KAK_{nama_kegiatan.replace(' ', '_').upper()[:35]}.docx"
            
            # Coba upload ke Google Drive
            drive_link = None
            try:
                drive_link = upload_kak_to_drive(doc_io, file_name)
            except Exception as e:
                st.warning(f"Catatan: File belum terunggah ke Google Drive ({e}).")

            doc_io.seek(0)
            st.success("✅ Dokumen KAK (.docx) berhasil dibuat!")
            if drive_link:
                st.info(f"📁 **Tersimpan di Google Drive:** [Buka Dokumen di Google Drive]({drive_link})")

            # Tombol Download Langsung
            st.download_button(
                label="📥 Unduh Dokumen Word Sekarang (.docx)",
                data=doc_io,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )