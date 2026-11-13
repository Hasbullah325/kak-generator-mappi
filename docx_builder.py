import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    """Memberi warna latar (shading) pada sel tabel."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def format_rupiah(nilai):
    """Format angka menjadi teks nominal rupiah dengan pemisah ribuan titik."""
    try:
        val = float(nilai)
        return f"{val:,.0f}".replace(",", ".")
    except (ValueError, TypeError):
        return "0"

def generate_kak_docx(data):
    """Menyusun seluruh dokumen Kerangka Acuan Kerja (KAK) format .docx resmi."""
    doc = docx.Document()

    # --- 1. SETTING MARGIN STANDAR (1 Inci / 2.54 cm) ---
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # --- 2. JUDUL DOKUMEN ---
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(12)

    r_title1 = p_title.add_run("KERANGKA ACUAN KERJA (KAK)\n")
    r_title1.bold = True
    r_title1.font.size = Pt(12)
    r_title1.font.name = "Arial"
    
    # Judul kegiatan UPPERCASE (Huruf Kapital)
    nama_kegiatan = str(data.get("nama_kegiatan", "")).strip().upper()
    r_title2 = p_title.add_run(f"{nama_kegiatan}\n")
    r_title2.bold = True
    r_title2.font.size = Pt(12)
    r_title2.font.name = "Arial"
    
    r_title3 = p_title.add_run("BADAN PUSAT STATISTIK KABUPATEN MAPPI")
    r_title3.bold = True
    r_title3.font.size = Pt(12)
    r_title3.font.name = "Arial"

    # --- 3. TABEL INFORMASI BAGAN AKUN STANDAR (BAS) ---
    bas_table = doc.add_table(rows=0, cols=3)
    bas_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    bas_table.autofit = False

    bas_rows = [
        ("KEMENTERIAN NEGARA/LEMBAGA", ":", "BADAN PUSAT STATISTIK (054)"),
        ("UNIT ESELON I", ":", "BADAN PUSAT STATISTIK (01)"),
        ("UNIT ESELON II / SATKER", ":", f"BPS PROVINSI PAPUA / BPS KABUPATEN MAPPI ({data.get('kode_satker', '668760')})"),
        ("PROGRAM", ":", str(data.get("program", "-"))),
        ("KEGIATAN", ":", str(data.get("kegiatan", "-"))),
        ("KLASIFIKASI RO", ":", str(data.get("kro", "-"))),
        ("RINCIAN OUTPUT", ":", str(data.get("ro", "-"))),
        ("KOMPONEN", ":", str(data.get("komponen", "-"))),
        ("SUB KOMPONEN", ":", str(data.get("sub_komponen", "-"))),
        ("AKUN BELANJA", ":", str(data.get("akun_belanja", "-"))),
        ("DETIL", ":", str(data.get("detil", "-")))
    ]

    for label, sep, val in bas_rows:
        row = bas_table.add_row().cells
        row[0].width = Inches(2.4)
        row[1].width = Inches(0.2)
        row[2].width = Inches(3.9)

        row[0].text = label
        row[1].text = sep
        row[2].text = val

        for cell in row:
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                for run in p.runs:
                    run.font.size = Pt(9)
                    run.font.name = "Arial"

    # --- MEMISAHKAN HALAMAN BAS DAN LATAR BELAKANG ---
    doc.add_page_break()

    # --- 4. LATAR BELAKANG DAN TUJUAN (HALAMAN 2) ---
    doc.add_heading("A. LATAR BELAKANG", level=2)
    doc.add_paragraph(str(data.get("latar_belakang", "-")))

    doc.add_heading("B. MAKSUD DAN TUJUAN", level=2)
    doc.add_paragraph(str(data.get("tujuan", "-")))

    doc.add_heading("C. METODOLOGI DAN PELAKSANAAN", level=2)
    doc.add_paragraph(str(data.get("metodologi", "-")))
    
    p_waktu = doc.add_paragraph()
    p_waktu.add_run("Waktu Pelaksanaan: ").bold = True
    p_waktu.add_run(str(data.get("waktu_pelaksanaan", "-")))
    
    p_lokasi = doc.add_paragraph()
    p_lokasi.add_run("Tempat/Lokasi: ").bold = True
    p_lokasi.add_run(str(data.get("lokasi_pelaksanaan", "-")))

    doc.add_heading("D. KELUARAN / OUTPUT", level=2)
    doc.add_paragraph(str(data.get("keluaran", "-")))

    # --- 5. SPESIFIKASI TEKNIS ---
    doc.add_heading("E. SPESIFIKASI TEKNIS", level=2)
    doc.add_paragraph(str(data.get("spesifikasi_teknis_default", "Spesifikasi teknis mengacu pada standar operasional prosedur dan panduan teknis yang ditetapkan oleh Badan Pusat Statistik.")))
    for item in data.get("spesifikasi_tambahan", []):
        if str(item).strip():
            doc.add_paragraph(str(item).strip(), style='List Bullet')

    # --- 6. NAMA ORGANISASI PENGADAAN ---
    doc.add_heading("F. NAMA ORGANISASI PENGADAAN", level=2)
    doc.add_paragraph("a. K/L: BADAN PUSAT STATISTIK (054)")
    doc.add_paragraph(f"b. Satuan Kerja: BPS KABUPATEN MAPPI ({data.get('kode_satker', '668760')})")
    doc.add_paragraph(f"c. Pejabat Pembuat Komitmen: {data.get('ppk_nama', '-')} (NIP. {data.get('ppk_nip', '-')})")

    # --- 7. SUMBER DANA DAN PERKIRAAN BIAYA ---
    doc.add_heading("G. SUMBER DANA DAN PERKIRAAN BIAYA", level=2)
    doc.add_paragraph(f"a. Sumber Dana: {data.get('sumber_dana_dipa', '-')}")

    doc.add_paragraph("b. Pembiayaan yang dibebankan pada kegiatan ini:")
    for beban in data.get("pembiayaan_beban", []):
        if str(beban).strip():
            doc.add_paragraph(str(beban).strip(), style='List Bullet')

    ket_tambahan = str(data.get("keterangan_tambahan_biaya", "")).strip()
    if ket_tambahan:
        doc.add_paragraph(f"c. Keterangan Pembebanan Lain: {ket_tambahan}")

    # --- 8. TABEL RINCIAN ANGGARAN BIAYA (RAB) ---
    doc.add_paragraph("Total perkiraan biaya yang diperlukan pada pelaksanaan kegiatan ini adalah sebagai berikut:")
    
    rab_table = doc.add_table(rows=1, cols=5)
    rab_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    rab_table.autofit = False

    headers = ["Nama Pembiayaan", "Volume", "Ukuran", "Harga Satuan (Rp)", "Total Biaya (Rp)"]
    col_widths = [Inches(2.7), Inches(0.8), Inches(0.9), Inches(1.3), Inches(1.3)]
    
    hdr_cells = rab_table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        hdr_cells[i].width = col_widths[i]
        set_cell_background(hdr_cells[i], "D9D9D9")
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        p.runs[0].font.size = Pt(9)
        p.runs[0].font.name = "Arial"

    total_anggaran_kegiatan = 0.0

    for item in data.get("tabel_biaya", []):
        row = rab_table.add_row().cells
        
        nama = str(item.get("Nama Pembiayaan") or item.get("nama") or "").strip()
        ukuran = str(item.get("Ukuran") or item.get("ukuran") or "").strip()
        
        try:
            vol = float(item.get("Volume") or item.get("volume") or 0.0)
        except (ValueError, TypeError):
            vol = 0.0
            
        try:
            harga = float(item.get("Harga Satuan (Rp)") or item.get("satuan_harga") or item.get("satuan") or 0.0)
        except (ValueError, TypeError):
            harga = 0.0

        nilai = vol * harga
        total_anggaran_kegiatan += nilai

        for idx, w in enumerate(col_widths):
            row[idx].width = w

        row[0].text = nama
        row[1].text = f"{vol:g}"
        row[2].text = ukuran
        row[3].text = format_rupiah(harga)
        row[4].text = format_rupiah(nilai)

        row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        row[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        for cell in row:
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.font.size = Pt(9)
                    run.font.name = "Arial"

    # Baris Total Akumulasi
    tot_row = rab_table.add_row().cells
    tot_row[0].merge(tot_row[3])
    tot_row[0].text = "TOTAL ANGGARAN"
    tot_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    tot_row[0].paragraphs[0].runs[0].bold = True
    tot_row[0].paragraphs[0].runs[0].font.size = Pt(9)
    tot_row[0].paragraphs[0].runs[0].font.name = "Arial"

    tot_row[4].text = f"Rp {format_rupiah(total_anggaran_kegiatan)}"
    tot_row[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tot_row[4].paragraphs[0].runs[0].bold = True
    tot_row[4].paragraphs[0].runs[0].font.size = Pt(9)
    tot_row[4].paragraphs[0].runs[0].font.name = "Arial"

    set_cell_background(tot_row[0], "F2F2F2")
    set_cell_background(tot_row[4], "F2F2F2")

    doc.add_paragraph()

    # --- 9. PENUTUP ---
    doc.add_heading("H. PENUTUP", level=2)
    doc.add_paragraph(str(data.get("penutup_default", "Demikian Kerangka Acuan Kerja (KAK) ini disusun dengan sebaik-baiknya untuk digunakan sebagaimana mestinya.")))
    for p_extra in data.get("penutup_tambahan", []):
        if str(p_extra).strip():
            doc.add_paragraph(str(p_extra).strip())

    # --- 10. LEMBAR PENGESAHAN / PENANDATANGANAN ---
    doc.add_paragraph()
    p_tgl = doc.add_paragraph(f"Kepi, {data.get('tanggal_penandatangan', '')}")
    p_tgl.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    sign_table = doc.add_table(rows=1, cols=2)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_cells = sign_table.rows[0].cells
    s_cells[0].width = Inches(3.5)
    s_cells[1].width = Inches(3.5)

    # Kolom Kiri: Ketua Tim
    jabatan_kt = data.get("jabatan_ketua_tim", "Ketua Tim")
    nama_kt = data.get("nama_ketua_tim", "")
    nip_kt = data.get("nip_ketua_tim", "")
    s_cells[0].paragraphs[0].text = f"{jabatan_kt}\n\n\n\n\n{nama_kt}\nNIP. {nip_kt}"
    s_cells[0].paragraphs[0].runs[0].font.name = "Arial"
    s_cells[0].paragraphs[0].runs[0].font.size = Pt(10)

    # Kolom Kanan: PPK
    ppk_nama = data.get("ppk_nama", "")
    ppk_nip = data.get("ppk_nip", "")
    s_cells[1].paragraphs[0].text = f"Pejabat Pembuat Komitmen\nBPS Kabupaten Mappi\n\n\n\n{ppk_nama}\nNIP. {ppk_nip}"
    s_cells[1].paragraphs[0].runs[0].font.name = "Arial"
    s_cells[1].paragraphs[0].runs[0].font.size = Pt(10)

    # --- 11. LEMBAR LAMPIRAN DINAMIS ---
    for lampiran in data.get("daftar_lampiran", []):
        nama_lamp = str(lampiran).strip()
        if nama_lamp:
            doc.add_page_break()
            p_lamp = doc.add_paragraph()
            p_lamp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_lamp.paragraph_format.space_before = Pt(24)
            
            r_lamp = p_lamp.add_run(f"LAMPIRAN: {nama_lamp.upper()}\n")
            r_lamp.bold = True
            r_lamp.font.size = Pt(12)
            r_lamp.font.name = "Arial"

    # Baris ini WAJIB berada di luar perulangan (menjorok 4 spasi dari margin kiri)
    return doc