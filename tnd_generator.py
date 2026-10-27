import os
import datetime

import streamlit as st
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT


def format_tgl_indo(date_obj, include_year=True):
    bulan = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
             "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    if include_year:
        return f"{date_obj.day} {bulan[date_obj.month]} {date_obj.year}"
    return f"{date_obj.day} {bulan[date_obj.month]}"


def add_kop_surat(doc):
    # 1. Create a 2-column invisible table (3.5 cm + 12.5 cm = 16 cm text width)
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Cm(3.5)
    table.columns[1].width = Cm(12.5)
    # Word honours per-cell widths, so set them too
    table.cell(0, 0).width = Cm(3.5)
    table.cell(0, 1).width = Cm(12.5)

    # 2. Insert Logo in Left Cell
    cell_left = table.cell(0, 0)
    cell_left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p_logo = cell_left.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_after = Pt(0)
    p_logo.paragraph_format.line_spacing = 1.0
    try:
        r_logo = p_logo.add_run()
        # Width only, so the aspect ratio stays locked
        r_logo.add_picture("logo_bpkp.png", width=Cm(3.0))
    except Exception:
        p_logo.add_run("[LOGO BPKP MISSING]").bold = True

    # 3. Insert Text in Right Cell
    cell_right = table.cell(0, 1)

    def fmt(p):
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0  # Single spacing for header

    def run(p, text, size, bold=False, italic=False):
        r = p.add_run(text)
        r.font.name = 'Arial'
        r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
        return r

    p1 = cell_right.paragraphs[0]
    fmt(p1)
    run(p1, "BADAN PENGAWASAN KEUANGAN DAN PEMBANGUNAN", 12, bold=True)

    p2 = cell_right.add_paragraph()
    fmt(p2)
    run(p2, "PERWAKILAN PROVINSI PAPUA", 12, bold=True)

    p3 = cell_right.add_paragraph()
    fmt(p3)
    run(p3, "Jalan Pasifik Indah III, Pasir Dua, Jayapura, Papua 99117\nTelepon (0967) 5144006, ", 9)
    run(p3, "Email", 9, italic=True)
    run(p3, " papua@bpkp.go.id\n", 9)
    run(p3, "Website", 9, italic=True)
    run(p3, " www.bpkp.go.id/papua", 9)

    # 4. Thick bottom border attached directly to the table
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        tblBorders = OxmlElement('w:tblBorders')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '24')  # 24 = 3.0 pt thickness
        bottom.set(qn('w:space'), '0')
        bottom.set(qn('w:color'), 'auto')
        tblBorders.append(bottom)
        tblPr[0].append(tblBorders)


def add_metadata_surat(doc, context):
    # Zero-height anchor paragraph: without any paragraph between them, Word merges
    # this table into the Kop Surat table. It occupies ~1 pt, so no visible gap.
    anchor = doc.add_paragraph()
    anchor.paragraph_format.space_before = Pt(0)
    anchor.paragraph_format.space_after = Pt(0)
    anchor.paragraph_format.line_spacing = Pt(1)
    anchor.add_run().font.size = Pt(1)

    widths = [Cm(2.5), Cm(0.5), Cm(9.0), Cm(4.0)]  # total 16.0 cm
    table = doc.add_table(rows=3, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False

    # Borderless table
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        tblBorders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement(f'w:{edge}')
            el.set(qn('w:val'), 'nil')
            tblBorders.append(el)
        tblPr[0].append(tblBorders)

    for i, w in enumerate(widths):
        table.columns[i].width = w
    for row in table.rows:
        for cell, w in zip(row.cells, widths):
            cell.width = w

    data = [
        ("Nomor", context['nomor_surat'], context['tanggal_surat']),
        ("Lampiran", context['lampiran'], ""),
        ("Hal", context['hal'], ""),
    ]
    for row, (label, value, tanggal) in zip(table.rows, data):
        for cell, text in zip(row.cells, (label, ":", value, tanggal)):
            if text:
                cell.paragraphs[0].add_run(text)

    # Strict single spacing for every paragraph in every cell
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.0

    # Date cell explicitly right-aligned
    table.cell(0, 3).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Apply 4pt space before the first row to create a precise gap below the header line
    for cell in table.rows[0].cells:
        for p in cell.paragraphs:
            p.paragraph_format.space_before = Pt(4)

    p_gap_yth = doc.add_paragraph()
    p_gap_yth.paragraph_format.space_after = Pt(0)
    p_gap_yth.paragraph_format.space_before = Pt(0)
    p_gap_yth.paragraph_format.line_spacing = 1.0

    p_yth = doc.add_paragraph("Yth. " + context['yth'])
    p_yth.paragraph_format.space_after = Pt(0)
    p_yth.paragraph_format.space_before = Pt(0)
    p_yth.paragraph_format.line_spacing = 1.0

    p_di = doc.add_paragraph("di " + context['tempat_tujuan'])
    p_di.paragraph_format.space_after = Pt(0)
    p_di.paragraph_format.space_before = Pt(0)
    p_di.paragraph_format.line_spacing = 1.0


def add_dasar_penugasan(doc, dasar_text):
    # Paragraph "Berdasarkan:"
    p_dasar = doc.add_paragraph("Berdasarkan:")
    p_dasar.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_dasar.paragraph_format.first_line_indent = Cm(1.25)
    p_dasar.paragraph_format.space_after = Pt(0)
    p_dasar.paragraph_format.line_spacing = 1.5

    # Numbered List with Hanging Indent
    dasar_list = [d.strip() for d in dasar_text.split('\n') if d.strip()]
    for i, text in enumerate(dasar_list, start=1):
        p_item = doc.add_paragraph()
        p_item.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_item.paragraph_format.left_indent = Cm(0.75)
        p_item.paragraph_format.first_line_indent = Cm(-0.75)
        p_item.paragraph_format.space_after = Pt(0)
        p_item.paragraph_format.line_spacing = 1.5
        p_item.add_run(f"{i}.\t{text}")


def add_body_paragraphs(doc, context):
    # Paragraph 1: kami menugasi (Flush left, NO indent)
    p1 = doc.add_paragraph(f"kami menugasi tim untuk melaksanakan {context['jenis_penugasan']} dengan susunan tim sesuai surat tugas terlampir.")
    p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p1.paragraph_format.space_after = Pt(0)
    p1.paragraph_format.line_spacing = 1.5

    # Paragraph 2: Penugasan ini dilaksanakan (First-line indent)
    p2 = doc.add_paragraph(f"Penugasan ini dilaksanakan selama {context['lama_hari']} hari kerja, terhitung mulai tanggal {context['tgl_mulai']} s.d. {context['tgl_selesai']}.")
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p2.paragraph_format.first_line_indent = Cm(1.25)
    p2.paragraph_format.space_after = Pt(0)
    p2.paragraph_format.line_spacing = 1.5

    # Paragraph 3: Seluruh biaya (First-line indent)
    p3 = doc.add_paragraph(f"Seluruh biaya yang berkaitan dengan penugasan ini menjadi beban anggaran {context['beban_biaya']}.")
    p3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p3.paragraph_format.first_line_indent = Cm(1.25)
    p3.paragraph_format.space_after = Pt(0)
    p3.paragraph_format.line_spacing = 1.5

    # Paragraph 4: Gratifikasi (First-line indent)
    p4 = doc.add_paragraph("Pegawai BPKP dalam melaksanakan tugas tidak menerima/meminta gratifikasi dan/atau suap.")
    p4.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p4.paragraph_format.first_line_indent = Cm(1.25)
    p4.paragraph_format.space_after = Pt(0)
    p4.paragraph_format.line_spacing = 1.5

    # Paragraph 5: Penugasan ini agar dilaksanakan (First-line indent)
    p5 = doc.add_paragraph("Penugasan ini agar dilaksanakan dengan penuh tanggung jawab.")
    p5.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p5.paragraph_format.first_line_indent = Cm(1.25)
    p5.paragraph_format.space_after = Pt(0)
    p5.paragraph_format.line_spacing = 1.5


def add_closing_and_signature(doc, context):
    # Determine indent for the first line ONLY
    is_plt_plh = context['jabatan_ttd'].startswith("Plt.") or context['jabatan_ttd'].startswith("Plh.")
    jabatan_indent = Cm(7.25) if is_plt_plh else Cm(8.0)
    base_indent = Cm(8.0)  # Always 8.0 for everything below the title

    # 1-line gap before signature block
    p_gap_sig = doc.add_paragraph()
    p_gap_sig.paragraph_format.space_after = Pt(0)
    p_gap_sig.paragraph_format.space_before = Pt(0)
    p_gap_sig.paragraph_format.line_spacing = 1.0

    # Update helper to accept a dynamic indent value
    def set_signature_format(p, indent_val):
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = indent_val
        p.paragraph_format.right_indent = Cm(0)
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0

    # 1. Jabatan (Uses dynamic indent)
    p_jab = doc.add_paragraph(context['jabatan_ttd'] + ",")
    set_signature_format(p_jab, jabatan_indent)

    # Spacer (Uses base_indent)
    p_gap1 = doc.add_paragraph()
    set_signature_format(p_gap1, base_indent)

    # 2. TTE Spacer / Text (All use base_indent)
    if context.get('use_tte', True):
        p_tte = doc.add_paragraph()
        set_signature_format(p_tte, base_indent)
        run_tte = p_tte.add_run("Ditandatangani secara elektronik oleh")
        run_tte.font.name = 'Arial'
        run_tte.font.size = Pt(10)
        run_tte.font.italic = True
        run_tte.font.color.rgb = RGBColor(166, 166, 166)

        p_gap2 = doc.add_paragraph()
        set_signature_format(p_gap2, base_indent)
    else:
        p_gap2 = doc.add_paragraph()
        set_signature_format(p_gap2, base_indent)
        p_gap3 = doc.add_paragraph()
        set_signature_format(p_gap3, base_indent)

    # 3. Nama (Uses base_indent)
    p_nama = doc.add_paragraph(context['nama_ttd'])
    set_signature_format(p_nama, base_indent)

    # --- TEMBUSAN LOGIC ---
    # Parse items and auto-detect count
    items = [t.strip() for t in context['daftar_tembusan'].split('\n') if t.strip()]

    # Smart helper for Hanging Indent Tembusan Items (Multiple)
    def add_tembusan_items(parent_doc, tembusan_list):
        for i, text in enumerate(tembusan_list, start=1):
            # Strip existing punctuation to prevent double symbols
            clean_text = text.rstrip(';.')

            # Add ';' for items 1 to N-1, and '.' for the final item
            if i < len(tembusan_list):
                final_text = f"{clean_text};"
            else:
                final_text = f"{clean_text}."

            p = parent_doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Cm(0.75)
            p.paragraph_format.first_line_indent = Cm(-0.75)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            p.add_run(f"{i}.\t{final_text}")

    if len(items) == 1:
        # RULE 8a: Single Item (Auto-detect)
        p_tembusan = doc.add_paragraph("Tembusan:")
        p_tembusan.paragraph_format.left_indent = Cm(0)
        p_tembusan.paragraph_format.first_line_indent = Cm(0)
        p_tembusan.paragraph_format.space_before = Pt(12)
        p_tembusan.paragraph_format.space_after = Pt(0)
        p_tembusan.paragraph_format.line_spacing = 1.0

        # No numbers, add "Yth.", no punctuation at the end
        p_item = doc.add_paragraph(f"Yth. {items[0].rstrip(';.')}")
        p_item.paragraph_format.left_indent = Cm(0)
        p_item.paragraph_format.first_line_indent = Cm(0)
        p_item.paragraph_format.space_after = Pt(0)
        p_item.paragraph_format.line_spacing = 1.0

    elif len(items) > 1:
        if context['posisi_tembusan'] == "Halaman 1 (Langsung)":
            # RULE 8b: Multiple Items on Page 1
            p_tembusan = doc.add_paragraph("Tembusan Yth.:")
            p_tembusan.paragraph_format.left_indent = Cm(0)
            p_tembusan.paragraph_format.first_line_indent = Cm(0)
            p_tembusan.paragraph_format.space_before = Pt(12)
            p_tembusan.paragraph_format.space_after = Pt(0)
            p_tembusan.paragraph_format.line_spacing = 1.0

            add_tembusan_items(doc, items)

        elif context['posisi_tembusan'] == "Halaman 2 (Daftar Terlampir)":
            # RULE 8c: Multiple Items on Page 2 (Header only on Page 1)
            p_tembusan = doc.add_paragraph("Tembusan Yth. (Daftar Terlampir)")
            p_tembusan.paragraph_format.left_indent = Cm(0)
            p_tembusan.paragraph_format.first_line_indent = Cm(0)
            p_tembusan.paragraph_format.space_before = Pt(12)
            p_tembusan.paragraph_format.space_after = Pt(0)
            p_tembusan.paragraph_format.line_spacing = 1.0

    # --- HOTLINE LOGIC ---
    p_hotline = doc.add_paragraph()
    p_hotline.paragraph_format.left_indent = Cm(0)
    p_hotline.paragraph_format.first_line_indent = Cm(0)
    p_hotline.paragraph_format.space_before = Pt(12)
    p_hotline.paragraph_format.space_after = Pt(0)
    p_hotline.paragraph_format.line_spacing = 1.0

    run_hl_title = p_hotline.add_run("Layanan Pengaduan/Hotline:\n")
    run_hl_title.font.name = 'Arial'
    run_hl_title.font.size = Pt(10)

    run_hl_desc = p_hotline.add_run("Apabila Pegawai BPKP menerima/meminta gratifikasi dan/atau suap, dapat dilaporkan melalui mekanisme penyampaian pengaduan pada wbs.bpkp.go.id. atau menghubungi 081389101133.")
    run_hl_desc.font.name = 'Arial'
    run_hl_desc.font.size = Pt(10)

    # --- PAGE 2 LAMPIRAN LOGIC (VIA HEADER) ---
    if len(items) > 1 and context['posisi_tembusan'] == "Halaman 2 (Daftar Terlampir)":
        # Add a new section (this automatically creates a Next Page Section Break)
        new_section = doc.add_section()

        # Adjust Margins specifically for Section 2
        new_section.header_distance = Cm(0.5)  # Pushes Header up to 0.5cm from top edge
        new_section.top_margin = Cm(2.5)       # Pulls Body text up, removing huge gaps

        # Unlink the header so Page 1 remains completely clean
        header = new_section.header
        header.is_linked_to_previous = False

        # Write to the Header with a shifted left indent (7.0 cm) to prevent text wrapping
        p_lamp = header.paragraphs[0] if len(header.paragraphs) > 0 else header.add_paragraph()
        p_lamp.text = "" 
        p_lamp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_lamp.paragraph_format.left_indent = Cm(7.0)  # Shifted left to 7.0cm
        p_lamp.paragraph_format.right_indent = Cm(0)
        p_lamp.paragraph_format.first_line_indent = Cm(0)
        p_lamp.paragraph_format.space_before = Pt(0)
        p_lamp.paragraph_format.space_after = Pt(0)
        p_lamp.paragraph_format.line_spacing = 1.0

        run1 = p_lamp.add_run("Lampiran\n")
        run2 = p_lamp.add_run("Surat Kepala Perwakilan BPKP Provinsi Papua\n")
        run3 = p_lamp.add_run(f"Nomor: {context['nomor_surat']}\n")
        run4 = p_lamp.add_run(f"Tanggal: {context['tanggal_surat']}")

        # Force font size 11 for the header
        for r in p_lamp.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(11)

        # Add exactly 1 blank line with 1.0 spacing before the title
        p_gap_judul = doc.add_paragraph()
        p_gap_judul.paragraph_format.space_before = Pt(0)
        p_gap_judul.paragraph_format.space_after = Pt(0)
        p_gap_judul.paragraph_format.line_spacing = 1.0

        # Write to the Body of Section 2
        p_judul = doc.add_paragraph("DAFTAR TEMBUSAN")
        p_judul.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_judul.paragraph_format.space_before = Pt(0) 
        p_judul.paragraph_format.space_after = Pt(12)
        p_judul.runs[0].bold = True

        add_tembusan_items(doc, items)


st.sidebar.title("🏠 TND Generator")
menu = st.sidebar.radio("Pilih Naskah Dinas:", ["Surat Pengantar", "Laporan (Coming Soon)"])

if menu == "Surat Pengantar":
    st.title("Drafting Surat Pengantar")

    # Build the Form
    with st.form("tnd_form"):
        st.subheader("Data Naskah Dinas")
        col1, col2 = st.columns(2)
        with col1:
            nomor_surat = st.text_input("Nomor Surat", value="")
            lampiran = st.text_input("Lampiran", value="")
            hal = st.text_input("Hal", value="")
            tanggal_surat_input = st.date_input("Tanggal Dokumen", datetime.date.today())

        with col2:
            yth = st.text_input("Yth.", value="")
            tempat_tujuan = st.text_input("Kota Tujuan (di)", value="")

        st.subheader("Isi Surat")
        ada_dasar_lain = st.checkbox("Tambahkan Dasar Aturan Lain (selain PP 60 & Perpres 2)", value=False)
        dasar_aturan = st.text_area("Isi Dasar Aturan Tambahan (jika dicentang)", value="")

        jenis_penugasan = st.text_input("kami menugasi tim untuk melaksanakan:", value="")

        col3, col4, col5 = st.columns(3)
        with col3:
            lama_hari = st.text_input("Lama Hari Kerja", value="", placeholder="misal: 36 (tiga puluh enam)")
        with col4:
            tgl_mulai_input = st.date_input("Mulai Tanggal", datetime.date.today())
        with col5:
            tgl_selesai_input = st.date_input("S.d. Tanggal", datetime.date.today())

        beban_biaya = st.text_input("Beban Anggaran", value="")

        st.subheader("Penandatangan")
        col6, col7 = st.columns(2)
        with col6:
            jabatan_ttd = st.text_input("Jabatan", value="")
            nama_ttd = st.text_input("Nama", value="")
        with col7:
            checker_tte = st.checkbox("Gunakan Tanda Tangan Elektronik (TTE)", value=True)
            checker_plt = st.selectbox("Status Jabatan", ["Definitif", "Plt.", "Plh."])

        st.subheader("Daftar Tembusan")
        posisi_tembusan = st.radio("Posisi Tembusan:", ["Halaman 1 (Langsung)", "Halaman 2 (Daftar Terlampir)"])
        daftar_tembusan = st.text_area("Isi Tembusan (satu baris per nama)")

        submit_button = st.form_submit_button("Generate Draft Surat")

    # Logic when form is submitted
    if submit_button:
        if not nomor_surat or not hal or not nama_ttd:
            st.error("⚠️ Harap isi form yang kosong terlebih dahulu!")
        else:
            # Update Context dynamically
            jabatan_raw = f"{checker_plt + ' ' if checker_plt != 'Definitif' else ''}{jabatan_ttd}"

            context = {
                'nomor_surat': nomor_surat,
                'tanggal_surat': format_tgl_indo(tanggal_surat_input),
                'lampiran': lampiran,
                'hal': hal,
                'yth': yth,
                'tempat_tujuan': tempat_tujuan,
                'jenis_penugasan': jenis_penugasan,
                'lama_hari': lama_hari,
                'tgl_mulai': format_tgl_indo(tgl_mulai_input, include_year=(tgl_mulai_input.year != tgl_selesai_input.year)),
                'tgl_selesai': format_tgl_indo(tgl_selesai_input),
                'beban_biaya': beban_biaya,
                'jabatan_ttd': jabatan_raw.strip(),  # STRIP to prevent center-alignment shifting
                'nama_ttd': nama_ttd.strip(),
                'use_tte': checker_tte,  # Capture the boolean from the checkbox
            }
            context['posisi_tembusan'] = posisi_tembusan
            context['daftar_tembusan'] = daftar_tembusan

            # Handle Dasar Instruksi conditionally with STRICT comma at the end
            dasar_tambahan = dasar_aturan.strip().rstrip(';.,')

            if ada_dasar_lain and dasar_tambahan:
                dasar_instruksi_final = (
                    "Peraturan Pemerintah Nomor 60 Tahun 2008 tentang Sistem Pengendalian Intern Pemerintah;\n"
                    "Peraturan Presiden Nomor 2 Tahun 2025 tentang Perubahan Kedua atas Peraturan Presiden Nomor 192 Tahun 2014 tentang Badan Pengawasan Keuangan dan Pembangunan; dan\n"
                    f"{dasar_tambahan},"
                )
            else:
                dasar_instruksi_final = (
                    "Peraturan Pemerintah Nomor 60 Tahun 2008 tentang Sistem Pengendalian Intern Pemerintah; dan\n"
                    "Peraturan Presiden Nomor 2 Tahun 2025 tentang Perubahan Kedua atas Peraturan Presiden Nomor 192 Tahun 2014 tentang Badan Pengawasan Keuangan dan Pembangunan,"
                )

            doc = docx.Document()

            # Global font: Arial 12
            style = doc.styles["Normal"]
            style.font.name = "Arial"
            style.font.size = Pt(12)

            # Page setup: A4 + margins
            for section in doc.sections:
                section.page_width = Cm(21.0)
                section.page_height = Cm(29.7)
                section.top_margin = Cm(1.0)
                section.bottom_margin = Cm(2.5)
                section.left_margin = Cm(3.0)
                section.right_margin = Cm(2.0)

            add_kop_surat(doc)
            add_metadata_surat(doc, context)

            # Gap before Berdasarkan
            p_gap_dasar = doc.add_paragraph()
            p_gap_dasar.paragraph_format.space_after = Pt(0)
            p_gap_dasar.paragraph_format.space_before = Pt(0)
            p_gap_dasar.paragraph_format.line_spacing = 1.0

            add_dasar_penugasan(doc, dasar_instruksi_final)
            add_body_paragraphs(doc, context)
            add_closing_and_signature(doc, context)

            preview_path = os.path.join(os.getcwd(), "preview_sementara.docx")
            doc.save(preview_path)
            os.startfile(preview_path)
            st.success("File Word berhasil dirakit dan terbuka otomatis di layar Anda.")

elif menu == "Laporan (Coming Soon)":
    st.title("Generator Laporan")
    st.info("Modul Laporan masih dalam tahap pengembangan. Nantikan *update* selanjutnya!")
