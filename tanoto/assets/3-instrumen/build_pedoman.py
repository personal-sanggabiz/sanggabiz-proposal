"""Bangun Pedoman Pendampingan Teknis (DOCX -> PDF): dokumen payung yang menaungi
Rubrik Observasi PPL dan Panduan FGD Pemetaan PPL di folder yang sama, dan menjelaskan
cara kerja metodologi Evidence-Based Continuous Quality Improvement (versi 2, 5 fase)
secara utuh. Ini CONTOH PEDOMAN disusun Sanggabiz, bukan Pedoman Pendampingan 2026 resmi
dari Direktorat PPG.
"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = os.path.dirname(os.path.abspath(__file__))
NAVY = RGBColor(0x09, 0x26, 0x47)
TEAL = RGBColor(0x2F, 0x7E, 0x9D)
GOLD = RGBColor(0xC8, 0x8A, 0x00)
GREY = RGBColor(0x5B, 0x66, 0x70)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x22, 0x22, 0x22)


def shade_cell(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_cell_border(cell, color="B0C0CF", sz=4):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for tag in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{tag}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def base_doc():
    d = Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.9)
    sec.top_margin = Cm(1.6)
    sec.bottom_margin = Cm(1.6)
    normal = d.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10)
    return d


def add_page_break(d):
    d.add_page_break()


def cover(d):
    for _ in range(4):
        d.add_paragraph()
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("SANGGABIZ")
    r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = NAVY

    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("CONTOH PEDOMAN KERJA")
    r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = GOLD
    p.paragraph_format.space_before = Pt(12)

    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Pedoman Pendampingan Teknis")
    r.font.size = Pt(28); r.font.bold = True; r.font.color.rgb = NAVY
    p.paragraph_format.space_before = Pt(6)

    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("LPTK Penyelenggara PPG, 2026")
    r.font.size = Pt(16); r.font.color.rgb = TEAL
    p.paragraph_format.space_after = Pt(20)

    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Metodologi Evidence-Based Continuous Quality Improvement")
    r.font.size = Pt(12); r.font.italic = True; r.font.color.rgb = GREY
    p.paragraph_format.space_after = Pt(60)

    tbl = d.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    shade_cell(cell, "FFF4C2")
    set_cell_border(cell, "C8930A", 6)
    cell.width = Cm(15)
    p = cell.paragraphs[0]
    r = p.add_run("STATUS: DRAF/CONTOH, BELUM FINAL. Disusun Sanggabiz berdasar RFP dan TOR untuk menunjukkan kesiapan kerja pada Proposal Teknis RFP No. 0648/TF-JKT/VIII/2026. Masih memerlukan masukan dan penyempurnaan lebih lanjut dari tim sebelum dipakai final. Ini bukan Pedoman Pendampingan 2026 resmi dari Direktorat PPG, dan akan diselaraskan dengan pedoman resmi tersebut saat kick-off.")
    r.font.size = Pt(9.5); r.font.italic = True; r.font.color.rgb = RGBColor(0x5A, 0x42, 0x06)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.margin_top = Cm(0.3); cell.margin_bottom = Cm(0.3); cell.margin_left = Cm(0.3); cell.margin_right = Cm(0.3)

    for _ in range(6):
        d.add_paragraph()
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Versi 1, 30 September 2026")
    r.font.size = Pt(9.5); r.font.color.rgb = GREY
    p2 = d.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p2.add_run("Disusun untuk Tanoto Foundation")
    r.font.size = Pt(9.5); r.font.color.rgb = GREY


def running_header(d, text):
    p = d.add_paragraph()
    r = p.add_run("SANGGABIZ")
    r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = NAVY
    r2 = p.add_run("   |   Pedoman Pendampingan Teknis, draf/contoh, belum final")
    r2.font.size = Pt(8); r2.font.color.rgb = GREY
    p.paragraph_format.space_after = Pt(10)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "10"); bottom.set(qn("w:color"), "B0C0CF")
    pbdr.append(bottom)
    pPr.append(pbdr)


def h1(d, num, text):
    p = d.add_paragraph()
    r = p.add_run(f"{num}.  {text}")
    r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = NAVY
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "14"); bottom.set(qn("w:color"), "FFC00C")
    pbdr.append(bottom)
    pPr.append(pbdr)


def h2(d, text):
    p = d.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(11.5); r.font.bold = True; r.font.color.rgb = TEAL
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)


def body(d, text, size=10, italic=False, color=None, space_after=6, bold=False):
    p = d.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(size); r.font.italic = italic; r.font.bold = bold
    r.font.color.rgb = color or INK
    p.paragraph_format.space_after = Pt(space_after)
    return p


def bullets(d, items, size=10):
    for it in items:
        p = d.add_paragraph(style="List Bullet")
        if isinstance(it, tuple):
            lead, rest = it
            r = p.add_run(lead); r.font.bold = True; r.font.size = Pt(size); r.font.color.rgb = NAVY
            r2 = p.add_run(rest); r2.font.size = Pt(size)
        else:
            r = p.add_run(it); r.font.size = Pt(size)
        p.paragraph_format.space_after = Pt(3)


def numbered(d, items, size=10):
    for it in items:
        p = d.add_paragraph(style="List Number")
        r = p.add_run(it); r.font.size = Pt(size)
        p.paragraph_format.space_after = Pt(3)


def styled_table(d, headers, rows, col_widths=None, header_fill="092647", font_size=9):
    tbl = d.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = True
    for j, htext in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        shade_cell(cell, header_fill)
        set_cell_border(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        r = p.add_run(htext); r.font.bold = True; r.font.size = Pt(font_size); r.font.color.rgb = WHITE
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = tbl.rows[i].cells[j]
            set_cell_border(cell)
            if i % 2 == 0:
                shade_cell(cell, "EEF2F6")
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            r = p.add_run(str(val)); r.font.size = Pt(font_size)
    if col_widths:
        for row in tbl.rows:
            for j, w in enumerate(col_widths):
                row.cells[j].width = Cm(w)
    return tbl


def note_box(d, text, fill="FFF4C2", border="C8930A", tcolor=RGBColor(0x5A, 0x42, 0x06)):
    tbl = d.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    shade_cell(cell, fill)
    set_cell_border(cell, border, 6)
    p = cell.paragraphs[0]
    r = p.add_run(text)
    r.font.size = Pt(9); r.font.italic = True; r.font.color.rgb = tcolor
    cell.margin_top = Cm(0.2); cell.margin_bottom = Cm(0.2); cell.margin_left = Cm(0.15); cell.margin_right = Cm(0.15)
    d.add_paragraph().paragraph_format.space_after = Pt(4)


def footer_note(d, text):
    d.add_paragraph().paragraph_format.space_after = Pt(2)
    p = d.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(7.5); r.font.color.rgb = GREY; r.font.italic = True


def build():
    d = base_doc()
    cover(d)
    add_page_break(d)

    # ================= 1. PENGANTAR =================
    running_header(d, "")
    h1(d, 1, "Pengantar dan Dasar")
    body(d, "Pedoman ini menjelaskan cara kerja tim pelaksana dalam mengoordinasikan pendampingan teknis di 12 LPTK penyelenggara PPG, sesuai RFP No. 0648/TF-JKT/VIII/2026 dari Tanoto Foundation dan TOR PPG Technical Assistance 2026. Pedoman ini adalah dokumen kerja internal tim, bukan pengganti Pedoman Pendampingan 2026 resmi yang diterbitkan Direktorat PPG.")
    h2(d, "1.1 Tujuan")
    bullets(d, [
        "Memberi satu acuan kerja yang sama bagi seluruh tim, sehingga setiap kunjungan dan laporan memakai standar yang konsisten.",
        "Menjelaskan kapan dan bagaimana tiap instrumen dipakai, siapa penanggung jawabnya, dan luaran apa yang dihasilkan.",
        "Menjadi dokumen hidup yang diperbarui begitu Pedoman Pendampingan 2026 resmi diterima dari Direktorat PPG.",
    ])
    h2(d, "1.2 Ruang Lingkup")
    body(d, "Pedoman ini mencakup seluruh pekerjaan mitra pelaksana sesuai RFP bagian C.4: menyiapkan perangkat pendampingan, mengoordinasikan kunjungan luring 12 LPTK, memfasilitasi tindak lanjut dan pemantauan, menyusun studi komparatif, serta menyusun analisis, laporan, dan rekomendasi. Pendampingan substantif ke LPTK sendiri dilakukan oleh Tim Pendamping Nasional yang ditetapkan Direktorat PPG, bukan oleh mitra pelaksana.")
    h2(d, "1.3 Dasar Rujukan")
    styled_table(d, ["Dokumen", "Fungsi"], [
        ["RFP No. 0648/TF-JKT/VIII/2026", "Dasar kontraktual: ruang lingkup, luaran, anggaran, jadwal"],
        ["TOR PPG Technical Assistance 2026", "Konteks, temuan Supervisi 2025, Teori Perubahan, daftar LPTK"],
        ["Permendikbudristek 53/2023", "Dasar SPMI dan siklus PPEPP"],
        ["IAPSK 3.0 ISK LAMDIK, 26 Mei 2026", "Rubrik dan ambang elemen akreditasi Unggul"],
        ["Pedoman Pendampingan 2026 (Direktorat PPG)", "Rujukan utama begitu diterbitkan resmi, menggantikan asumsi di pedoman ini"],
    ], col_widths=[7, 10])

    # ================= 2. METODOLOGI =================
    add_page_break(d)
    running_header(d, "")
    h1(d, 2, "Metodologi: Evidence-Based Continuous Quality Improvement")
    body(d, "Metodologi ini menggabungkan dua prinsip yang sudah mapan. Continuous Quality Improvement adalah cara kerja siklus mutu yang berakar dari Plan-Do-Check-Act atau Siklus Deming, dan menjadi dasar filosofi SPMI atau Sistem Penjaminan Mutu Internal di Indonesia. Evidence-based practice adalah prinsip bahwa setiap rekomendasi dan kesimpulan harus berbasis bukti yang dapat ditelusuri, bukan asumsi. Pekerjaan tim dibagi menjadi lima fase berikut, yang dipetakan langsung ke lima poin ruang lingkup mitra pelaksana pada RFP bagian C.4.")

    styled_table(d, ["Fase", "Setara RFP C.4", "Aktivitas utama", "Luaran"], [
        ["01 Diagnose", "Poin 1 (awal)", "Telaah baseline dari hasil Supervisi 2025 dan RTL 12 LPTK, skor awal 1 sampai 4 pada lima elemen", "Baseline per LPTK"],
        ["02 Design", "Poin 1 (lanjutan)", "Penyusunan atau penyempurnaan pedoman, modul, dan instrumen, rencana kerja berklaster", "Instrumen, rencana kerja, jadwal"],
        ["03 Deliver", "Poin 2", "Pendampingan luring 12 LPTK, workshop, FGD, telaah teknis, penyusunan rencana aksi", "Laporan Pendampingan per LPTK"],
        ["04 Monitor & Verify", "Poin 3", "Pemantauan RTL, konsultasi berkala, verifikasi bukti perbaikan", "RTL Tracker, Evidence Register terisi"],
        ["05 Evaluate & Improve", "Poin 4 dan 5", "Studi komparatif, analisis pelaksanaan PPL, sintesis lintas LPTK, rekomendasi", "Paket Laporan Akhir"],
    ], col_widths=[3.2, 2.6, 7.5, 3.7], font_size=8.7)

    note_box(d, "Fase 02 Design menghasilkan Rubrik Observasi PPL dan Panduan FGD Pemetaan PPL, yaitu dua instrumen di folder yang sama dengan dokumen ini. Keduanya dipakai langsung pada Fase 03 Deliver.")

    # ================= 3. TATA KELOLA =================
    h1(d, 3, "Tata Kelola dan Batas Peran")
    h2(d, "3.1 Tiga Lapis Pengawasan")
    styled_table(d, ["Lapis", "Anggota", "Fungsi"], [
        ["Steering Committee", "Kemendikdasmen (Direktorat PPG), Tanoto Foundation, Direktur Sanggabiz", "Arah strategis, keputusan lingkup, jadwal, anggaran"],
        ["Technical Committee", "Perwakilan Direktorat PPG/TPN, unsur Kemdiktisaintek, Project Manager, Advisor, Analis", "Menjaga mutu luaran, instrumen, jadwal kunjungan"],
        ["Advisory Board", "Ahli PPG, akreditasi, penjaminan mutu, evaluasi", "Reviu metodologi dan draf laporan"],
    ], col_widths=[3.8, 7.5, 5.7], font_size=8.8)
    h2(d, "3.2 Batas Peran Mitra Pelaksana")
    body(d, "Mitra pelaksana tidak menggantikan Tim Pendamping Nasional dan tidak menjanjikan pendampingan akreditasi yang bersifat substantif. Peran mitra adalah memastikan pendampingan itu terkoordinasi, terdata, dan terbukti. Koordinasi dengan Kemdiktisaintek dilakukan melalui Direktorat PPG, bukan langsung, sesuai struktur tim pendampingan bersama.")

    # ================= 4. PROSEDUR PER FASE =================
    running_header(d, "")
    h1(d, 4, "Prosedur Kerja per Fase")

    h2(d, "Fase 01. Diagnose, sebelum kunjungan pertama")
    numbered(d, [
        "Telaah laporan Supervisi 2025 dan RTL untuk masing-masing dari 12 LPTK.",
        "Isi skor baseline 1 sampai 4 pada lima elemen memakai Readiness Scorecard, dengan status bukti tiap elemen dicatat di Evidence Register.",
        "Konfirmasi kelompok pembanding studi komparatif bersama Tanoto Foundation dan Direktorat PPG, mengacu TOR Annex 1 dan 2.",
        "Susun Laporan Awal, mencakup rencana kerja, pemetaan jadwal Tim Pendamping Nasional, dan instrumen yang akan dipakai.",
    ])

    h2(d, "Fase 02. Design, sebelum dan menjelang kunjungan")
    numbered(d, [
        "Sempurnakan atau susun instrumen: Rubrik Observasi PPL dan Panduan FGD Pemetaan PPL (lihat dua dokumen terkait di folder ini).",
        "Susun jadwal kunjungan berklaster mengikuti lima klaster lokasi RFP bagian C.3.",
        "Siapkan Paket Kesiapan Dokumen untuk tiap LPTK, dikirim tujuh hari sebelum kunjungan.",
        "Lakukan kick-off dan penyelarasan bersama Tanoto Foundation dan Direktorat PPG.",
    ])

    h2(d, "Fase 03. Deliver, saat kunjungan luring")
    numbered(d, [
        "Koordinator Lapangan mendampingi kunjungan Tim Pendamping Nasional: mengatur jadwal, kesiapan dokumen, dan kebutuhan operasional.",
        "Observasi PPL memakai Rubrik Observasi PPL, diisi bersama DPL dan guru pamong.",
        "FGD pemetaan PPL dilaksanakan memakai Panduan FGD, untuk empat kelompok informan.",
        "Setiap bukti yang terkumpul dicatat di Evidence Register dengan status: terverifikasi, diklaim, atau belum ada.",
        "Susun Laporan Pendampingan dan Paket Perbaikan Akreditasi, paling lambat 14 hari setelah kunjungan.",
    ])

    h2(d, "Fase 04. Monitor and Verify, berjalan sepanjang program")
    numbered(d, [
        "Perbarui RTL Tracker secara berkala: status tiap rencana tindak lanjut adalah rencana, berjalan, selesai, atau terbukti.",
        "Lakukan konsultasi berkala dengan Tim Pendamping Nasional dan LPTK untuk menghimpun bukti kemajuan.",
        "Verifikasi independen dilakukan sebelum status RTL dinaikkan dari selesai menjadi terbukti. Lihat indikator lengkap di bagian 5.",
    ])

    h2(d, "Fase 05. Evaluate and Improve, menjelang akhir program")
    numbered(d, [
        "Verifikasi kemajuan akhir dan analisis lintas 12 LPTK.",
        "Susun studi komparatif memakai data yang tersedia dan pengumpulan informasi daring, tanpa kunjungan lapangan tambahan.",
        "Susun sintesis praktik baik, hambatan, dan rekomendasi untuk siklus pendampingan berikutnya.",
        "Serahkan Paket Laporan Akhir dan repositori data, paling lambat 7 Desember 2026.",
    ])

    # ================= 5. INDIKATOR STATUS RTL =================
    add_page_break(d)
    running_header(d, "")
    h1(d, 5, "Indikator Status RTL")
    body(d, "Setiap rencana tindak lanjut, disingkat RTL, dinilai memakai empat status berikut. Status selesai dan terbukti dibedakan dengan sengaja, mengikuti praktik baku pelacakan tindak lanjut audit internal: tindakan yang sudah dikerjakan dan tindakan yang sudah diverifikasi efektif adalah dua tahap yang berbeda.")
    styled_table(d, ["Status", "Definisi", "Syarat"], [
        ["Rencana", "Belum dimulai", "RTL tercatat di rencana 2025, belum ada bukti tindakan apa pun"],
        ["Berjalan", "Sedang dikerjakan", "Ada bukti awal seperti draf dokumen atau notulen rapat, belum tuntas"],
        ["Selesai", "Diklaim tuntas oleh LPTK", "Output sudah jadi, namun ini laporan LPTK sendiri, status buktinya diklaim"],
        ["Terbukti", "Diverifikasi independen", "Mitra atau Tim Pendamping Nasional sudah mengecek langsung, status buktinya terverifikasi"],
    ], col_widths=[3, 5.5, 8.5], font_size=9)
    note_box(d, "Prinsip: selesai berarti klaim sepihak, terbukti berarti sudah dicek pihak lain. Klaim tanpa dokumen dicatat sebagai diklaim, bukan terverifikasi. Tidak ada bukti yang dibuat atau direkayasa.")

    # ================= 6. STANDAR MUTU =================
    h1(d, 6, "Standar Mutu Luaran dan Pengelolaan Data")
    h2(d, "6.1 Prinsip Mutu Luaran")
    bullets(d, [
        "Ikuti format yang disepakati Direktorat PPG dan Tanoto Foundation.",
        "Bedakan jelas antara kondisi awal, kondisi terkini, rekomendasi, tindakan, dan hasil.",
        "Setiap klaim didukung bukti yang dapat ditelusuri.",
        "Setiap tindak lanjut punya penanggung jawab dan target waktu.",
        "Gabungkan analisis kuantitatif dan kualitatif secara tepat.",
        "Serahkan PDF final dan versi yang dapat diedit.",
        "Akomodasi semua komentar Direktorat PPG dan Tanoto Foundation.",
    ])
    h2(d, "6.2 Pengelolaan Data")
    bullets(d, [
        "Repositori terpusat dengan akses berlapis menurut peran, penamaan dan versi baku, log akses dan pencadangan.",
        "Minimasi data pribadi. Persetujuan atau consent diminta untuk FGD dan rekaman.",
        "Tidak mempublikasikan data atau hasil program tanpa persetujuan tertulis Tanoto Foundation dan Direktorat PPG.",
    ])

    # ================= 7. DAFTAR INSTRUMEN =================
    add_page_break(d)
    running_header(d, "")
    h1(d, 7, "Daftar Instrumen dan Alat Kerja")
    styled_table(d, ["Alat", "Fase", "Bentuk", "Lokasi"], [
        ["Readiness Scorecard", "01 Diagnose, 04 Monitor & Verify", "Tabel skor 1 sampai 4 per elemen per LPTK", "Dashboard Kesiapan Akreditasi"],
        ["Rubrik Observasi PPL", "02 Design, 03 Deliver", "Dokumen cetak, enam dimensi PPL", "Rubrik-Observasi-PPL.pdf/.docx (folder ini)"],
        ["Panduan FGD Pemetaan PPL", "02 Design, 03 Deliver", "Dokumen cetak, empat sesi kelompok informan", "Panduan-FGD-Pemetaan-PPL.pdf/.docx (folder ini)"],
        ["Evidence Register", "03 Deliver, 04 Monitor & Verify", "Tabel bukti dengan status", "Dashboard Kesiapan Akreditasi"],
        ["RTL Tracker", "04 Monitor & Verify", "Daftar RTL dengan PIC, tenggat, status", "Dashboard Kesiapan Akreditasi"],
        ["Dashboard Kesiapan Akreditasi", "Semua fase", "Alat interaktif, ringkasan dan detail per LPTK", "assets/1-dashboard/ (index-v2.html)"],
    ], col_widths=[4, 3.8, 5, 4.2], font_size=8.5)

    # ================= 8. DAFTAR ISTILAH =================
    add_page_break(d)
    running_header(d, "")
    h1(d, 8, "Daftar Istilah")
    styled_table(d, ["Istilah", "Kepanjangan atau arti"], [
        ["LPTK", "Lembaga Pendidikan Tenaga Kependidikan"],
        ["PPG", "Pendidikan Profesi Guru"],
        ["TPN", "Tim Pendamping Nasional, ditetapkan Direktorat PPG"],
        ["RTL", "Rencana Tindak Lanjut, disusun tiap LPTK setelah Supervisi 2025"],
        ["PPL", "Praktik Pengalaman Lapangan"],
        ["DPL", "Dosen Pembimbing Lapangan"],
        ["SPMI", "Sistem Penjaminan Mutu Internal"],
        ["PPEPP", "Penetapan, Pelaksanaan, Evaluasi, Pengendalian, Peningkatan; lima tahap siklus SPMI"],
        ["OBE", "Outcome-Based Education, pendidikan berbasis capaian pembelajaran"],
        ["CPL", "Capaian Pembelajaran Lulusan"],
        ["DTPS", "Dosen Tetap Program Studi"],
        ["UPPS", "Unit Pengelola Program Studi"],
        ["LAMDIK", "Lembaga Akreditasi Mandiri Kependidikan"],
        ["Kemdiktisaintek", "Kementerian Pendidikan Tinggi, Sains, dan Teknologi"],
        ["Kemendikdasmen", "Kementerian Pendidikan Dasar dan Menengah, membawahi Direktorat PPG"],
    ], col_widths=[4, 13], font_size=9)

    footer_note(d, "Pedoman Pendampingan Teknis, versi 1 (draf), 30 September 2026. Disusun Sanggabiz untuk Proposal Teknis RFP No. 0648/TF-JKT/VIII/2026 kepada Tanoto Foundation. Status: draf/contoh, memerlukan masukan dan penyempurnaan lebih lanjut. Dokumen ini akan diperbarui begitu Pedoman Pendampingan 2026 resmi dari Direktorat PPG diterima.")

    out = os.path.join(HERE, "Pedoman-Pendampingan-Teknis.docx")
    d.save(out)
    return out


if __name__ == "__main__":
    p = build()
    print("saved", p)
