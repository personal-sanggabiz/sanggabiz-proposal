"""Bangun dua instrumen kerja (DOCX -> PDF) untuk Lampiran: Rubrik Observasi PPL & Panduan FGD Pemetaan PPL.
Isi ini adalah CONTOH INSTRUMEN yang disusun Sanggabiz berdasar RFP dan TOR. Ini bukan instrumen resmi Direktorat PPG.
Diselaraskan dengan metodologi Evidence-Based Continuous Quality Improvement versi 2 (lihat assets/4-slide8-metodologi-v2):
kedua instrumen ini adalah produk Fase 02 Design, dan dipakai langsung saat Fase 03 Deliver.
"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = os.path.dirname(os.path.abspath(__file__))
NAVY = RGBColor(0x09, 0x26, 0x47)
TEAL = RGBColor(0x2F, 0x7E, 0x9D)
YELLOW = RGBColor(0xC8, 0x8A, 0x00)  # emas lebih gelap biar terbaca di kertas putih
GREY = RGBColor(0x5B, 0x66, 0x70)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


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
    sec.left_margin = sec.right_margin = Cm(1.8)
    sec.top_margin = Cm(1.6)
    sec.bottom_margin = Cm(1.6)
    normal = d.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10)
    return d


def phase_chip(d, fase_disusun, fase_dipakai):
    tbl = d.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, (label, val) in enumerate([("Disusun pada", fase_disusun), ("Dipakai pada", fase_dipakai)]):
        cell = tbl.rows[0].cells[i]
        shade_cell(cell, "0D2A4E")
        set_cell_border(cell, "2F7E9D", 4)
        p = cell.paragraphs[0]
        r1 = p.add_run(label + "  ")
        r1.font.size = Pt(8); r1.font.color.rgb = RGBColor(0xB0, 0xC0, 0xCF); r1.font.bold = False
        r2 = p.add_run(val)
        r2.font.size = Pt(9); r2.font.bold = True; r2.font.color.rgb = RGBColor(0xFF, 0xC0, 0x0C)
        cell.width = Cm(8.5)
    d.add_paragraph().paragraph_format.space_after = Pt(4)


def header_block(d, kicker, title, subtitle):
    p = d.add_paragraph()
    r = p.add_run("SANGGABIZ")
    r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = NAVY
    r2 = p.add_run("   |   Contoh Instrumen untuk Proposal RFP No. 0648/TF-JKT/VIII/2026")
    r2.font.size = Pt(9); r2.font.color.rgb = GREY
    p.paragraph_format.space_after = Pt(2)

    p2 = d.add_paragraph()
    r = p2.add_run(kicker.upper())
    r.font.size = Pt(8.5); r.font.bold = True; r.font.color.rgb = YELLOW
    p2.paragraph_format.space_after = Pt(2)

    p3 = d.add_paragraph()
    r = p3.add_run(title)
    r.font.size = Pt(18); r.font.bold = True; r.font.color.rgb = NAVY
    p3.paragraph_format.space_after = Pt(4)

    p4 = d.add_paragraph()
    r = p4.add_run(subtitle)
    r.font.size = Pt(9.5); r.font.color.rgb = GREY
    p4.paragraph_format.space_after = Pt(8)

    # garis pembatas
    p5 = d.add_paragraph()
    pPr = p5._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "18"); bottom.set(qn("w:color"), "092647")
    pbdr.append(bottom)
    pPr.append(pbdr)
    p5.paragraph_format.space_after = Pt(10)


def note_box(d, text):
    tbl = d.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    shade_cell(cell, "FFF4C2")
    set_cell_border(cell, "C8930A", 6)
    p = cell.paragraphs[0]
    r = p.add_run(text)
    r.font.size = Pt(8.5); r.font.italic = True; r.font.color.rgb = RGBColor(0x5A, 0x42, 0x06)
    cell.margin_top = Cm(0.2); cell.margin_bottom = Cm(0.2)
    d.add_paragraph().paragraph_format.space_after = Pt(4)


def h2(d, text):
    p = d.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = TEAL
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)


def body(d, text, size=9.5, bold=False, italic=False, color=None, space_after=4):
    p = d.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color or RGBColor(0x22, 0x22, 0x22)
    p.paragraph_format.space_after = Pt(space_after)
    return p


def bullets(d, items, size=9.5):
    for it in items:
        p = d.add_paragraph(style="List Bullet")
        r = p.add_run(it)
        r.font.size = Pt(size)
        p.paragraph_format.space_after = Pt(2)


def styled_table(d, headers, rows, col_widths=None, header_fill="092647", font_size=8.8):
    tbl = d.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = True
    for j, htext in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        shade_cell(cell, header_fill)
        set_cell_border(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
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


def footer_note(d, text):
    d.add_paragraph().paragraph_format.space_after = Pt(2)
    p = d.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(7.5); r.font.color.rgb = GREY; r.font.italic = True


# ============================================================
# INSTRUMEN 1: RUBRIK OBSERVASI PPL
# ============================================================
def build_rubrik():
    d = base_doc()
    header_block(
        d, "Contoh Instrumen: Pemetaan Pelaksanaan PPL",
        "Rubrik Observasi Pelaksanaan PPL",
        "Dipakai Koordinator Lapangan bersama Dosen Pembimbing Lapangan (DPL) dan guru pamong saat kunjungan ke LPTK. Enam dimensi mengikuti RFP bagian C.4, yaitu standar, peran, pendampingan, observasi, umpan balik, dan penilaian. Skala 1 sampai 4 selaras dengan rubrik LAMDIK dan BAN-PT yang dipakai di seluruh proposal.",
    )
    phase_chip(d, "Fase 02 Design", "Fase 03 Deliver")
    note_box(d, "STATUS: DRAF/CONTOH, BELUM FINAL. Disusun Sanggabiz berdasar RFP dan TOR untuk menunjukkan kesiapan kerja. Masih memerlukan masukan dan penyempurnaan lebih lanjut dari tim. Bukan instrumen resmi Direktorat PPG, akan diselaraskan dengan Pedoman Pendampingan 2026 saat kick-off.")

    h2(d, "A. Identitas Observasi")
    styled_table(d, ["Field", "Isian"], [
        ["LPTK / Program Studi", "…………………………………………………………"],
        ["Tanggal dan lokasi observasi", "…………………………………………………………"],
        ["Observer (Koordinator Lapangan)", "…………………………………………………………"],
        ["Mahasiswa PPG yang diobservasi", "…………………………………………………………"],
        ["DPL pendamping", "…………………………………………………………"],
        ["Guru pamong pendamping", "…………………………………………………………"],
        ["Tahap PPL", "☐ PPL I (Terbimbing)     ☐ PPL II (Mandiri)"],
    ], col_widths=[6.5, 10.5])

    h2(d, "B. Enam Dimensi Penilaian")
    body(d, "Skor tiap dimensi: 4 berarti melampaui standar, 3 berarti memenuhi, 2 berarti sebagian memenuhi, dan 1 berarti belum memenuhi. Lingkari skor yang sesuai, lalu isi catatan bukti.", italic=True, color=GREY)

    dims = [
        ("1. Standar pelaksanaan",
         "Pedoman PPL diikuti secara konsisten. Durasi, tahapan (PPL I dan II), serta penempatan di sekolah mitra sesuai ketentuan.",
         "Pedoman PPL LPTK, SK penempatan, jadwal PPL"),
        ("2. Pembagian peran",
         "Peran DPL, guru pamong, dan koordinator prodi jelas dan dijalankan. Rasio DPL dengan mahasiswa wajar, dan kualifikasi guru pamong sesuai syarat.",
         "SK penugasan DPL dan guru pamong, daftar rasio bimbingan"),
        ("3. Pendampingan",
         "Frekuensi dan kolaborasi antara DPL dan guru pamong memadai. Siklus supervisi klinis berjalan: pra-observasi, observasi, umpan balik, lalu refleksi.",
         "Jurnal bimbingan, jadwal pendampingan"),
        ("4. Observasi",
         "Observasi kelas memakai instrumen terstandar. Partisipasi aktif dan keterlibatan mahasiswa turut diamati, sesuai Temuan 2 Supervisi 2025.",
         "Lembar observasi terisi, rekaman atau catatan kelas"),
        ("5. Umpan balik",
         "Umpan balik diberikan tepat waktu, spesifik, dan berorientasi perbaikan. Mahasiswa melakukan refleksi diri.",
         "Catatan umpan balik, lembar refleksi mahasiswa"),
        ("6. Penilaian",
         "Rubrik penilaian jelas dan dipakai konsisten. Ada kalibrasi antara nilai DPL dan guru pamong, dan keterkaitan dengan capaian pembelajaran lulusan disebutkan.",
         "Rubrik penilaian, rekap nilai, berita acara kalibrasi"),
    ]
    for title, desc, bukti in dims:
        p = d.add_paragraph()
        r = p.add_run(title); r.font.bold = True; r.font.size = Pt(10); r.font.color.rgb = NAVY
        p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(2)
        body(d, desc, size=9)
        pb = d.add_paragraph()
        r = pb.add_run("Bukti yang dicari: "); r.font.bold = True; r.font.size = Pt(8.5); r.font.color.rgb = GREY
        r2 = pb.add_run(bukti); r2.font.size = Pt(8.5); r2.font.color.rgb = GREY
        pb.paragraph_format.space_after = Pt(2)
        ps = d.add_paragraph()
        r = ps.add_run("Skor:   1   2   3   4        ")
        r.font.bold = True; r.font.size = Pt(9.5); r.font.color.rgb = TEAL
        r2 = ps.add_run("Catatan bukti: ……………………………………………………………………………")
        r2.font.size = Pt(9)
        ps.paragraph_format.space_after = Pt(6)

    h2(d, "C. Ringkasan dan Rekomendasi")
    styled_table(d, ["Dimensi", "Skor (1 sampai 4)", "Prioritas tindak lanjut"], [
        ["1. Standar pelaksanaan", "", ""],
        ["2. Pembagian peran", "", ""],
        ["3. Pendampingan", "", ""],
        ["4. Observasi", "", ""],
        ["5. Umpan balik", "", ""],
        ["6. Penilaian", "", ""],
    ], col_widths=[6, 3, 8])

    footer_note(d, "Rubrik Observasi PPL, versi 2 (draf), 30 September 2026. Disusun Sanggabiz untuk Proposal Teknis RFP No. 0648/TF-JKT/VIII/2026 kepada Tanoto Foundation. Status: draf/contoh, memerlukan masukan dan penyempurnaan lebih lanjut. Skala skor selaras dengan IAPSK 3.0 ISK LAMDIK, 26 Mei 2026.")
    out = os.path.join(HERE, "Rubrik-Observasi-PPL.docx")
    d.save(out)
    return out


# ============================================================
# INSTRUMEN 2: PANDUAN FGD PEMETAAN PPL
# ============================================================
def build_fgd():
    d = base_doc()
    header_block(
        d, "Contoh Instrumen: Pemetaan Pelaksanaan PPL",
        "Panduan FGD Pemetaan PPL",
        "Empat sesi FGD terpisah untuk tiap kelompok informan, yaitu DPL, guru pamong, mahasiswa PPG, dan pimpinan prodi atau UPPS. Durasi 45 sampai 60 menit per sesi, dilaksanakan daring atau saat kunjungan luring. Pertanyaan disusun mengikuti enam dimensi PPL RFP bagian C.4, dan menjawab langsung Temuan 2 dan Temuan 3 Supervisi 2025 tentang pembelajaran berpusat mahasiswa serta mutu praktikum dan umpan balik.",
    )
    phase_chip(d, "Fase 02 Design", "Fase 03 Deliver")
    note_box(d, "STATUS: DRAF/CONTOH, BELUM FINAL. Disusun Sanggabiz berdasar RFP dan TOR untuk menunjukkan kesiapan kerja. Masih memerlukan masukan dan penyempurnaan lebih lanjut dari tim. Bukan instrumen resmi Direktorat PPG, akan diselaraskan dengan Pedoman Pendampingan 2026 saat kick-off.")

    h2(d, "Persiapan Sesi")
    styled_table(d, ["Field", "Isian"], [
        ["LPTK / Program Studi", "…………………………………………………………"],
        ["Kelompok informan", "☐ DPL   ☐ Guru pamong   ☐ Mahasiswa PPG   ☐ Pimpinan prodi/UPPS"],
        ["Tanggal, moda (daring/luring)", "…………………………………………………………"],
        ["Fasilitator / pencatat", "…………………………………………………………"],
        ["Jumlah peserta", "…………………………………………………………"],
        ["Persetujuan rekaman (consent)", "☐ Ya, direkam     ☐ Tidak, catatan tertulis saja"],
    ], col_widths=[6.5, 10.5])

    h2(d, "Pembukaan (5 menit)")
    bullets(d, [
        "Perkenalan fasilitator dan tujuan FGD, yaitu memetakan pelaksanaan PPL untuk mendukung kesiapan Akreditasi Unggul.",
        "Jelaskan kerahasiaan. Hasil dianalisis lintas LPTK dan tidak dikutip per nama tanpa izin. Laporan tidak dipublikasikan tanpa persetujuan tertulis Tanoto Foundation dan Direktorat PPG.",
        "Minta persetujuan rekaman atau catatan.",
    ])

    sections = [
        ("1. Standar pelaksanaan dan pembagian peran (10 menit)", [
            "Bagaimana durasi dan tahapan PPL (PPL I dan II) dijalankan di prodi ini dibandingkan pedoman yang berlaku?",
            "Bagaimana peran DPL, guru pamong, dan koordinator prodi dibagi? Apakah rasio bimbingan terasa memadai?",
            "Apa tantangan utama dalam penetapan sekolah mitra dan penempatan mahasiswa?",
        ]),
        ("2. Pendampingan dan kolaborasi DPL dengan guru pamong (10 menit)", [
            "Seberapa sering DPL dan guru pamong berkoordinasi, dan bagaimana bentuknya? Tatap muka, daring, atau campuran?",
            "Bagaimana siklus pra-observasi, observasi, umpan balik, dan refleksi dijalankan dalam praktiknya?",
            "Apa yang membuat kolaborasi antara DPL dan guru pamong berjalan baik, atau justru terhambat?",
        ]),
        ("3. Pembelajaran berpusat mahasiswa, Temuan 2 Supervisi 2025 (10 menit)", [
            "Sejauh mana mahasiswa PPG didorong berpartisipasi aktif selama praktik mengajar, bukan sekadar mengikuti arahan?",
            "Bagaimana dukungan diberikan sesuai kebutuhan masing-masing mahasiswa, bukan dengan pendekatan yang seragam?",
            "Contoh konkret, kapan terakhir Anda melihat mahasiswa mengambil inisiatif dalam merancang pembelajaran?",
        ]),
        ("4. Observasi, umpan balik, dan penilaian, Temuan 3 Supervisi 2025 (15 menit)", [
            "Instrumen apa yang dipakai untuk mengobservasi praktik mengajar? Apakah dipakai konsisten oleh semua DPL dan guru pamong?",
            "Bagaimana umpan balik diberikan? Seberapa cepat dan sespesifik apa, serta apakah mahasiswa diberi ruang refleksi diri?",
            "Bagaimana penilaian dikalibrasi antara DPL dan guru pamong, supaya hasilnya tidak terlalu berbeda?",
            "Apa satu perubahan yang paling membantu meningkatkan mutu observasi dan umpan balik di prodi ini?",
        ]),
    ]
    for title, qs in sections:
        h2(d, title)
        bullets(d, qs)

    h2(d, "Penutup (5 menit)")
    bullets(d, [
        "Ringkasan singkat oleh fasilitator, lalu konfirmasi ke peserta apakah ada yang perlu diluruskan.",
        "Informasikan langkah selanjutnya, dan cara peserta bisa menghubungi tim jika ada tambahan masukan.",
    ])

    h2(d, "Catatan Analisis (diisi fasilitator setelah sesi)")
    styled_table(d, ["Dimensi", "Ringkasan temuan", "Kutipan atau bukti pendukung"], [
        ["Standar dan peran", "", ""],
        ["Pendampingan", "", ""],
        ["Berpusat mahasiswa", "", ""],
        ["Observasi dan umpan balik", "", ""],
        ["Praktik baik teridentifikasi", "", ""],
        ["Hambatan teridentifikasi", "", ""],
    ], col_widths=[4.5, 6.5, 6])

    footer_note(d, "Panduan FGD Pemetaan PPL, versi 2 (draf), 30 September 2026. Disusun Sanggabiz untuk Proposal Teknis RFP No. 0648/TF-JKT/VIII/2026 kepada Tanoto Foundation. Status: draf/contoh, memerlukan masukan dan penyempurnaan lebih lanjut.")
    out = os.path.join(HERE, "Panduan-FGD-Pemetaan-PPL.docx")
    d.save(out)
    return out


if __name__ == "__main__":
    p1 = build_rubrik()
    p2 = build_fgd()
    print("saved", p1)
    print("saved", p2)
