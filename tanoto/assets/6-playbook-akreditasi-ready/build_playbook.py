"""Bangun Playbook Akreditasi-Ready (DOCX -> PDF): contoh ringkas, menunjukkan struktur saja.
Beda tujuan dari Pedoman Pendampingan Teknis (SOP internal tim untuk 12 LPTK):
Playbook ini adalah produk diseminasi untuk 172 LPTK secara nasional, hasil Fase 05
Evaluate & Improve, dijanjikan di kartu "Replikasi ke LPTK lain" (Strategi Keberlanjutan).
CONTOH, sengaja singkat, hanya menunjukkan struktur yang akan dipakai versi lengkap.
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


def cover(d):
    for _ in range(3):
        d.add_paragraph()
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("SANGGABIZ"); r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = NAVY

    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("CONTOH PLAYBOOK, VERSI RINGKAS"); r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = GOLD
    p.paragraph_format.space_before = Pt(12)

    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Playbook Akreditasi-Ready"); r.font.size = Pt(30); r.font.bold = True; r.font.color.rgb = NAVY
    p.paragraph_format.space_before = Pt(6)

    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Panduan Ringkas Mandiri untuk LPTK Penyelenggara PPG")
    r.font.size = Pt(15); r.font.color.rgb = TEAL
    p.paragraph_format.space_after = Pt(16)

    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Disarikan dari pendampingan teknis 12 LPTK bersama Tanoto Foundation, 2026, untuk dipakai 172 LPTK penyelenggara PPG secara nasional")
    r.font.size = Pt(11); r.font.italic = True; r.font.color.rgb = GREY
    p.paragraph_format.space_after = Pt(40)

    tbl = d.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    shade_cell(cell, "FFF4C2"); set_cell_border(cell, "C8930A", 6); cell.width = Cm(15)
    p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("STATUS: DRAF/CONTOH RINGKAS, BELUM FINAL. Versi ini sengaja dipersingkat untuk menunjukkan struktur saja, bukan playbook lengkap. Disusun Sanggabiz untuk Proposal Teknis RFP No. 0648/TF-JKT/VIII/2026. Versi lengkap disusun pada Fase 05 Evaluate & Improve, setelah pembelajaran dari 12 LPTK terkumpul, dan diselaraskan dengan Direktorat PPG serta Tanoto Foundation sebelum disebarluaskan.")
    r.font.size = Pt(9.5); r.font.italic = True; r.font.color.rgb = RGBColor(0x5A, 0x42, 0x06)
    cell.margin_top = Cm(0.3); cell.margin_bottom = Cm(0.3); cell.margin_left = Cm(0.3); cell.margin_right = Cm(0.3)

    for _ in range(5):
        d.add_paragraph()
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Versi 1 (contoh ringkas), 30 September 2026"); r.font.size = Pt(9.5); r.font.color.rgb = GREY


def running_header(d):
    p = d.add_paragraph()
    r = p.add_run("SANGGABIZ"); r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = NAVY
    r2 = p.add_run("   |   Playbook Akreditasi-Ready, contoh ringkas, belum final")
    r2.font.size = Pt(8); r2.font.color.rgb = GREY
    p.paragraph_format.space_after = Pt(10)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "10"); bottom.set(qn("w:color"), "B0C0CF")
    pbdr.append(bottom)
    pPr.append(pbdr)


def h1(d, num, text_):
    p = d.add_paragraph()
    r = p.add_run(f"{num}.  {text_}"); r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = NAVY
    p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "14"); bottom.set(qn("w:color"), "FFC00C")
    pbdr.append(bottom)
    pPr.append(pbdr)


def h2(d, text_):
    p = d.add_paragraph()
    r = p.add_run(text_); r.font.size = Pt(11.5); r.font.bold = True; r.font.color.rgb = TEAL
    p.paragraph_format.space_before = Pt(10); p.paragraph_format.space_after = Pt(4)


def body(d, text_, size=10, italic=False, color=None, space_after=6, bold=False):
    p = d.add_paragraph()
    r = p.add_run(text_); r.font.size = Pt(size); r.font.italic = italic; r.font.bold = bold
    r.font.color.rgb = color or INK
    p.paragraph_format.space_after = Pt(space_after)
    return p


def bullets(d, items, size=10):
    for it in items:
        p = d.add_paragraph(style="List Bullet")
        r = p.add_run(it); r.font.size = Pt(size)
        p.paragraph_format.space_after = Pt(3)


def numbered(d, items, size=10):
    # nomor ditulis manual (bukan style "List Number" Word) supaya tiap pemanggilan
    # selalu mulai dari 1, tidak melanjutkan hitungan dari daftar sebelumnya di dokumen yang sama.
    for i, it in enumerate(items, start=1):
        p = d.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.first_line_indent = Cm(-0.6)
        r = p.add_run(f"{i}.  "); r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = NAVY
        r2 = p.add_run(it); r2.font.size = Pt(size)
        p.paragraph_format.space_after = Pt(3)


def styled_table(d, headers, rows, col_widths=None, header_fill="092647", font_size=8.7):
    tbl = d.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = True
    for j, htext in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        shade_cell(cell, header_fill); set_cell_border(cell)
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


def note_box(d, text_, fill="FFF4C2", border="C8930A", tcolor=RGBColor(0x5A, 0x42, 0x06)):
    tbl = d.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    shade_cell(cell, fill); set_cell_border(cell, border, 6)
    p = cell.paragraphs[0]
    r = p.add_run(text_); r.font.size = Pt(9); r.font.italic = True; r.font.color.rgb = tcolor
    cell.margin_top = Cm(0.2); cell.margin_bottom = Cm(0.2); cell.margin_left = Cm(0.15); cell.margin_right = Cm(0.15)
    d.add_paragraph().paragraph_format.space_after = Pt(4)


def footer_note(d, text_):
    d.add_paragraph().paragraph_format.space_after = Pt(2)
    p = d.add_paragraph()
    r = p.add_run(text_); r.font.size = Pt(7.5); r.font.color.rgb = GREY; r.font.italic = True


def build():
    d = base_doc()
    cover(d)
    d.add_page_break()

    # ============ 1. TENTANG PLAYBOOK ============
    running_header(d)
    h1(d, 1, "Tentang Playbook Ini")
    body(d, "Playbook ini disiapkan untuk seluruh LPTK penyelenggara PPG di Indonesia, bukan hanya 12 LPTK yang didampingi langsung dalam program ini. Isinya disarikan dari pembelajaran pendampingan teknis 12 LPTK bersama Tanoto Foundation sepanjang Oktober sampai Desember 2026, mengikuti metodologi Evidence-Based Continuous Quality Improvement.")
    h2(d, "1.1 Cara Memakai Playbook Ini")
    numbered(d, [
        "Isi Checklist Kesiapan Akreditasi Unggul pada bagian 2, sebagai asesmen mandiri program studi.",
        "Diskusikan hasilnya bersama pimpinan Unit Pengelola Program Studi dan gugus penjaminan mutu.",
        "Susun rencana tindak lanjut sendiri untuk celah yang ditemukan, mengikuti format sederhana pada bagian 2.",
        "Hubungi Direktorat PPG apabila memerlukan pendampingan teknis resmi lebih lanjut.",
    ])
    h2(d, "1.2 Lima Langkah Menuju Akreditasi Unggul")
    styled_table(d, ["Langkah", "Fokus"], [
        ["1. Kualitas dosen", "Kualifikasi dan jabatan akademik dosen tetap program studi"],
        ["2. Penguatan OBE", "Keselarasan tujuan, desain perkuliahan, dan asesmen capaian pembelajaran"],
        ["3. Optimalisasi PPL", "Standar pelaksanaan, peran, pendampingan, dan umpan balik praktik lapangan"],
        ["4. Penjaminan mutu", "Sistem Penjaminan Mutu Internal berjalan penuh, siklus PPEPP"],
        ["5. Karya dan publikasi", "Potensi publikasi dan karya inovatif dosen, mahasiswa, guru pamong"],
    ], col_widths=[5, 12])

    # ============ 2. CHECKLIST ============
    d.add_page_break()
    running_header(d)
    h1(d, 2, "Checklist Kesiapan Akreditasi Unggul")
    body(d, "Jawab tiap pertanyaan dengan jujur. Beri tanda pada kolom yang sesuai kondisi program studi saat ini.", italic=True, color=GREY)
    styled_table(d, ["Langkah", "Pertanyaan diagnostik mandiri", "Syarat minimal Unggul", "Sudah / Belum"], [
        ["1. Dosen", "Berapa persen dosen tetap bergelar doktor, dan berapa yang berjabatan lektor kepala ke atas?", "Sebagian besar doktor, minimal tiga lektor kepala", "☐ Sudah   ☐ Belum"],
        ["2. OBE", "Apakah capaian pembelajaran lulusan dinilai, dievaluasi, dan ditindaklanjuti untuk mata kuliah penciri?", "Minimal 15% mata kuliah penciri, dengan bukti sahih", "☐ Sudah   ☐ Belum"],
        ["3. PPL", "Apakah ada laboratorium microteaching yang layak, dan mahasiswa praktik lebih dari sekali dengan refleksi?", "Minimal dua kali praktik per mahasiswa, delapan keterampilan dasar", "☐ Sudah   ☐ Belum"],
        ["4. SPMI", "Apakah lima tahap penjaminan mutu internal dijalankan penuh, dari menetapkan sampai meningkatkan standar?", "Lima tahap PPEPP dijalankan, idealnya berbasis sistem informasi", "☐ Sudah   ☐ Belum"],
        ["5. Publikasi", "Berapa persen dosen tetap punya publikasi terindeks dalam tiga tahun terakhir?", "Minimal seperlima dosen tetap, di jurnal terakreditasi atau bereputasi", "☐ Sudah   ☐ Belum"],
    ], col_widths=[2.7, 7.3, 5, 2.5], font_size=8.3)
    note_box(d, "Elemen yang dijawab Belum adalah prioritas rencana tindak lanjut program studi. Catat rencana, penanggung jawab, dan target waktu untuk tiap elemen yang belum terpenuhi.")

    # ============ 3. PELAJARAN & LANGKAH SELANJUTNYA ============
    d.add_page_break()
    running_header(d)
    h1(d, 3, "Pelajaran dari Program Pendampingan 2026")
    body(d, "Pola umum yang ditemukan pada pendampingan 12 LPTK tahun ini, dan berlaku luas bagi LPTK lain yang berada pada tahap serupa.")
    bullets(d, [
        "Implementasi kurikulum berbasis capaian pembelajaran sering kali baru sebatas dokumen, belum diuji dengan data hasil belajar mahasiswa yang sebenarnya.",
        "Kolaborasi antara dosen pembimbing lapangan dan guru pamong paling efektif ketika ada jadwal koordinasi rutin, bukan hanya saat observasi berlangsung.",
        "Penjaminan mutu internal lebih sering berhenti di tahap evaluasi. Tahap pengendalian dan peningkatan jarang terdokumentasi dengan bukti yang jelas.",
        "Program studi yang menunjuk penanggung jawab dan tenggat yang jelas untuk tiap rencana tindak lanjut, jauh lebih cepat menuntaskannya.",
    ])
    h1(d, 4, "Langkah Selanjutnya")
    numbered(d, [
        "Simpan hasil checklist ini sebagai baseline internal program studi.",
        "Susun rencana tindak lanjut untuk maksimal tiga prioritas tertinggi terlebih dahulu, jangan mencoba menuntaskan semua elemen sekaligus.",
        "Tinjau ulang checklist ini setiap semester untuk memantau kemajuan.",
        "Pelajari Pedoman Pendampingan Teknis dan dua instrumen pendukungnya, Rubrik Observasi PPL dan Panduan FGD Pemetaan PPL, sebagai referensi metode yang lebih rinci.",
    ])

    h2(d, "Referensi")
    bullets(d, [
        "Permendikbudristek 53/2023 tentang Penjaminan Mutu Pendidikan Tinggi.",
        "IAPSK 3.0 Instrumen Suplemen Konversi, LAMDIK, 26 Mei 2026.",
        "Pedoman Pendampingan 2026, Direktorat PPG (rujukan utama begitu diterbitkan resmi).",
    ])

    footer_note(d, "Playbook Akreditasi-Ready, versi 1, contoh ringkas, 30 September 2026. Disusun Sanggabiz untuk Proposal Teknis RFP No. 0648/TF-JKT/VIII/2026 kepada Tanoto Foundation. Status: draf/contoh, memerlukan masukan dan penyempurnaan lebih lanjut. Versi lengkap disusun pada Fase 05 Evaluate & Improve program pendampingan.")

    out = os.path.join(HERE, "Playbook-Akreditasi-Ready.docx")
    d.save(out)
    return out


if __name__ == "__main__":
    p = build()
    print("saved", p)
