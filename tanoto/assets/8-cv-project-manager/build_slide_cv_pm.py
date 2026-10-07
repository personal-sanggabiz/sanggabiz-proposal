"""Slide berdiri sendiri: CV Personel Kunci untuk posisi Project Manager dan Koordinator
Nasional (Muhammad Muchson Attoyib). Format mengikuti gaya 'Konsultan Kami' yang sudah ada
di v2 (latar putih, kartu navy, kotak pengalaman kerja bersusun), tapi isinya dipilih hanya
yang relevan untuk peran ini di RFP Tanoto PPG. Untuk diimpor manual ke Google Slides.
"""
import os, sys
from PIL import Image, ImageDraw, ImageFont

BUILD_DIR = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "_build"))
sys.path.insert(0, BUILD_DIR)
from lib import *  # noqa

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_PPTX = os.path.join(HERE, "Slide - CV Project Manager (Muchson Attoyib).pptx")

INK = "1D2733"
LGREY = "EEF1F5"
TEAL_TINT = "E6F0F4"

# ---- avatar placeholder (inisial, karena belum ada foto) ----
def make_avatar(path, initials="MMA"):
    S = 4
    size = 256 * S
    im = Image.new("RGB", (size, size), tuple(int(NAVY_D[i:i + 2], 16) for i in (0, 2, 4)))
    d = ImageDraw.Draw(im)
    try:
        font = ImageFont.truetype("arialbd.ttf", int(size * 0.34))
    except Exception:
        font = ImageFont.load_default()
    bbox = d.textbbox((0, 0), initials, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    d.text(((size - tw) / 2 - bbox[0], (size - th) / 2 - bbox[1]), initials, font=font,
           fill=tuple(int(YELLOW[i:i + 2], 16) for i in (0, 2, 4)))
    im = im.resize((256, 256), Image.LANCZOS)
    im.save(path, quality=92)
    return path


avatar_path = os.path.join(HERE, "avatar_placeholder.jpg")
make_avatar(avatar_path)

prs = new_prs()
s = blank(prs, "white")

# ---- header ----
text(s, ML, 0.4, CW - 2.4, 0.3, "KUALITAS DAN KAPABILITAS TIM · CV PERSONEL KUNCI",
     size=11.5, bold=True, color=TEAL, caps=True, tracking=120)
text(s, SW - ML - 2.2, 0.4, 2.2, 0.3, "sanggabiz", size=13, bold=True, color=NAVY_D, align="r", font=F_XB)
text(s, ML, 0.7, CW - 2.4, 0.6, "Project Manager dan Koordinator Nasional", size=26, bold=True, color=NAVY_D, font=F_XB, spacing=1.0)

col_gap = 0.35
col1_x = ML
col1_w = 3.75
col2_x = col1_x + col1_w + col_gap
col2_w = CW - col1_w - col_gap

y0 = 1.55

# ---- kolom kiri: foto + kartu nama ----
av_size = 1.05
photo(s, avatar_path, col1_x, y0, av_size, av_size, rounded=True)
name_card_x = col1_x + av_size + 0.16
name_card_w = col1_w - av_size - 0.16
card(s, name_card_x, y0, name_card_w, av_size, fill=NAVY_D, line=None, radius=0.1)
text(s, name_card_x + 0.14, y0 + 0.14, name_card_w - 0.28, 0.4, "Muhammad Muchson Attoyib",
     size=13.5, bold=True, color=WHITE, spacing=0.95)
text(s, name_card_x + 0.14, y0 + 0.14 + 0.4, name_card_w - 0.28, 0.5,
     "5+ tahun manajemen program pendidikan dan teknologi, lintas lembaga nasional",
     size=8.8, color=GREY, spacing=1.0)

y1 = y0 + av_size + 0.2
text(s, col1_x, y1, col1_w, 0.28, "PENDIDIKAN", size=10.5, bold=True, color=NAVY_D, caps=True, tracking=60)
text(s, col1_x, y1 + 0.26, col1_w, 0.75, [
    [("S2 ", {"bold": True}), ("- Pendidikan Kejuruan (Vocational Education)", {})],
    "Universitas Negeri Malang (2018-2020)",
    [("S1 ", {"bold": True}), ("- Pendidikan Informatika", {})],
    "Universitas Negeri Malang (2012-2017)",
], size=9.3, color=INK, spacing=1.12, after=2)

y2 = y1 + 1.05
text(s, col1_x, y2, col1_w, 0.28, "KERANGKA KERJA DAN KEAHLIAN", size=10.5, bold=True, color=NAVY_D, caps=True, tracking=60)
text(s, col1_x, y2 + 0.26, col1_w, 0.9,
     "ADDIE, Bloom's Taxonomy, dan Kirkpatrick Evaluation Model untuk desain dan evaluasi program. "
     "Koordinasi pemangku kepentingan lintas lembaga, manajemen linimasa proyek, dan administrasi LMS.",
     size=9.3, color=INK, spacing=1.15)

y3 = y2 + 1.15
hl_h = 1.15
card(s, col1_x, y3, col1_w, hl_h, fill=TEAL_TINT, line=TEAL, lw=1, radius=0.1)
text(s, col1_x + 0.16, y3 + 0.14, col1_w - 0.32, hl_h - 0.28,
     "Pernah mengelola proyek pendidikan bersama Tanoto Foundation, Kampus Merdeka, Prakerja, "
     "UN Women, UNHCR, dan Bappenas sebagai Project Coordinator di Alterra Academy (2022-2024).",
     size=9.5, bold=True, italic=True, color=NAVY_D, spacing=1.15)

# ---- kolom kanan: pengalaman kerja relevan ----
text(s, col2_x, y0, col2_w, 0.3, "PENGALAMAN KERJA RELEVAN", size=12, bold=True, color=NAVY_D, caps=True, tracking=60)
rect(s, col2_x, y0 + 0.32, col2_w, 0.02, fill=TEAL)

jobs = [
    ("Artificial Intelligence Talent Factory - Komdigi", "AI Project Manager · Sep 2025-Mei 2026, Malang (Hybrid)",
     "Memimpin program pengembangan talenta AI untuk 150 peserta sebagai penghubung utama antara dosen "
     "universitas, Komdigi, dan PMO. Merancang alur kerja manajemen proyek: pemetaan tugas, penetapan aturan, "
     "dan pelacakan progres mingguan. Menyelenggarakan workshop dan pertemuan pemangku kepentingan skala besar, "
     "dihadiri Wakil Menteri, Kepala Pusat Komdigi, dan Rektor Universitas Brawijaya. Mencapai tingkat "
     "penyelesaian modul 100% di tiga jalur pembelajaran melalui pemantauan terstruktur."),
    ("Alterra Academy", "Project Coordinator · Juli 2022-Desember 2024, Remote",
     "Mengelola proyek pendidikan skala besar bersama Kampus Merdeka, Prakerja, UN Women, UNHCR, Tanoto "
     "Foundation, dan Bappenas. Mengelola LMS Kampus Merdeka: unggah konten, manajemen akun, pelaporan "
     "analitik progres. Merencanakan linimasa proyek dan alokasi sumber daya, menerapkan sistem pemantauan "
     "hasil belajar. Mengevaluasi efektivitas pelatihan dengan Kirkpatrick Framework (NPS, CSAT, tingkat "
     "putus, kehadiran) dan merancang sistem observasi sejawat untuk menjaga mutu pengajaran."),
    ("Digital Skola", "Curriculum Developer · April 2025-Mei 2026, Remote",
     "Menyusun modul kurikulum dan jalur belajar bersama tim produk dan instruktur sesuai kebutuhan industri. "
     "Mengelola linimasa pengembangan program agar tiap tahap selesai tepat waktu, mengoordinasikan rapat tim "
     "lintas fungsi, dan mereviu konten untuk penjaminan mutu lintas tim desain, teknologi, dan pemasaran."),
]

jy = y0 + 0.48
job_h = 1.42
gap = 0.12
for i, (org, role, desc) in enumerate(jobs):
    fill = LGREY if i % 2 == 0 else WHITE
    card(s, col2_x, jy, col2_w, job_h, fill=fill, line=None, radius=0.06)
    text(s, col2_x + 0.18, jy + 0.1, col2_w - 0.36, 0.3, org, size=11, bold=True, color=NAVY_D, spacing=1.0)
    text(s, col2_x + 0.18, jy + 0.1 + 0.24, col2_w - 0.36, 0.24, role, size=9, bold=True, color=TEAL, spacing=1.0)
    text(s, col2_x + 0.18, jy + 0.1 + 0.48, col2_w - 0.36, job_h - 0.6, desc, size=8.8, color=INK, spacing=1.12)
    jy += job_h + gap

# ---- footer ----
text(s, ML + 1.6, 7.12, 8, 0.25, "CV lengkap dilampirkan sesuai RFP Lampiran 1 poin 7 (Deskripsi Organisasi dan Staf)",
     size=8.5, color="7A8899")
text(s, SW - ML - 1.0, 7.12, 1.0, 0.25, "22", size=10, bold=True, color=TEAL, align="r")

prs.save(OUT_PPTX)
print("saved", OUT_PPTX)
