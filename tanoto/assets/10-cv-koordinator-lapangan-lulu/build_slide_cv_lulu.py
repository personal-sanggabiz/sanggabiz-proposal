"""Slide berdiri sendiri: CV Personel Kunci untuk posisi Koordinator Lapangan
(Lu'lu' Al Adhliyani). Format sama dengan slide CV Project Manager (folder 8) dan
Administrator Data (folder 9). Isi diambil dari bio yang sudah ada di v2 (slide TIM |
Konsultan) — tidak ada kompetensi tambahan yang dikarang. Foto dipotong dari slide asli.
Untuk diimpor manual ke Google Slides.
"""
import os, sys

BUILD_DIR = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "_build"))
sys.path.insert(0, BUILD_DIR)
from lib import *  # noqa

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_PPTX = os.path.join(HERE, "Slide - CV Koordinator Lapangan (Lulu Al Adhliyani).pptx")
PHOTO = os.path.join(HERE, "lulu_photo.jpg")

INK = "1D2733"
LGREY = "EEF1F5"
TEAL_TINT = "E6F0F4"

prs = new_prs()
s = blank(prs, "white")

text(s, ML, 0.4, CW - 2.4, 0.3, "KUALITAS DAN KAPABILITAS TIM · CV PERSONEL KUNCI",
     size=11.5, bold=True, color=TEAL, caps=True, tracking=120)
text(s, SW - ML - 2.2, 0.4, 2.2, 0.3, "sanggabiz", size=13, bold=True, color=NAVY_D, align="r", font=F_XB)
text(s, ML, 0.7, CW - 2.4, 0.6, "Koordinator Lapangan", size=26, bold=True, color=NAVY_D, font=F_XB, spacing=1.0)

col_gap = 0.35
col1_x = ML
col1_w = 3.75
col2_x = col1_x + col1_w + col_gap
col2_w = CW - col1_w - col_gap

y0 = 1.55

av_size = 1.05
name_card_h = 1.35
photo(s, PHOTO, col1_x, y0, av_size, av_size, rounded=True, focus=(0.5, 0.15))
name_card_x = col1_x + av_size + 0.16
name_card_w = col1_w - av_size - 0.16
card(s, name_card_x, y0, name_card_w, name_card_h, fill=NAVY_D, line=None, radius=0.1)
text(s, name_card_x + 0.14, y0 + 0.14, name_card_w - 0.28, 0.4, "Lu'lu' Al Adhliyani",
     size=13.5, bold=True, color=WHITE, spacing=0.95)
text(s, name_card_x + 0.14, y0 + 0.14 + 0.4, name_card_w - 0.28, 0.8,
     "Saat ini Academic Product Manager Trainee di Sanggabiz; berpengalaman koordinasi lapangan dan acara skala besar",
     size=8.6, color=GREY, spacing=1.05)

y1 = y0 + name_card_h + 0.2
text(s, col1_x, y1, col1_w, 0.28, "PENDIDIKAN", size=10.5, bold=True, color=NAVY_D, caps=True, tracking=60)
text(s, col1_x, y1 + 0.26, col1_w, 0.55, [
    [("S1 ", {"bold": True}), ("- Ekonomi Pertanian (Sarjana Pertanian, S.P.)", {})],
    "Universitas Gadjah Mada",
], size=9.3, color=INK, spacing=1.1, after=2)

y2 = y1 + 0.9
text(s, col1_x, y2, col1_w, 0.28, "KEAHLIAN", size=10.5, bold=True, color=NAVY_D, caps=True, tracking=60)
text(s, col1_x, y2 + 0.26, col1_w, 1.0,
     "Training and Development (HR), Product Management, Event Management. Perancangan kurikulum "
     "(learning journey), kemitraan pendidikan, dan pengelolaan acara berskala besar.",
     size=9.3, color=INK, spacing=1.15)

y3 = y2 + 1.25
hl_h = 1.15
card(s, col1_x, y3, col1_w, hl_h, fill=TEAL_TINT, line=TEAL, lw=1, radius=0.1)
text(s, col1_x + 0.16, y3 + 0.14, col1_w - 0.32, hl_h - 0.28,
     "Mengoordinasikan logistik dan lapangan untuk acara 3.000+ peserta (AgriFest, pendanaan "
     "Rp250 juta) dan memfasilitasi 251 mahasiswa dari 61 universitas untuk program MBKM.",
     size=9.5, bold=True, italic=True, color=NAVY_D, spacing=1.15)

text(s, col2_x, y0, col2_w, 0.3, "PENGALAMAN KERJA RELEVAN", size=12, bold=True, color=NAVY_D, caps=True, tracking=60)
rect(s, col2_x, y0 + 0.32, col2_w, 0.02, fill=TEAL)

jobs = [
    ("Pusat Inovasi Agroteknologi (PIAT) UGM", "Field Coordinator & Organizing Committee",
     "Mengoordinasikan kegiatan lapangan dan logistik acara skala besar, termasuk AgriFest dengan "
     "3.000+ peserta dan pendanaan Rp250.000.000. Pengalaman langsung sebagai koordinator lapangan "
     "untuk acara multi-hari dengan banyak pemangku kepentingan."),
    ("Universitas Gadjah Mada", "Panitia MBKM dan Tim Teknis",
     "Memfasilitasi 251 mahasiswa dari 61 universitas untuk program Merdeka Belajar Kampus Merdeka "
     "(MBKM), mengoordinasikan penempatan, administrasi, dan komunikasi lintas institusi pendidikan "
     "tinggi."),
    ("Skill Academy by Ruangguru & Dompet Dhuafa", "Product Management dan Kemitraan Program",
     "Pengalaman product management pada program pelatihan skala nasional serta kerja sama program "
     "sosial-pendidikan. Mengoordinasikan program magang untuk ratusan siswa dan mahasiswa."),
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

text(s, ML + 1.6, 7.12, 8, 0.25, "CV lengkap dilampirkan sesuai RFP Lampiran 1 poin 7 (Deskripsi Organisasi dan Staf)",
     size=8.5, color="7A8899")
text(s, SW - ML - 1.0, 7.12, 1.0, 0.25, "22", size=10, bold=True, color=TEAL, align="r")

prs.save(OUT_PPTX)
print("saved", OUT_PPTX)
