"""Slide berdiri sendiri: CV Personel Kunci untuk posisi Administrator Data dan Dokumentasi
(Annisa Nastiti). Format sama dengan slide CV Project Manager (folder 8). Isi diambil murni
dari CV asli (CV Annisa Nastiti_Agustus.pdf) — tidak ada kompetensi tambahan yang dikarang.
Untuk diimpor manual ke Google Slides.
"""
import os, sys
from PIL import Image, ImageDraw, ImageFont

BUILD_DIR = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "_build"))
sys.path.insert(0, BUILD_DIR)
from lib import *  # noqa

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_PPTX = os.path.join(HERE, "Slide - CV Administrator Data dan Dokumentasi (Annisa Nastiti).pptx")

INK = "1D2733"
LGREY = "EEF1F5"
TEAL_TINT = "E6F0F4"


def make_avatar(path, initials="AN"):
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

text(s, ML, 0.4, CW - 2.4, 0.3, "KUALITAS DAN KAPABILITAS TIM · CV PERSONEL KUNCI",
     size=11.5, bold=True, color=TEAL, caps=True, tracking=120)
text(s, SW - ML - 2.2, 0.4, 2.2, 0.3, "sanggabiz", size=13, bold=True, color=NAVY_D, align="r", font=F_XB)
text(s, ML, 0.7, CW - 2.4, 0.6, "Administrator Data dan Dokumentasi", size=26, bold=True, color=NAVY_D, font=F_XB, spacing=1.0)

col_gap = 0.35
col1_x = ML
col1_w = 3.75
col2_x = col1_x + col1_w + col_gap
col2_w = CW - col1_w - col_gap

y0 = 1.55

av_size = 1.05
photo(s, avatar_path, col1_x, y0, av_size, av_size, rounded=True)
name_card_x = col1_x + av_size + 0.16
name_card_w = col1_w - av_size - 0.16
card(s, name_card_x, y0, name_card_w, av_size, fill=NAVY_D, line=None, radius=0.1)
text(s, name_card_x + 0.14, y0 + 0.14, name_card_w - 0.28, 0.4, "Annisa Nastiti",
     size=13.5, bold=True, color=WHITE, spacing=0.95)
text(s, name_card_x + 0.14, y0 + 0.14 + 0.4, name_card_w - 0.28, 0.5,
     "Teliti dan berorientasi detail, berpengalaman mengelola arsip dan basis data",
     size=8.8, color=GREY, spacing=1.0)

y1 = y0 + av_size + 0.2
text(s, col1_x, y1, col1_w, 0.28, "PENDIDIKAN", size=10.5, bold=True, color=NAVY_D, caps=True, tracking=60)
text(s, col1_x, y1 + 0.26, col1_w, 0.75, [
    [("S1 ", {"bold": True}), ("- Sastra Indonesia", {})],
    "Universitas Negeri Yogyakarta (2021-2025)",
    "IPK 3,72/4,00 (Dengan Pujian)",
    "Pertukaran Mahasiswa Merdeka, Universitas Udayana (2023)",
], size=9.3, color=INK, spacing=1.1, after=2)

y2 = y1 + 1.15
text(s, col1_x, y2, col1_w, 0.28, "SERTIFIKASI DAN KEAHLIAN", size=10.5, bold=True, color=NAVY_D, caps=True, tracking=60)
text(s, col1_x, y2 + 0.26, col1_w, 0.9,
     "ProTEFL (550/677), Sertifikasi Microsoft Office Desktop Application (88,89/100). "
     "Microsoft Excel (VLOOKUP/HLOOKUP), sistem arsip digital, metadata dan katalogisasi dokumen.",
     size=9.3, color=INK, spacing=1.15)

y3 = y2 + 1.15
hl_h = 1.15
card(s, col1_x, y3, col1_w, hl_h, fill=TEAL_TINT, line=TEAL, lw=1, radius=0.1)
text(s, col1_x + 0.16, y3 + 0.14, col1_w - 0.32, hl_h - 0.28,
     "Berpengalaman langsung pada pengarsipan dan pendokumentasian data, dari naskah kuno "
     "beraksara Jawa hingga sistem arsip digital dan basis data pelanggan.",
     size=9.5, bold=True, italic=True, color=NAVY_D, spacing=1.15)

text(s, col2_x, y0, col2_w, 0.3, "PENGALAMAN KERJA RELEVAN", size=12, bold=True, color=NAVY_D, caps=True, tracking=60)
rect(s, col2_x, y0 + 0.32, col2_w, 0.02, fill=TEAL)

jobs = [
    ("Komunitas Jangkah Nusantara", "Konservatoris Arsip dan Naskah Kuno · Juli-Desember 2024",
     "Mentransliterasi, mendeskripsikan, dan menyusun metadata arsip serta naskah kuno beraksara Jawa dan "
     "Latin dalam beberapa bahasa. Berpartisipasi dalam lokakarya restorasi, preservasi, dan digitalisasi "
     "arsip bersama komunitas dan mitra. Mendapat predikat “Sangat Baik” dalam sertifikat magang."),
    ("LBB Unggulan Education", "Staf Administrasi · Januari-Juni 2024",
     "Mengembangkan sistem arsip digital dengan VLOOKUP/HLOOKUP untuk mempermudah penyortiran dokumen. "
     "Memperbarui dan mereorganisasi database pelanggan untuk meningkatkan aksesibilitas dan akurasi data. "
     "Menginisiasi pembuatan slip gaji, formulir komunikasi kerja, dan kuitansi pembayaran, serta "
     "mengoordinasikan komunikasi antara siswa, orang tua, dan pengajar."),
    ("Nurani Center Solo", "Administrasi Keuangan · Juli 2026-sekarang",
     "Mengelola dan melakukan pelaporan keuangan, termasuk pembuatan invoice, kwitansi, dan dokumen "
     "administrasi keuangan lain secara berkelanjutan."),
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
