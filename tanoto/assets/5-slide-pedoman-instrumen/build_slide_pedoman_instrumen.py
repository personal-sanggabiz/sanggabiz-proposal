"""Slide berdiri sendiri: 'Pedoman dan instrumen' (pasangan kartu Strategi Keberlanjutan
dari 'Repositori dan dashboard'). Menampilkan cuplikan isi + link Google Drive untuk
tiga dokumen di assets/3-instrumen/. Untuk diimpor manual ke Google Slides.
"""
import os, sys
BUILD_DIR = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "_build"))
sys.path.insert(0, BUILD_DIR)
from lib import *  # noqa

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_PPTX = os.path.join(HERE, "Slide - Pedoman dan Instrumen.pptx")

prs = new_prs()
s = blank(prs, "grad")

text(s, ML, 0.42, CW, 0.3, "STRATEGI KEBERLANJUTAN", size=11.5, bold=True, color=YELLOW, caps=True, tracking=120)
text(s, ML, 0.72, CW, 0.7, "Pedoman dan instrumen", size=32, color=WHITE, font=F_XB, spacing=1.0)
text(s, ML, 1.42, CW, 0.35,
     "Rubrik, instrumen observasi PPL, FGD, dan pedoman pendampingan yang disempurnakan, siap dipakai ulang.",
     size=12, color=GREY)

def hyperlink(shape, url):
    r = shape.text_frame.paragraphs[0].runs[0]
    r.hyperlink.address = url


docs = [
    ("Pedoman Pendampingan Teknis", "Metodologi 5 fase, tata kelola, prosedur kerja, indikator RTL, daftar istilah", "prev_pedoman.png",
     "https://drive.google.com/file/d/1ha5-qChCFAWESvPBs_YCIk1wztmZecn7/view"),
    ("Rubrik Observasi PPL", "Enam dimensi penilaian PPL, skala 1 sampai 4", "prev_rubrik.png",
     "https://drive.google.com/file/d/1IoTR5Qby_PAVQTZDJweBhtjcmWByYDcr/view"),
    ("Panduan FGD Pemetaan PPL", "Empat sesi kelompok informan, menjawab Temuan 2 dan 3", "prev_fgd.png",
     "https://drive.google.com/file/d/1QONJRSzPWMIEqa3gu9igzXVSQa41UhnV/view"),
]

n = len(docs)
gap = 0.22
col_w = (CW - gap * (n - 1)) / n
y0 = 1.85
title_h = 0.4
sub_h = 0.5
img_h = 2.85
link_h = 0.38

for i, (title, sub, img, link) in enumerate(docs):
    x = ML + i * (col_w + gap)
    card(s, x, y0, col_w, title_h + sub_h + img_h + link_h + 0.26, fill=NAVY_DD, line=TEAL, lw=1.5, radius=0.14)
    text(s, x + 0.16, y0 + 0.1, col_w - 0.32, title_h, title, size=13.5, bold=True, color=WHITE, spacing=0.95)
    text(s, x + 0.16, y0 + 0.08 + title_h, col_w - 0.32, sub_h, sub, size=9, color=GREY, spacing=1.0)
    iy = y0 + 0.1 + title_h + sub_h
    photo(s, os.path.join(HERE, img), x + 0.14, iy, col_w - 0.28, img_h, focus=(0.5, 0.0), line=WHITE)
    ly = iy + img_h + 0.1
    pill(s, x + 0.14, ly, col_w - 0.28, link_h, fill=NAVY_P, line=YELLOW, lw=1)
    icon(s, "link", x + 0.22, ly + 0.05, 0.26)
    btn = text(s, x + 0.58, ly, col_w - 0.72, link_h, "Buka di Google Drive", size=9.5, bold=True, color=YELLOW, anchor="m")
    hyperlink(btn, link)

# baris link teks penuh di bawah (mudah disalin, dan kelihatan tanpa klik)
ly2 = y0 + title_h + sub_h + img_h + link_h + 0.42
for i, (title, sub, img, link) in enumerate(docs):
    x = ML + i * (col_w + gap)
    lk = text(s, x + 0.14, ly2, col_w - 0.28, 0.38, link, size=7.8, color=TEAL, spacing=0.95)
    hyperlink(lk, link)

logo(s, ML, 7.05, 0.24)
text(s, ML + 1.35, 7.08, 8, 0.25, "Slide berdiri sendiri, dipakai di Google Slides langsung (bukan bagian PPTX/PDF utama)", size=8.5, color=GREY)

prs.save(OUT_PPTX)
print("saved", OUT_PPTX)
