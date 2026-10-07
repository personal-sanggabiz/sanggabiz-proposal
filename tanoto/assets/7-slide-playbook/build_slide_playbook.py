"""Slide berdiri sendiri: 'Replikasi ke LPTK lain' (kartu Strategi Keberlanjutan ke-4).
Menampilkan cuplikan isi Playbook Akreditasi-Ready + link Google Drive. Untuk diimpor
manual ke Google Slides, satu set dengan 'Repositori dan dashboard' dan 'Pedoman dan instrumen'.
"""
import os, sys
BUILD_DIR = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "_build"))
sys.path.insert(0, BUILD_DIR)
from lib import *  # noqa

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_PPTX = os.path.join(HERE, "Slide - Replikasi ke LPTK Lain.pptx")
LINK = "https://drive.google.com/file/d/1Rp2BrCTw9iHn3NjE76ZsKqHxOob5SSna/view"

prs = new_prs()
s = blank(prs, "grad")


def hyperlink(shape, url):
    r = shape.text_frame.paragraphs[0].runs[0]
    r.hyperlink.address = url


text(s, ML, 0.42, CW, 0.3, "STRATEGI KEBERLANJUTAN", size=11.5, bold=True, color=YELLOW, caps=True, tracking=120)
text(s, ML, 0.72, CW, 0.7, "Replikasi ke LPTK lain", size=32, color=WHITE, font=F_XB, spacing=1.0)
text(s, ML, 1.42, CW, 0.35,
     "Playbook Akreditasi-Ready: panduan ringkas dan rekomendasi tahapan untuk 172 LPTK penyelenggara PPG.",
     size=12, color=GREY)

panels = [
    ("Tentang dan 5 Langkah Menuju Akreditasi Unggul", "prev_p2.png"),
    ("Checklist Kesiapan Akreditasi Unggul", "prev_p3.png"),
]
gap = 0.3
col_w = (CW - gap) / 2
y0 = 1.85
title_h = 0.38
img_h = 3.55

for i, (title, img) in enumerate(panels):
    x = ML + i * (col_w + gap)
    card(s, x, y0, col_w, title_h + img_h + 0.2, fill=NAVY_DD, line=TEAL, lw=1.5, radius=0.14)
    text(s, x + 0.18, y0 + 0.1, col_w - 0.36, title_h, title, size=13, bold=True, color=WHITE, spacing=0.95)
    photo(s, os.path.join(HERE, img), x + 0.14, y0 + title_h + 0.06, col_w - 0.28, img_h, focus=(0.5, 0.0), line=WHITE)

card_bottom = y0 + title_h + img_h + 0.2

# baris fakta + tombol dalam satu baris
by = card_bottom + 0.14
facts = ["4 halaman", "contoh ringkas", "untuk 172 LPTK"]
fx = ML
for f in facts:
    w = 0.3 + 0.095 * len(f)
    pill(s, fx, by, w, 0.32, fill=NAVY_P, line=TEAL, lw=1)
    text(s, fx, by, w, 0.32, f, size=9.5, bold=True, color=WHITE, align="c", anchor="m")
    fx += w + 0.12

bw = 3.6
bx = SW - ML - bw
pill(s, bx, by, bw, 0.32, fill=NAVY_P, line=YELLOW, lw=1.25)
icon(s, "link", bx + 0.12, by + 0.03, 0.24)
btn = text(s, bx + 0.44, by, bw - 0.5, 0.32, "Buka di Google Drive", size=9.5, bold=True, color=YELLOW, anchor="m")
hyperlink(btn, LINK)

ly2 = by + 0.42
lk = text(s, ML, ly2, CW, 0.28, LINK, size=8.5, color=TEAL, spacing=0.9)
hyperlink(lk, LINK)

logo(s, ML, 7.05, 0.24)
text(s, ML + 1.35, 7.08, 8, 0.25, "Slide berdiri sendiri, dipakai di Google Slides langsung (bukan bagian PPTX/PDF utama)", size=8.5, color=GREY)

prs.save(OUT_PPTX)
print("saved", OUT_PPTX)
