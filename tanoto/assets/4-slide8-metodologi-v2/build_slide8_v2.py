"""Slide 8 versi 2 — berdiri sendiri (BUKAN bagian build_deck.py/PPTX utama).
Metodologi: Evidence-Based Continuous Quality Improvement (5 fase) + Key Deliverables.
Dipakai user untuk diimpor manual ke Google Slides mereka. Deck lokal (Proposal Teknis...pptx)
slide 8-nya TETAP versi PDCA/PPEPP — jangan disamakan otomatis, ini varian terpisah yang disengaja.
"""
import os, sys
BUILD_DIR = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "_build"))
sys.path.insert(0, BUILD_DIR)
from lib import *  # noqa

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_PPTX = os.path.join(OUT_DIR, "Slide 8 v2 - Metodologi (Evidence-Based CQI).pptx")

prs = new_prs()
s = blank(prs, "grad")

header(s, "Pendekatan pendampingan", "Evidence-Based Continuous Quality Improvement", tsize=27)
text(s, ML, 1.28, CW, 0.4,
     "Lima fase, mengikuti langsung ruang lingkup mitra pelaksana RFP bagian C.4.",
     size=12, color=GREY)

# ---------- 5 fase (ribbon) ----------
phases = [
    ("01", "Diagnose", "Okt", TEAL, WHITE,
     ["Review hasil Supervisi 2025 & baseline assessment", "Pemetaan kesenjangan mutu dan kesiapan akreditasi"]),
    ("02", "Design", "Okt", NAVY_DD, WHITE,
     ["Penyusunan pedoman, modul, dan instrumen", "Rencana pendampingan berbasis kebutuhan LPTK"]),
    ("03", "Deliver", "Okt–Nov", TEAL, WHITE,
     ["Pendampingan luring di 12 LPTK", "Workshop, FGD, technical review, action planning"]),
    ("04", "Monitor & Verify", "Okt–Des", NAVY_DD, WHITE,
     ["Monitoring RTL & konsultasi berkala", "Verifikasi bukti perbaikan"]),
    ("05", "Evaluate & Improve", "Nov–Des", YELLOW, NAVY_D,
     ["Studi komparatif & analisis pelaksanaan PPL", "Sintesis praktik baik, hambatan, rekomendasi"]),
]
n = len(phases)
gap = 0.16
card_w = (CW - gap * (n - 1)) / n
y0 = 1.65
num_h = 0.5
card_h = 1.55

for i, (num, name, when, fc, tc, bl) in enumerate(phases):
    x = ML + i * (card_w + gap)
    text(s, x, y0, card_w, num_h, num, size=28, font=F_XB, color=YELLOW, align="l")
    if i < n - 1:
        text(s, x + card_w - 0.32, y0, 0.5, num_h, "→", size=16, bold=True, color=YELLOW, align="c", anchor="m")
    card(s, x, y0 + num_h, card_w, card_h, fill=fc, line=WHITE if fc != YELLOW else NAVY_D, lw=1.75, radius=0.14)
    text(s, x + 0.15, y0 + num_h + 0.08, card_w - 0.3, 0.35, name, size=12.5, bold=True, color=tc, spacing=0.9)
    text(s, x + 0.15, y0 + num_h + 0.42, card_w - 0.3, 0.2, when, size=9, bold=True, color=tc)
    text(s, x + 0.15, y0 + num_h + 0.64, card_w - 0.28, card_h - 0.66, bl, size=8.6, bullet="•", after=4, color=tc, spacing=1.0)

# ---------- grounding note ----------
gy = y0 + num_h + card_h + 0.14
card(s, ML, gy, CW, 0.56, fill=NAVY_DD, line=YELLOW, lw=1.25, radius=0.13)
text(s, ML + 0.2, gy, CW - 0.4, 0.56,
     [[("Dasar pendekatan: ", {"bold": True, "color": YELLOW, "size": 9}),
       ("Continuous Quality Improvement (akar PDCA/Deming, dasar filosofi mutu banyak sistem akreditasi pendidikan tinggi termasuk PPEPP/SPMI) dipadukan dengan evidence-based practice — rekomendasi berbasis bukti yang dapat ditelusuri, bukan asumsi.", {"size": 9})]],
     anchor="m")

# ---------- Key Deliverables (9) ----------
dy = gy + 0.68
text(s, ML, dy, 4, 0.26, "KEY DELIVERABLES", size=10.5, bold=True, color=YELLOW, tracking=100)
deliverables = [
    "Pedoman dan modul pendampingan teknis",
    "Instrumen dan dashboard monitoring",
    "Baseline assessment 12 LPTK",
    "Paket perbaikan akreditasi 12 LPTK",
    "Laporan pendampingan 12 LPTK",
    "Laporan kemajuan RTL Supervisi 2025",
    "Studi komparatif pendampingan",
    "Analisis pelaksanaan PPL",
    "Laporan sintesis dan rekomendasi akhir",
]
cols = 3
cw2 = (CW - 0.18 * (cols - 1)) / cols
ry = dy + 0.28
rh = 0.4
for i, d in enumerate(deliverables):
    col, row = i % cols, i // cols
    x = ML + col * (cw2 + 0.18)
    y = ry + row * (rh + 0.08)
    pill(s, x, y, cw2, rh, fill=NAVY_P, line=TEAL, lw=1)
    text(s, x + 0.12, y, 0.3, rh, str(i + 1), size=11, bold=True, color=YELLOW, anchor="m")
    text(s, x + 0.4, y, cw2 - 0.52, rh, d, size=8.5, bold=True, anchor="m", spacing=0.92)

logo(s, ML, 7.12, 0.2)
text(s, ML + 1.15, 7.14, 8, 0.22, "Slide 8 v2 — dipakai di Google Slides langsung (bukan bagian PPTX/PDF utama)", size=8, color=GREY)

prs.save(OUT_PPTX)
print("saved", OUT_PPTX)
