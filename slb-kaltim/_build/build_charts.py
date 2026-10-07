"""Membuat grafik untuk proposal SLB Kaltim. Data dan sumber ada di dokumen proposal."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from pathlib import Path

FONTS = Path(__file__).resolve().parents[2] / "tanoto" / "fonts"
for f in FONTS.glob("*.ttf"):
    fm.fontManager.addfont(str(f))
plt.rcParams["font.family"] = "Montserrat"

NAVY, GOLD, INK, MUTED, GRID = "#0B396D", "#FFC00C", "#1D1D1F", "#5C5F66", "#D9DCE1"
OUT = Path(__file__).resolve().parent / "assets"

# Grafik 1: guru pembimbing khusus di satuan pendidikan dengan PDBK, Desember 2023
total, punya = 40164, 5956
belum = total - punya
fig, ax = plt.subplots(figsize=(6.6, 2.1), dpi=220)
ax.barh([0], [belum], color=NAVY, height=0.42, edgecolor="white", linewidth=2)
ax.barh([0], [punya], left=[belum], color=GOLD, height=0.42, edgecolor="white", linewidth=2)
ax.set_xlim(0, total)
ax.set_ylim(-0.55, 0.95)
ax.axis("off")
ax.text(belum / 2, 0, "34.208 sekolah belum punya\nguru pembimbing khusus (85,17%)",
        color="white", ha="center", va="center", fontsize=8.5, fontweight="semibold", linespacing=1.4)
ax.annotate("5.956 sekolah punya (14,83%)", xy=(belum + punya / 2, 0.21), xytext=(belum + punya / 2, 0.34),
            ha="right", va="bottom", fontsize=8.5, fontweight="semibold", color=INK,
            arrowprops=dict(arrowstyle="-", color=MUTED, lw=0.8))
ax.text(0, 0.72, "40.164 satuan pendidikan formal memiliki peserta didik berkebutuhan khusus",
        ha="left", va="bottom", fontsize=8, color=INK, fontweight="medium")
ax.text(0, -0.34, "Sumber: Kemendikbudristek, data per Desember 2023 (dikutip Medcom dan ANTARA)",
        ha="left", va="top", fontsize=6.8, color=MUTED)
fig.savefig(OUT / "grafik-guru-pembimbing-khusus.png", bbox_inches="tight", facecolor="white")
plt.close(fig)

# Grafik 2: siswa SLB Negeri Balikpapan per jenjang
jenjang, jumlah = ["SD", "SMP", "SMA"], [219, 96, 86]
fig, ax = plt.subplots(figsize=(6.6, 2.7), dpi=220)
bars = ax.bar(jenjang, jumlah, color=NAVY, width=0.42)
for b, v in zip(bars, jumlah):
    ax.text(b.get_x() + b.get_width() / 2, v + 6, str(v), ha="center", va="bottom",
            fontsize=9.5, fontweight="semibold", color=INK)
ax.set_ylim(0, 260)
ax.set_yticks([0, 100, 200])
ax.tick_params(axis="y", labelsize=7.5, colors=MUTED, length=0)
ax.tick_params(axis="x", labelsize=9, colors=INK, length=0)
ax.grid(axis="y", color=GRID, linewidth=0.6)
ax.set_axisbelow(True)
for s in ("top", "right", "left"):
    ax.spines[s].set_visible(False)
ax.spines["bottom"].set_color(GRID)
ax.set_title("Siswa SLB Negeri Balikpapan per jenjang (total 401)", loc="left", fontsize=8.5,
             fontweight="medium", color=INK, pad=10)
fig.text(0.125, -0.02, "Sumber: ANTARA Kaltim, berita bantuan PT Angkasa Pura I untuk SLBN Balikpapan",
         ha="left", va="top", fontsize=6.8, color=MUTED)
fig.savefig(OUT / "grafik-siswa-slbn-balikpapan.png", bbox_inches="tight", facecolor="white")
plt.close(fig)
print("selesai")

# Gambar skema susunan ruang (tidak berskala)
from matplotlib.patches import FancyBboxPatch, Circle, Rectangle
import math

fig, ax = plt.subplots(figsize=(6.6, 3.7), dpi=220)
ax.set_xlim(0, 100)
ax.set_ylim(0, 56)
ax.axis("off")
ax.add_patch(FancyBboxPatch((1, 1), 98, 54, boxstyle="round,pad=0,rounding_size=1.5", fc="#F7F9FC", ec=NAVY, lw=1.6))
ax.add_patch(Rectangle((30, 50), 40, 3.2, fc=NAVY, ec="none"))
ax.text(50, 51.6, "Layar dan area pemateri", ha="center", va="center", fontsize=7, color="white", fontweight="semibold")
cols, rows = [18, 40, 62, 84], [36, 13]
n = 1
for ry in rows:
    for cx in cols:
        ax.add_patch(Circle((cx, ry), 5.2, fc=NAVY, ec="white", lw=1))
        ax.text(cx, ry, f"Kelompok\n{n}", ha="center", va="center", fontsize=5.8, color="white", fontweight="semibold", linespacing=1.2)
        for k in range(10):
            a = 2 * math.pi * k / 10
            ax.add_patch(Circle((cx + 7.6 * math.cos(a), ry + 7.6 * math.sin(a) * 0.95), 1.15, fc=GOLD, ec="none"))
        n += 1
ax.add_patch(Rectangle((2.5, 1.8), 14, 3.2, fc="white", ec=MUTED, lw=0.8))
ax.text(9.5, 3.4, "Meja registrasi", ha="center", va="center", fontsize=5.6, color=INK)
ax.add_patch(Rectangle((1, 21), 1.4, 7, fc=GOLD, ec="none"))
ax.text(3.6, 24.5, "Pintu", ha="left", va="center", fontsize=6, color=INK, rotation=90)
ax.annotate("", xy=(97, 24.5), xytext=(6, 24.5), arrowprops=dict(arrowstyle="<->", color=MUTED, lw=0.8, ls="--"))
ax.text(52, 25.4, "Jalur utama selebar dua kursi roda", ha="center", va="bottom", fontsize=6, color=MUTED)
ax.text(50, -1.5, "Skema susunan 80 peserta dalam 8 kelompok. Tidak berskala.", ha="center", va="top", fontsize=6.8, color=MUTED)
fig.savefig(OUT / "skema-susunan-ruang.png", bbox_inches="tight", facecolor="white")
plt.close(fig)
print("skema selesai")
