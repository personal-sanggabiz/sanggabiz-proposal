"""Ganti identitas Levner menjadi Sanggabiz pada dokumen v1 (format dan isi tidak diubah).

Pakai: python3 -I rebrand_v1_ke_sanggabiz.py <v1.docx> <keluaran.docx>
Logo Sanggabiz diambil dari compro 2026 (siluet putih bertransparansi), diwarnai navy 0B396D.
"""
import io, re, sys, zipfile
from pathlib import Path
from PIL import Image

NAVY = (0x0B, 0x39, 0x6D)
ASSETS = Path(__file__).resolve().parent / "assets"

def logo(width):
    src = Image.open(ASSETS / "sanggabiz_logo_alpha_src.png").convert("RGBA")
    a = src.split()[3]
    bbox = a.getbbox()
    a = a.crop(bbox)
    h = round(a.size[1] * width / a.size[0])
    a = a.resize((width, h), Image.LANCZOS)
    out = Image.new("RGBA", (width, h), NAVY + (0,))
    out.putalpha(a)
    return out

def swap(png, box, width):
    """Tutup logo Levner di kotak box (x0,y0,x1,y1) dengan warna latar, lalu tempel logo Sanggabiz."""
    im = Image.open(io.BytesIO(png)).convert("RGB")
    x0, y0, x1, y1 = box
    bg = im.getpixel((x0 - 4, y0 - 4))
    im.paste(bg, box)
    lg = logo(width)
    cx, cy = (x0 + x1) // 2, (y0 + y1) // 2
    im.paste(lg, (cx - lg.size[0] // 2, cy - lg.size[1] // 2), lg)
    buf = io.BytesIO(); im.save(buf, "PNG", optimize=True)
    return buf.getvalue()

def swap_text(x):
    x = x.replace("LEVNER CONSULTING", "SANGGABIZ")
    x = x.replace("Levner Consulting", "Sanggabiz")
    return re.sub(r"\bLevner\b", "Sanggabiz", x)

def main(src, dst):
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename in ("word/document.xml", "word/header1.xml", "word/footer1.xml"):
                data = swap_text(data.decode("utf8")).encode("utf8")
            elif it.filename == "word/media/image2.png":   # latar halaman isi, logo di kanan atas
                data = swap(data, (1190, 55, 1325, 125), 128)
            elif it.filename == "word/media/image3.png":   # sampul
                data = swap(data, (1040, 160, 1235, 255), 175)
            zout.writestr(it, data)

if __name__ == "__main__":
    main(*sys.argv[1:3])
