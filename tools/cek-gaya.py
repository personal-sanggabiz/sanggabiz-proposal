#!/usr/bin/env python3
"""Pemeriksa gaya Sanggabiz. Lihat ATURAN-PENULISAN.md.

Pemakaian: python3 tools/cek-gaya.py <file atau folder> ...
Mengecek .md dan .docx. Blok kode, kode inline, dan baris kutipan (diawali >) dilewati.
Kode keluar 1 jika ada temuan.
"""
import re
import sys
import zipfile
from pathlib import Path

DILARANG_KARAKTER = {
    "—": "tanda pisah panjang",
    "–": "tanda pisah sedang",
    ";": "titik koma",
    "→": "tanda panah",
    "➜": "tanda panah",
    "⇒": "tanda panah",
    "✓": "tanda centang",
    "✔": "tanda centang",
    "✅": "emoji centang",
}
FRASA = [
    "di era digital", "dalam era digital", "di tengah dinamika", "holistik",
    "komprehensif", "sinergi", "memberdayakan", "mengoptimalkan", "menyelami",
    "menggali lebih dalam", "lanskap", "tonggak", "revolusioner", "solusi terbaik",
    "menjembatani kesenjangan", "peran krusial", "sangat penting untuk",
    "perlu dicatat bahwa", "dapat disimpulkan bahwa", "singkatnya", "pada akhirnya",
    "bukan sekadar", "tidak hanya", "mari kita", "tentu saja",
    "leverage", "seamless", "robust", "game-changer", "delve", "tapestry",
]
NEGATIF = ["sempit", "kecil", "murah", "mepet", "lemah", "buruk", "minim"]  # hanya untuk .docx (dokumen klien)
EMOJI = re.compile("[\U0001F300-\U0001FAFF☀-⛿✀-➿]")
ABAIKAN = {"ATURAN-PENULISAN.md"}


def baca_docx(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf8")
    xml = re.sub(r"</w:p>", "\n", xml)
    xml = re.sub(r"<w:tab/>", " ", xml)
    teks = re.sub(r"<[^>]+>", "", xml)
    for a, b in (("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"), ("&quot;", '"'), ("&apos;", "'")):
        teks = teks.replace(a, b)
    return teks.splitlines()


def baris_md(path):
    dalam_blok = False
    for no, baris in enumerate(path.read_text(encoding="utf8").splitlines(), 1):
        if baris.strip().startswith("```"):
            dalam_blok = not dalam_blok
            continue
        if dalam_blok or baris.lstrip().startswith(">"):
            continue
        yield no, re.sub(r"`[^`]*`", "", baris)


def periksa(path):
    path = Path(path)
    if path.name in ABAIKAN or "sumber" in path.parts or "reference" in path.parts:
        return 0
    if path.suffix == ".md" and "<!-- salinan-sumber -->" in path.read_text(encoding="utf8"):
        return 0
    if path.suffix == ".docx":
        sumber = enumerate(baca_docx(path), 1)
    elif path.suffix == ".md":
        sumber = baris_md(path)
    else:
        return 0
    n = 0
    for no, baris in sumber:
        rendah = baris.lower()
        for k, nama in DILARANG_KARAKTER.items():
            if k in baris:
                print(f"{path}:{no}: {nama} '{k}': {baris.strip()[:90]}")
                n += 1
        if EMOJI.search(baris):
            print(f"{path}:{no}: emoji: {baris.strip()[:90]}")
            n += 1
        if path.suffix == ".docx":
            for w in NEGATIF:
                if re.search(r"\b" + w + r"\b", rendah):
                    print(f"{path}:{no}: kata negatif '{w}': {baris.strip()[:90]}")
                    n += 1
        for f in FRASA:
            if f in rendah:
                print(f"{path}:{no}: frasa '{f}': {baris.strip()[:90]}")
                n += 1
    return n


def main(args):
    if not args:
        print(__doc__)
        return 2
    total = 0
    for a in args:
        p = Path(a)
        files = sorted(p.rglob("*")) if p.is_dir() else [p]
        for f in files:
            if f.is_file() and f.suffix in (".md", ".docx") and ".git" not in f.parts:
                total += periksa(f)
    print(f"Temuan: {total}")
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
