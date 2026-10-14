"""Membuat v4 proposal: salinan v3 (desain asli, identitas Sanggabiz) dengan bagian 7.4 diganti kurikulum baru.

Hanya isi bagian 7.4 yang berubah. Desain, font, header, dan bagian lain tidak disentuh.
Pakai: python3 -I susun_proposal_v4.py <v3.docx> <keluaran.docx>
"""
import re
import sys
import zipfile
from xml.sax.saxutils import escape

NAVY, ALT, LINE = "102a4c", "eaf2f7", "d9d9d9"
SZ = 16  # 8 pt, sedikit di bawah tabel asli (8,5 pt) agar enam kolom muat di halaman tegak
WIDTHS = [1500, 1650, 2200, 1250, 1550, 1375]  # jumlah 9525, sama dengan lebar tabel asli
HEAD = ["Topik atau modul", "Outline materi", "Aktivitas pembelajaran", "Alat atau media", "Indikator capaian", "Output konkret"]

HARI = [
    ("Hari 1. Memahami Ukuran dan Merancang Pengumpulan Data", [
        ("1. Gambaran ABK dan hasil akhir yang dituju",
         "Menjelaskan langkah ABK dan membedakan beban kerja, kinerja, dan jumlah pegawai.",
         ["Hasil akhir ABK: lembar hitung satu posisi", "Enam langkah ABK", "Beban kerja, kinerja, dan jumlah pegawai", "Temuan studi PIM sebelumnya"],
         ["Peserta menebak FTE satu posisi tanpa alat hitung dan menempelkannya di flipchart.",
          "Fasilitator menunjukkan sebaran tebakan, membuka lembar hitung yang sudah jadi, dan menguraikannya menjadi enam langkah.",
          "Tiap kelompok menulis lima masalah beban kerja dari kasusnya."],
         "Flipchart, kartu, paket kasus",
         "Tiap kelompok menulis minimal lima masalah. Minimal 80 persen peserta menjawab benar 2 dari 3 butir kuis.",
         "Daftar masalah beban kerja awal"),
        ("2. Menyusun daftar pekerjaan satu posisi",
         "Menyusun daftar pekerjaan satu posisi lengkap dengan hasil kerja, pemicu, satuan jumlah, dan frekuensi.",
         ["Tingkat rincian pekerjaan yang tepat", "Hasil kerja, pemicu, dan satuan jumlah", "Frekuensi dan masa sibuk", "Pekerjaan di luar uraian jabatan"],
         ["Fasilitator menjelaskan kolom formulir dengan contoh yang tepat dan yang terlalu luas.",
          "Tiap peserta mengisi daftar minimal 8 pekerjaan untuk satu posisi dari kasus kelompoknya, dengan merujuk pada SOP yang ada.",
          "Peserta bertukar daftar berpasangan dan memeriksanya dengan daftar periksa."],
         "Formulir daftar pekerjaan, daftar periksa",
         "Minimal 80 persen peserta mengisi minimal 8 pekerjaan dengan semua kolom wajib terisi.",
         "Daftar pekerjaan satu posisi"),
        ("3. Menghitung beban kerja dan FTE",
         "Menghitung FTE satu posisi dari jumlah pekerjaan, waktu pengerjaan, dan jam kerja efektif.",
         ["Beban jam: jumlah per tahun dikali waktu satu kali", "Jam kerja efektif setahun", "FTE dan empat kategori beban", "Tingkat keyakinan atas data"],
         ["Fasilitator menghitung satu pekerjaan dan jam kerja efektif di layar.",
          "Tiap peserta menghitung FTE dari daftar pekerjaannya dengan lembar hitung baku.",
          "Kelompok membandingkan hasil dan mendaftar data yang masih perlu dibuktikan."],
         "Lembar hitung FTE, kalkulator",
         "Minimal 80 persen peserta menghasilkan FTE dengan selisih paling banyak 0,10 dari kunci.",
         "FTE awal berdasarkan perkiraan dan daftar data yang perlu dibuktikan"),
        ("4. Merencanakan pengumpulan data",
         "Memilih sumber data untuk tiap pekerjaan dan menyusun rencana pengumpulan data yang dapat diperiksa.",
         ["Empat sumber data: wawancara, catatan kerja harian, pengamatan langsung, dan data transaksi", "Kekuatan dan batas tiap sumber", "Pekerjaan berjam terbanyak memakai minimal dua sumber", "Kerahasiaan data pribadi"],
         ["Fasilitator menjelaskan empat sumber data dengan contoh pengurusan surat masuk.",
          "Tiap kelompok memilih sumber data untuk pekerjaan berjam terbanyak dan mengisi rencana: siapa, kapan, sumber, dan bukti.",
          "Kelompok saling menilai rencana dengan tiga pertanyaan, lalu memperbaikinya."],
         "Formulir rencana data, kartu penilaian",
         "Pekerjaan berjam terbanyak punya minimal dua sumber data. Semua baris rencana memuat penanggung jawab, tanggal, dan bukti.",
         "Rencana pengumpulan data kelompok"),
    ]),
    ("Hari 2. Mengumpulkan dan Memeriksa Data", [
        ("5. Teknik wawancara ABK",
         "Menyusun panduan wawancara yang menggali pekerjaan di luar uraian jabatan, pekerjaan yang diulang, waktu menunggu, dan masa sibuk.",
         ["Susunan wawancara", "Pertanyaan terbuka dan pertanyaan lanjutan", "Kekeliruan yang sering terjadi: melebihkan, lupa, menjawab sesuai harapan", "Meminta contoh nyata dan angka"],
         ["Fasilitator menjelaskan susunan wawancara dan jenis pertanyaan.",
          "Dua fasilitator memperagakan wawancara yang menggiring dan yang netral. Peserta menilainya dengan lembar pengamatan.",
          "Peserta berpasangan menyusun panduan wawancara untuk satu posisi."],
         "Lembar pengamatan, skenario peragaan",
         "Panduan memuat minimal 10 pertanyaan terbuka, termasuk 3 tentang pekerjaan di luar uraian jabatan.",
         "Panduan wawancara"),
        ("6. Praktik wawancara dengan bermain peran",
         "Mewawancarai dengan pertanyaan netral dan mencatat bukti dengan tertib.",
         ["Tiga peran: pewawancara, narasumber, pengamat", "Mencatat angka dan bukti", "Pertanyaan lanjutan yang netral", "Umpan balik dengan lembar pengamatan"],
         ["Fasilitator menjelaskan skenario dan kartu peran. Narasumber menyembunyikan tiga fakta.",
          "Peserta berlatih dalam kelompok tiga orang, tiga putaran, dan bergantian peran.",
          "Tiap kelompok menuliskan dua hal yang akan diubah dalam cara mewawancarai."],
         "Kartu peran, lembar pengamatan",
         "Minimal 75 persen peserta menggali minimal 2 dari 3 fakta tersembunyi pada putaran terakhir.",
         "Catatan wawancara dan lembar pengamatan terisi"),
        ("7. Mengisi lembar kerja dan mencocokkan data",
         "Mengisi lembar kerja ABK dari catatan wawancara dan mencocokkannya dengan catatan harian serta data transaksi.",
         ["Dari catatan wawancara ke lembar kerja", "Membaca catatan kerja harian", "Data transaksi sebagai pembanding", "Mencatat selisih antarsumber dan memilih sumber yang dipakai"],
         ["Fasilitator membagikan paket data: catatan wawancara, catatan harian lima hari, dan cuplikan register.",
          "Kelompok mengisi lembar kerja tiga posisi dan mencatat selisih antarsumber beserta alasan memilih sumber.",
          "Fasilitator memeriksa daftar masalah data tiap kelompok."],
         "Paket data, lembar kerja ABK",
         "Tiap kelompok mencatat minimal 3 selisih beserta alasan memilih sumber.",
         "Lembar kerja tiga posisi dan daftar masalah data"),
        ("8. Memeriksa mutu data",
         "Memeriksa mutu data dengan daftar periksa dan memutuskan perbaikannya.",
         ["Kelengkapan dan pekerjaan yang tercatat dua kali", "Satuan dan periode yang tidak cocok", "Angka janggal dan hitungan ganda", "Uji kewajaran jam per orang"],
         ["Fasilitator menjelaskan daftar periksa enam butir.",
          "Kelompok memeriksa lembar latihan berisi enam kesalahan sisipan, lalu memeriksa lembar kerja kelompok lain.",
          "Pemilik lembar memutuskan menerima atau menolak tiap temuan dan memperbaikinya."],
         "Lembar latihan, daftar periksa",
         "Tiap kelompok menemukan minimal 5 dari 6 kesalahan sisipan.",
         "Daftar masalah data yang sudah diputuskan dan lembar kerja yang diperbaiki"),
    ]),
    ("Hari 3. Menghitung, Membaca Hasil, dan Menindaklanjuti", [
        ("9. Menghitung ulang FTE dan menguji asumsi",
         "Menghitung FTE dari data yang sudah diperiksa dan menganalisis bagaimana asumsi mengubah kategori beban.",
         ["Hitung ulang FTE dan bandingkan dengan FTE awal", "Tiga pilihan jam kerja efektif", "Rentang FTE dan kategori beban", "Tingkat keyakinan atas data"],
         ["Fasilitator memberi contoh latihan: beban 1.600 jam dengan jam kerja efektif 1.250, 1.400, dan 1.500 menghasilkan FTE 1,28, 1,14, dan 1,07. Angka ini contoh latihan, bukan angka PIM.",
          "Kelompok menghitung ulang FTE tiga posisi pada tiga pilihan asumsi dan memberi tingkat keyakinan.",
          "Kelompok membahas keputusan yang aman diambil dan yang menunggu data tambahan."],
         "Lembar hitung FTE dengan tabel asumsi",
         "Tiap posisi punya FTE pada tiga pilihan asumsi dan tingkat keyakinan beserta alasannya.",
         "FTE akhir tiga posisi dengan rentang hasil"),
        ("10. Membaca heatmap dan menelusuri akar masalah",
         "Membaca heatmap, menelusuri penyebab beban tinggi, dan mempertahankan hasilnya di depan manajer.",
         ["Membaca heatmap per posisi dan unit", "Gejala dan akar masalah", "Pohon akar masalah", "Catatan keputusan"],
         ["Fasilitator menunjukkan contoh heatmap dan pohon akar masalah.",
          "Kelompok menyusun heatmap dan pohon akar masalah untuk satu posisi yang padat atau berlebih.",
          "Simulasi panel manajer: kelompok memaparkan hasil, menjawab pertanyaan, dan mencatat keputusan."],
         "Heatmap, kartu pertanyaan panel",
         "Tiap pohon akar masalah memuat minimal dua penyebab beserta buktinya.",
         "Heatmap, pohon akar masalah, dan catatan keputusan"),
        ("11. Memilih tindak lanjut dan menguji dampaknya pada RACI",
         "Memilih tindak lanjut yang sesuai akar masalah dan menguji dampaknya pada pembagian peran RACI.",
         ["Pilihan tindak lanjut: menyederhanakan, membagi ulang, merancang ulang, otomasi, melatih ulang, menambah kapasitas", "Mencocokkan akar masalah dengan tindakan", "Uji RACI: beban penerima peran, penanggung jawab akhir ganda atau kosong", "Urutan prioritas"],
         ["Fasilitator menjelaskan pencocokan akar masalah dengan tindakan dan contoh uji RACI.",
          "Kelompok memilih tiga temuan, menetapkan tindakan, membuat potongan RACI sebelum dan sesudah, dan menghitung ulang FTE penerima peran.",
          "Kelompok bertukar matriks dan mencari satu masalah pada hasil kelompok lain."],
         "Formulir RACI, matriks prioritas",
         "Tiap tindakan punya penanggung jawab, perkiraan perubahan FTE, dan hasil uji RACI.",
         "Matriks tindak lanjut dan potongan RACI sebelum dan sesudah"),
        ("12. Rencana 30-60-90 hari dan penutupan",
         "Menyusun rencana 30-60-90 hari untuk unitnya dan menunjukkan penguasaan lewat post-test.",
         ["Rencana 30-60-90 hari dengan penanggung jawab dan ukuran", "Kaitan dengan klinik daring 30 hari dan tinjauan 60 sampai 90 hari setelah pelatihan", "Komitmen penerapan"],
         ["Fasilitator menjelaskan cara menulis target yang terukur.",
          "Tiap peserta menyusun rencana 30-60-90 hari untuk unitnya dan berbagi dengan dua rekan.",
          "Peserta mengerjakan post-test, mengisi evaluasi akhir, dan menerima sertifikat elektronik."],
         "Formulir rencana, soal post-test",
         "Skor post-test minimal 75 dan peningkatan rata-rata minimal 20 persen dari nilai awal.",
         "Rencana 30-60-90 hari, hasil post-test, dan komitmen"),
    ]),
]

INTRO = [
    "Kurikulum disusun mengikuti urutan kerja ABK: merencanakan dan mengumpulkan data, memeriksa mutu data, menghitung FTE, membaca hasil, lalu menindaklanjuti. Peserta bekerja dalam lima kelompok dengan satu kasus yang sama dari Modul 1 sampai Modul 12, sehingga hasil satu modul menjadi bahan modul berikutnya. Setiap kelompok memegang tiga posisi fungsi Umum yang dipilih bersama PIM, dengan data yang namanya sudah dihapus. Bila data belum dapat dibagikan, kami menyiapkan kasus simulasi yang ditandai jelas.",
    "Tiap modul berdurasi 2 JPL (90 menit). Fasilitator menyesuaikan waktu tiap langkah dengan kondisi kelas. FTE dihitung dua kali, yaitu dari perkiraan pada Modul 3 dan dari data yang sudah diperiksa pada Modul 9, sehingga peserta melihat pengaruh mutu data pada kategori beban. Hasil wawancara dicocokkan dengan catatan harian dan data transaksi supaya hasil tidak bergantung pada satu sumber. Pelatihan tidak merevisi SOP.",
]

ARIAL = '<w:rFonts w:ascii="Arial" w:cs="Arial" w:eastAsia="Arial" w:hAnsi="Arial"/>'


def run(text, bold=False, color=None, sz=SZ):
    rpr = ARIAL + ('<w:b w:val="1"/><w:bCs w:val="1"/>' if bold else "") + (f'<w:color w:val="{color}"/>' if color else "") + f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/><w:rtl w:val="0"/>'
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def para(runs, after=40, jc="left", sz=SZ, color=None, bold=False):
    ppr = f'<w:spacing w:after="{after}" w:line="240" w:lineRule="auto"/>' + (f'<w:jc w:val="{jc}"/>' if jc else "")
    ppr += "<w:rPr>" + ARIAL + (f'<w:color w:val="{color}"/>' if color else "") + f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    return f"<w:p><w:pPr>{ppr}</w:pPr>{runs}</w:p>"


def cell(paras, width, fill=None, head=False):
    border = lambda c: f'w:color="{LINE}" w:space="0" w:sz="5" w:val="single"'
    tcb = "<w:tcBorders>" + "".join(f"<w:{s} {border(s)}/>" for s in ("top", "left", "bottom", "right")) + "</w:tcBorders>"
    shd = f'<w:shd w:fill="{fill}" w:val="clear"/>' if fill else ""
    mar = '<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tcMar>'
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{tcb}{shd}{mar}<w:vAlign w:val="top"/></w:tcPr>{"".join(paras)}</w:tc>'


def table(rows):
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in WIDTHS)
    tblpr = (f'<w:tblPr><w:tblStyle w:val="Table18"/><w:tblW w:w="{sum(WIDTHS)}" w:type="dxa"/><w:jc w:val="left"/>'
             '<w:tblBorders><w:top w:color="000000" w:space="0" w:sz="0" w:val="nil"/><w:left w:color="000000" w:space="0" w:sz="0" w:val="nil"/>'
             '<w:bottom w:color="000000" w:space="0" w:sz="0" w:val="nil"/><w:right w:color="000000" w:space="0" w:sz="0" w:val="nil"/>'
             '<w:insideH w:color="000000" w:space="0" w:sz="0" w:val="nil"/><w:insideV w:color="000000" w:space="0" w:sz="0" w:val="nil"/></w:tblBorders>'
             '<w:tblLayout w:type="fixed"/><w:tblLook w:val="0600"/></w:tblPr>')
    head = '<w:tr><w:trPr><w:cantSplit w:val="1"/><w:tblHeader w:val="1"/></w:trPr>' + "".join(
        cell([para(run(h, bold=True, color="ffffff"), after=0, jc="center")], WIDTHS[i], fill=NAVY, head=True) for i, h in enumerate(HEAD)) + "</w:tr>"
    body = ""
    for ri, (judul, tujuan, outline, akt, alat, ukur, out) in enumerate(rows):
        fill = ALT if ri % 2 else None
        cells = [
            [para(run(judul, bold=True)), para(run("Tujuan: " + tujuan[0].lower() + tujuan[1:]))],
            [para(run(f"{i + 1}. {t}")) for i, t in enumerate(outline)],
            [para(run(f"{i + 1}. {t}")) for i, t in enumerate(akt)],
            [para(run(alat))],
            [para(run(ukur))],
            [para(run(out))],
        ]
        body += '<w:tr><w:trPr><w:cantSplit w:val="1"/></w:trPr>' + "".join(cell(c, WIDTHS[i], fill=fill) for i, c in enumerate(cells)) + "</w:tr>"
    return f"<w:tbl>{tblpr}<w:tblGrid>{grid}</w:tblGrid>{head}{body}</w:tbl>"


def normal_para(text):
    ppr = '<w:spacing w:after="120" w:line="268.8" w:lineRule="auto"/><w:rPr>' + ARIAL + "</w:rPr>"
    r = f'<w:r><w:rPr>{ARIAL}<w:rtl w:val="0"/></w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
    return f"<w:p><w:pPr>{ppr}</w:pPr>{r}</w:p>"


PB_HARI_PERTAMA = False  # True: judul hari pertama pindah ke halaman baru bersama tabelnya


def sub_heading(text, pb=False):
    rpr = ARIAL + '<w:color w:val="0c7788"/><w:sz w:val="24"/><w:szCs w:val="24"/>'
    ppr = f'<w:pStyle w:val="heading2"/><w:keepNext w:val="1"/><w:keepLines w:val="0"/><w:spacing w:after="100" w:before="160" w:lineRule="auto"/><w:rPr>{rpr}</w:rPr>'
    # pemisah halaman eksplisit (bukan pageBreakBefore) karena dikenali semua program pembaca docx
    br = f'<w:r><w:rPr>{rpr}</w:rPr><w:br w:type="page"/></w:r>' if pb else ""
    return f'<w:p><w:pPr>{ppr}</w:pPr>{br}<w:r><w:rPr>{rpr}<w:rtl w:val="0"/></w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def build(src, dst):
    zin = zipfile.ZipFile(src)
    x = zin.read("word/document.xml").decode("utf8")
    i = x.index("7.4 Kurikulum Pelatihan Tiga Hari")
    t0 = x.index("<w:tbl>", i)
    t1 = x.index("</w:tbl>", t0) + len("</w:tbl>")
    isi = "".join(normal_para(t) for t in INTRO)
    for n, (judul, modul) in enumerate(HARI):
        isi += sub_heading(judul, pb=(n == 0 and PB_HARI_PERTAMA)) + table(modul)
    x2 = x[:t0] + isi + x[t1:]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == "word/document.xml":
                data = x2.encode("utf8")
            zout.writestr(it, data)


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2])
