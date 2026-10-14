"""Membuat v5: proposal lengkap dengan pelatihan dua hari (16 JPL, delapan modul).

Dasarnya v3 (desain asli, identitas Sanggabiz). Bagian 7.4 diganti kurikulum dua hari memakai penyusun tabel dari
susun_proposal_v4.py, lalu semua penyebutan durasi di bagian lain disesuaikan dari tiga hari menjadi dua hari.
Pakai: python3 -I susun_proposal_v5_2hari.py <v3.docx> <keluaran.docx>
"""
import importlib.util
import os
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

here = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("v4", os.path.join(here, "susun_proposal_v4.py"))
v4 = importlib.util.module_from_spec(spec)
spec.loader.exec_module(v4)

# Kolom lebih lebar untuk aktivitas karena modul memuat lebih banyak langkah
v4.WIDTHS = [1500, 1650, 2250, 1200, 1550, 1375]
v4.PB_HARI_PERTAMA = True

# tag langkah (K penjelasan, P praktik, R review) dan bobot waktu hanya untuk memeriksa porsi, tidak ditampilkan
PLAN = {
    1: [35, 45, 10], 2: [25, 50, 15], 3: [15, 65, 10], 4: [15, 60, 15],
    5: [20, 50, 20], 6: [20, 60, 10], 7: [20, 45, 25], 8: [15, 35, 40],
}

v4.HARI = [
    ("Hari 1. Memahami Ukuran, Merencanakan, dan Mengumpulkan Data", [
        ("1. Gambaran ABK dan menghitung FTE",
         "Menjelaskan langkah ABK dan menghitung FTE satu posisi dari jumlah pekerjaan, waktu pengerjaan, dan jam kerja efektif.",
         ["Hasil akhir ABK: lembar hitung satu posisi", "Enam langkah ABK dan beda beban kerja, kinerja, jumlah pegawai", "Beban jam, jam kerja efektif, dan FTE", "Empat kategori beban"],
         ["Peserta menebak FTE satu posisi tanpa alat hitung dan menempelkannya di flipchart.",
          "Fasilitator menunjukkan sebaran tebakan, membuka lembar hitung yang sudah jadi, menguraikannya menjadi enam langkah, dan menghitung satu pekerjaan di layar.",
          "Tiap peserta menghitung FTE satu posisi dari kasus kelompoknya dengan lembar hitung baku. Kelompok lalu mendaftar data yang masih perlu dibuktikan."],
         "Flipchart, lembar hitung FTE, paket kasus",
         "Minimal 80 persen peserta menghasilkan FTE dengan selisih paling banyak 0,10 dari kunci. Tiap kelompok mencatat minimal 3 data yang perlu dibuktikan.",
         "FTE awal berdasarkan perkiraan dan daftar data yang perlu dibuktikan"),
        ("2. Menyusun daftar pekerjaan dan merencanakan pengumpulan data",
         "Menyusun daftar pekerjaan satu posisi dan memilih sumber data tiap pekerjaan dalam rencana pengumpulan data yang dapat diperiksa.",
         ["Tingkat rincian pekerjaan yang tepat, hasil kerja, pemicu, satuan jumlah, dan frekuensi", "Pekerjaan di luar uraian jabatan", "Empat sumber data: wawancara, catatan kerja harian, pengamatan langsung, dan data transaksi", "Pekerjaan berjam terbanyak memakai minimal dua sumber, dan kerahasiaan data pribadi"],
         ["Fasilitator menjelaskan kolom formulir dengan contoh yang tepat dan yang terlalu luas, lalu menjelaskan empat sumber data dengan contoh pengurusan surat masuk.",
          "Tiap peserta mengisi daftar minimal 8 pekerjaan untuk satu posisi dari kasus kelompoknya, dengan merujuk pada SOP yang ada.",
          "Tiap kelompok memilih sumber data untuk pekerjaan berjam terbanyak dan mengisi rencana: siapa, kapan, sumber, dan bukti.",
          "Kelompok saling menilai rencana dengan tiga pertanyaan, lalu memperbaikinya."],
         "Formulir daftar pekerjaan, formulir rencana data, kartu penilaian",
         "Minimal 80 persen peserta mengisi minimal 8 pekerjaan dengan semua kolom wajib terisi. Pekerjaan berjam terbanyak punya minimal dua sumber data.",
         "Daftar pekerjaan satu posisi dan rencana pengumpulan data kelompok"),
        ("3. Teknik dan praktik wawancara ABK",
         "Menyusun panduan wawancara yang menggali pekerjaan di luar uraian jabatan, pekerjaan yang diulang, waktu menunggu, dan masa sibuk, lalu mewawancarai dengan pertanyaan netral dan mencatat bukti dengan tertib.",
         ["Susunan wawancara, pertanyaan terbuka, dan pertanyaan lanjutan", "Kekeliruan yang sering terjadi: melebihkan, lupa, menjawab sesuai harapan", "Meminta contoh nyata dan angka, serta mencatat bukti", "Tiga peran latihan: pewawancara, narasumber, pengamat"],
         ["Fasilitator menjelaskan susunan wawancara dan jenis pertanyaan.",
          "Dua fasilitator memperagakan wawancara yang menggiring dan yang netral. Peserta menilainya dengan lembar pengamatan.",
          "Peserta berpasangan menyusun panduan wawancara untuk satu posisi.",
          "Peserta berlatih dalam kelompok tiga orang, tiga putaran singkat, dan bergantian peran. Narasumber menyembunyikan tiga fakta pada kartu peran.",
          "Tiap kelompok menuliskan dua hal yang akan diubah dalam cara mewawancarai."],
         "Lembar pengamatan, kartu peran, skenario peragaan",
         "Panduan memuat minimal 10 pertanyaan terbuka, termasuk 3 tentang pekerjaan di luar uraian jabatan. Minimal 75 persen peserta menggali minimal 2 dari 3 fakta tersembunyi pada putaran terakhir.",
         "Panduan wawancara, catatan wawancara, dan lembar pengamatan terisi"),
        ("4. Mengisi lembar kerja dan mencocokkan data",
         "Mengisi lembar kerja ABK dari catatan wawancara dan mencocokkannya dengan catatan harian serta data transaksi.",
         ["Dari catatan wawancara ke lembar kerja", "Membaca catatan kerja harian", "Data transaksi sebagai pembanding", "Mencatat selisih antarsumber dan memilih sumber yang dipakai"],
         ["Fasilitator membagikan paket data: catatan wawancara, catatan harian lima hari, dan cuplikan register.",
          "Kelompok mengisi lembar kerja tiga posisi dan mencatat selisih antarsumber beserta alasan memilih sumber.",
          "Fasilitator memeriksa daftar masalah data tiap kelompok."],
         "Paket data, lembar kerja ABK",
         "Tiap kelompok mencatat minimal 3 selisih beserta alasan memilih sumber.",
         "Lembar kerja tiga posisi dan daftar masalah data"),
    ]),
    ("Hari 2. Memeriksa Data, Membaca Hasil, dan Menindaklanjuti", [
        ("5. Memeriksa mutu data",
         "Memeriksa mutu data dengan daftar periksa dan memutuskan perbaikannya.",
         ["Kelengkapan dan pekerjaan yang tercatat dua kali", "Satuan dan periode yang tidak cocok", "Angka janggal dan hitungan ganda", "Uji kewajaran jam per orang"],
         ["Fasilitator menjelaskan daftar periksa enam butir.",
          "Kelompok memeriksa lembar latihan berisi enam kesalahan sisipan, lalu memeriksa lembar kerja kelompok lain.",
          "Pemilik lembar memutuskan menerima atau menolak tiap temuan dan memperbaikinya."],
         "Lembar latihan, daftar periksa",
         "Tiap kelompok menemukan minimal 5 dari 6 kesalahan sisipan.",
         "Daftar masalah data yang sudah diputuskan dan lembar kerja yang diperbaiki"),
        ("6. Menghitung ulang FTE, membaca heatmap, dan menelusuri akar masalah",
         "Menghitung FTE dari data yang sudah diperiksa, menganalisis bagaimana asumsi mengubah kategori beban, dan menelusuri penyebab beban tinggi.",
         ["Hitung ulang FTE dan bandingkan dengan FTE awal", "Tiga pilihan jam kerja efektif, rentang FTE, dan tingkat keyakinan atas data", "Membaca heatmap per posisi dan unit", "Gejala, akar masalah, dan pohon akar masalah"],
         ["Fasilitator memberi contoh latihan: beban 1.600 jam dengan jam kerja efektif 1.250, 1.400, dan 1.500 menghasilkan FTE 1,28, 1,14, dan 1,07. Angka ini contoh latihan, bukan angka PIM.",
          "Kelompok menghitung ulang FTE tiga posisi pada tiga pilihan asumsi dan memberi tingkat keyakinan.",
          "Kelompok menyusun heatmap dari hasilnya dan pohon akar masalah untuk satu posisi yang padat atau berlebih.",
          "Kelompok membahas keputusan yang aman diambil sekarang dan yang menunggu data tambahan."],
         "Lembar hitung FTE dengan tabel asumsi, heatmap",
         "Tiap posisi punya FTE pada tiga pilihan asumsi dan tingkat keyakinan beserta alasannya. Tiap pohon akar masalah memuat minimal dua penyebab beserta buktinya.",
         "FTE akhir tiga posisi, heatmap, dan pohon akar masalah"),
        ("7. Memilih tindak lanjut, menguji RACI, dan mencocokkan dengan manajer",
         "Memilih tindak lanjut yang sesuai akar masalah, menguji dampaknya pada pembagian peran RACI, dan mempertahankan hasilnya di depan manajer.",
         ["Pilihan tindak lanjut: menyederhanakan, membagi ulang, merancang ulang, otomasi, melatih ulang, menambah kapasitas", "Uji RACI: beban penerima peran, penanggung jawab akhir ganda atau kosong", "Urutan prioritas", "Pencocokan dengan manajer dan catatan keputusan"],
         ["Fasilitator menjelaskan pencocokan akar masalah dengan tindakan dan contoh uji RACI.",
          "Kelompok memilih tiga temuan, menetapkan tindakan, membuat potongan RACI sebelum dan sesudah, dan menghitung ulang FTE penerima peran.",
          "Simulasi panel manajer: peserta kelompok lain berperan sebagai manajer. Kelompok memaparkan hasil, menjawab pertanyaan, dan mencatat keputusan."],
         "Formulir RACI, matriks prioritas, kartu pertanyaan panel",
         "Tiap tindakan punya penanggung jawab, perkiraan perubahan FTE, dan hasil uji RACI. Tiap catatan keputusan memuat alasan dan penanggung jawab.",
         "Matriks tindak lanjut, potongan RACI sebelum dan sesudah, dan catatan keputusan"),
        ("8. Rencana 30-60-90 hari dan penutupan",
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

v4.INTRO = [
    "Kurikulum disusun mengikuti urutan kerja ABK: menghitung dan merencanakan data, mengumpulkan dan memeriksa data, menghitung ulang FTE, membaca hasil, lalu menindaklanjuti. Peserta bekerja dalam lima kelompok dengan satu kasus yang sama dari Modul 1 sampai Modul 8, sehingga hasil satu modul menjadi bahan modul berikutnya. Setiap kelompok memegang tiga posisi fungsi Umum yang dipilih bersama PIM, dengan data yang namanya sudah dihapus. Bila data belum dapat dibagikan, kami menyiapkan kasus simulasi yang ditandai jelas.",
    "Tiap modul berdurasi 2 JPL (90 menit), dan tiap hari terdiri atas empat modul. Fasilitator menyesuaikan waktu tiap langkah dengan kondisi kelas. FTE dihitung dua kali, yaitu dari perkiraan pada Modul 1 dan dari data yang sudah diperiksa pada Modul 6, sehingga peserta melihat pengaruh mutu data pada kategori beban. Hasil wawancara dicocokkan dengan catatan harian dan data transaksi supaya hasil tidak bergantung pada satu sumber. Pelatihan dua hari berfokus pada keterampilan inti ABK, dan pendalaman dilakukan pada klinik daring 30 hari setelah pelatihan. Pelatihan tidak merevisi SOP.",
]

# penggantian durasi di bagian lain proposal (tiap kalimat sudah dipastikan muncul satu kali di v3)
GANTI = [
    ("Training luring selama 3 hari", "Training luring selama 2 hari"),
    ("selama tiga hari berturut-turut pada November 2026 di Medan, dengan total 24 jam pelajaran atau 18 jam efektif",
     "selama dua hari berturut-turut pada November 2026 di Medan, dengan total 16 jam pelajaran atau 12 jam efektif"),
    ("3 hari / 24 JPL", "2 hari / 16 JPL"),
    ("7.4 Kurikulum Pelatihan Tiga Hari", "7.4 Kurikulum Pelatihan Dua Hari"),
    ("Pulse survey setiap hari dan evaluasi akhir Hari 3.", "Pulse survey setiap hari dan evaluasi akhir Hari 2."),
    ("Hari 1 sampai 3", "Hari 1 sampai 2"),
    ("format tiga hari luring", "format dua hari luring"),
    ("3 hari pada November 2026", "2 hari pada November 2026"),
    ("24 JPL di Medan sesuai kurikulum.", "16 JPL di Medan sesuai kurikulum."),
    ("Kurikulum; 3 hari luring di Medan;", "Kurikulum; 2 hari luring di Medan;"),
    ("satu gelombang pelatihan tiga hari di Medan", "satu gelombang pelatihan dua hari di Medan"),
    ("Hari 2 merupakan praktik pengumpulan data dan wawancara menggunakan sampel kasus terpilih.",
     "Praktik pengumpulan data dan wawancara dalam pelatihan menggunakan sampel kasus terpilih."),
]


def porsi():
    k = sum(v[0] for v in PLAN.values())
    p = sum(v[1] for v in PLAN.values())
    r = sum(v[2] for v in PLAN.values())
    assert all(sum(v) == 90 for v in PLAN.values()), "tiap modul harus 90 menit"
    assert k + p + r == 720, "total harus 720 menit (16 JPL)"
    return [round(n * 100 / 720) for n in (k, p, r)]


def build(src, dst):
    print("Porsi rancangan (penjelasan, praktik, review) persen:", porsi())
    tmp = tempfile.mktemp(suffix=".docx")
    v4.build(src, tmp)
    zin = zipfile.ZipFile(tmp)
    x = zin.read("word/document.xml").decode("utf8")
    for lama, baru in GANTI:
        if x.count(lama) != 1:
            raise SystemExit(f"Kalimat tidak ditemukan tepat satu kali ({x.count(lama)}): {lama[:60]}")
        x = x.replace(lama, baru)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == "word/document.xml":
                data = x.encode("utf8")
            zout.writestr(it, data)
    os.remove(tmp)


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2])
