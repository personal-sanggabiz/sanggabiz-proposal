# ATURAN PENULISAN SANGGABIZ

Berlaku untuk semua yang kita buat: proposal (docx, pptx), dokumen kerja, README, laporan, jawaban di chat, pesan commit, dan deskripsi PR. Tujuannya satu: tulisan terbaca seperti ditulis orang yang paham pekerjaannya, bukan keluaran mesin.

Pengecualian: kutipan langsung dan salinan dokumen sumber (misalnya chat klien) tidak diubah sama sekali.

## 1. Dasar menulis yang baik
Aturan ini dirangkum dari panduan penulisan yang sudah teruji: aturan menulis George Orwell, *On Writing Well* (William Zinsser), pedoman plain language, dan kaidah kalimat efektif PUEBI.

1. **Mulai dari inti.** Kalimat pertama paragraf menyampaikan hal yang paling penting. Tidak ada pembuka basa-basi.
2. **Konkret.** Pakai nama, angka, tanggal, tempat, dan sumber. Kalau tidak ada datanya, tulis terus terang bahwa datanya belum ada.
3. **Kata sehari-hari yang baku.** Kalau ada kata pendek yang artinya sama, pakai kata itu.
4. **Buang kata yang bisa dibuang.** Baca ulang, hapus semua kata yang tidak mengubah arti.
5. **Kalimat aktif.** Sebut siapa yang melakukan apa. Di proposal pakai "kami" untuk Sanggabiz.
6. **Satu gagasan per kalimat.** Panjang kalimat boleh berbeda-beda. Ada yang pendek, ada yang agak panjang.
7. **Tulis untuk satu pembaca tertentu.** Tanya dulu siapa yang membaca dan apa yang ingin ia putuskan setelah membaca.
8. **Jujur soal batas.** Tulis "belum ada data", "perkiraan kami", atau "perlu dicek" langsung di tempatnya. Jangan menjanjikan hasil yang tidak bisa kita pegang.
9. **Pakai kosakata klien.** Kalau RFP atau klien menyebut "capacity building", pakai istilah itu.
10. **Baca keras-keras.** Kalimat yang tidak enak diucapkan biasanya perlu ditulis ulang.

## 2. Tanda baca yang tidak dipakai
| Jangan | Gantinya |
|---|---|
| Tanda pisah panjang `—` dan sedang `–` sebagai jeda kalimat | Titik atau koma. Pecah jadi dua kalimat |
| Titik koma `;` | Titik, atau koma dengan kata penghubung |
| Tanda panah `→` `➜` `⇒` | Kata: "lalu", "menjadi", "sehingga" |
| Emoji dan simbol hias (centang, bintang) | Tidak dipakai di dokumen resmi |
| Tanda seru | Tidak dipakai |
| Tanda kurung berlapis | Pisah jadi kalimat sendiri |

Boleh dipakai: tanda hubung biasa `-` untuk kata ulang (guru-guru), kata majemuk, dan rentang jam atau angka di tabel (08.00 - 08.30). Titik dua dipakai untuk membuka daftar atau penjelasan, tidak untuk efek dramatis.

## 3. Pola yang membuat tulisan terbaca seperti mesin
Hindari pola berikut.
- Pembuka klise: "Di era digital saat ini", "Dalam dunia yang terus berkembang", "Di tengah dinamika".
- Kalimat kontras retoris: "bukan sekadar X, melainkan Y", "tidak hanya X, tetapi juga Y". Pakai sesekali saja, jangan jadi kebiasaan.
- Tiga serangkai yang muncul di hampir tiap paragraf. Sebut sebanyak yang memang ada.
- Paragraf yang semuanya sama panjang dan sama bentuknya.
- Penutup yang mengulang isi: "Singkatnya", "Pada akhirnya", "Dapat disimpulkan bahwa".
- Kalimat "moral" di akhir tiap paragraf.
- Judul bergaya slogan. Judul menyebut isinya.
- Daftar poin untuk hal yang cukup ditulis dua kalimat.
- Tebal (bold) di mana-mana. Pakai hanya untuk istilah kunci atau angka penting.
- Basa-basi lawan bicara: "Tentu saja", "Pertanyaan bagus", "Mari kita".

## 4. Kata yang dihindari
Kata ini tidak dilarang mutlak, tetapi biasanya dipakai tanpa makna. Ganti dengan hal yang konkret.

holistik, komprehensif, sinergi, ekosistem (kecuali yang benar-benar ekosistem), transformasi, mengoptimalkan, memberdayakan, menyelami, menggali lebih dalam, lanskap, tonggak, inovatif, terdepan, revolusioner, solusi terbaik, menjembatani kesenjangan, memainkan peran krusial, sangat penting untuk, perlu dicatat bahwa.

Dari bahasa Inggris: leverage, seamless, robust, game-changer, unlock, delve, tapestry.

Contoh perbaikan:
- Buruk: "Program ini memberdayakan guru secara holistik melalui pendekatan komprehensif."
- Baik: "Guru belajar menyusun satu modul ajar yang bisa dipakai di kelasnya minggu depan."

## 5. Format dan angka
- Angka Indonesia: titik untuk ribuan, koma untuk desimal. Rp 13.500.000 dan 14,83 persen.
- Tanggal: 7 Oktober 2026. Jam: 08.00 WIB.
- Persentase ditulis "persen" di kalimat, "%" di tabel.
- Istilah asing hanya kalau istilah itu dipakai klien atau tidak ada padanan yang lazim (workshop, capacity building, pre-test).
- Setiap angka dan klaim di proposal punya sumber di daftar rujukan atau di catatan kaki.
- Bagian yang harus diisi tim ditandai dengan stabilo kuning dan tulisan `[ISI TIM: ...]`. Tidak ada bagian lain yang dibiarkan kosong atau berisi teks contoh.

## 6. Cara memeriksa
1. Jalankan pemeriksa otomatis: `python3 tools/cek-gaya.py <file atau folder>`. Pemeriksa menandai tanda baca terlarang dan frasa klise. Bagian di dalam blok kode dan kutipan (`>`) dilewati.
2. Baca ulang satu kali dengan pertanyaan: apakah ada kalimat yang bisa dihapus tanpa mengubah arti? Apakah ada angka tanpa sumber?
3. Pemeriksa tidak menggantikan pembacaan. Tulisan bisa lolos pemeriksa dan tetap kaku.
