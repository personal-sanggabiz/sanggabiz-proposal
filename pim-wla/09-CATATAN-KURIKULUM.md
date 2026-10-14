# CATATAN TIM: Kurikulum pelatihan tiga hari ABK PIM

Isi catatan ini hanya untuk tim. Dokumen klien terbaru ada di `kurikulum/v3 Kurikulum Pelatihan Tiga Hari ABK - PT Pupuk Iskandar Muda.docx` dan `.pdf` (15 halaman). Dokumen itu lampiran proposal yang menggantikan bagian 7.4 pada `proposal/v3`. `proposal/v3` dan `proposal/v2` tidak diubah.

| Versi | Isi | Skrip pembuat |
|---|---|---|
| `kurikulum/v1` | Tabel lima kolom dengan menit per langkah | `_build/build_kurikulum_v1.js` |
| `kurikulum/v2` | Enam kolom, outline di kolom sendiri, tanpa menit per langkah | `_build/build_kurikulum_v2.js` |
| `kurikulum/v3` | Siap kirim: tanpa sampul dan tanpa kata internal, bahasa awam, daftar istilah, rujukan buku untuk Bloom, Kolb, dan backward design | `_build/build_kurikulum.js` |

Versi yang masuk ke proposal: `proposal/v4` (31 halaman). Desain dan isinya sama dengan `proposal/v3` (turunan langsung berkas sumber, font Arial, header dan latar halaman asli), hanya bagian 7.4 yang diganti. Skripnya `_build/susun_proposal_v4.py` mengedit salinan v3 secara langsung, sehingga tidak ada perubahan font atau tata letak di bagian lain. Isi 7.4 sengaja ringkas: pengantar dua paragraf dan tiga tabel per hari dengan enam kolom (huruf 8 pt). Daftar istilah dan rujukan tidak dimasukkan. Dokumen kurikulum terpisah (`kurikulum/v3`) tetap dipakai bila perlu versi rinci.

Aturan baru dari revisi ini ada di `ATURAN-PROPOSAL.md` bagian 3 (tanpa menit per langkah, outline kolom sendiri), bagian 6 butir 4 (bahasa awam), dan bagian 9 (tidak ada catatan internal di dokumen klien).

## 1. Yang diubah dari 7.4 di v3
| Hal | Di v3 (7.4) | Di kurikulum baru |
|---|---|---|
| Urutan | Hari 1 membaca heatmap dan menghitung FTE, padahal data belum dikumpulkan. Pengumpulan data baru Hari 2 | Mengikuti urutan kerja ABK: daftar pekerjaan, rencana data, wawancara, lembar kerja, pemeriksaan mutu, hitung FTE, baca hasil, tindak lanjut |
| Kolom tabel | 4 kolom (hari, modul, materi, bukti belajar) | 6 kolom: topik dan tujuan, outline materi, aktivitas, alat, indikator, output |
| Aktivitas | Daftar istilah per modul | Langkah bernomor: siapa melakukan apa, bahan apa yang dipegang peserta, berlabel konsep, praktik, atau review, tanpa pembagian menit |
| Kasus | Tiap modul berdiri sendiri | Satu kasus dari Modul 1 sampai 12. Lima kelompok masing-masing memegang tiga posisi, output satu modul menjadi bahan modul berikutnya |
| Hitung FTE | Sekali, dengan contoh kasus | Dua kali: Modul 3 (perkiraan) dan Modul 9 (data teruji), plus uji sensitivitas jam kerja efektif |
| Wawancara | Teori dan bermain peran dalam modul yang padat | Dua modul: teknik (dengan peragaan dua versi) dan bermain peran tiga putaran dengan kartu peran berisi fakta tersembunyi, sehingga keberhasilan terukur |
| Hari 3 | Modul 12 memuat rencana tindak, post-test, dan penutupan dalam satu baris | Modul 12 dipecah menjadi lima langkah berurutan: penjelasan target terukur, rencana 30-60-90, berbagi, post-test, evaluasi dan penutupan |
| Pembuka | Pre-test | Peserta menebak FTE tanpa alat hitung lebih dulu, lalu melihat sebaran tebakan sebelum menerima rumus (Kolb) |

Total tetap 12 modul, 24 JPL, 30 peserta, dua fasilitator. Porsi rancangan: 25 persen penjelasan, 54 persen praktik, 21 persen review. Angka ini dihitung skrip dari bobot waktu per langkah yang hanya dipakai untuk perhitungan dan tidak ditampilkan di dokumen. Skrip menolak modul yang tidak berjumlah 90 menit.

## 2. Keputusan dan hal yang perlu dikonfirmasi
1. **Data kasus.** Rancangan memakai tiga posisi fungsi Umum per kelompok, data anonim dari PIM, dipilih satu sampai dua minggu sebelum pelatihan. Apakah PIM mau membagikan? Bila tidak, tim menyiapkan kasus simulasi. Bobot kerja menyiapkannya perlu dihitung.
2. **Angka latihan buatan kami.** Contoh 1.600 jam dengan jam efektif 1.250, 1.400, dan 1.500 (FTE 1,28, 1,14, 1,07), ambang selisih 20 persen antarsumber, dan 8 aktivitas per posisi bukan angka PIM. Di dokumen sudah dilabeli sebagai contoh latihan.
3. **Bank soal.** Pre-test dan post-test 15 butir (12 pilihan ganda, 3 hitungan) belum ada. Perlu dibuat dengan kisi-kisi yang sama.
4. **Beban persiapan.** Kartu peran, lembar uji berisi enam kesalahan sisipan, paket data (catatan wawancara, catatan harian lima hari, cuplikan register), dan kunci jawaban lebih banyak daripada di v3. Cek apakah komponen pelatihan Rp40 juta masih cukup.
5. **Jam kelas.** Kurikulum menambah sesi ulasan 30 menit pukul 08.30 sampai 09.00 pada Hari 2 dan 3, di luar 24 JPL. Cek dengan jam kerja PIM.
6. **Ruang.** Wawancara bermain peran memakai 10 kelompok tiga orang, jadi ruang perlu cukup untuk berpencar. Ini perlu masuk ke kebutuhan tempat (bagian 8.7 di v2).

## 3. Dampak ke bagian lain bila kurikulum ini dipakai
| Bagian | Yang perlu disesuaikan |
|---|---|
| 7.3 Tujuan pembelajaran dan porsi belajar | Tertulis 25 persen konsep, 55 persen praktik, 20 persen review. Hasil rancangan 25, 54, 21. Selisihnya ringan, tetapi samakan |
| 7.3 Tujuan pembelajaran | Masih memakai istilah Inggris (hidden work, seasonality, annualisasi). Tidak diubah karena fokus revisi hanya 7.4, tetapi tidak sejalan dengan bahasa awam di 7.4 baru |
| 7.5 Kirkpatrick | Tidak diubah, sesuai arahan. Rujukan pulse survey harian dan post-test di Modul 12 sudah sejalan |
| 7.6 Luaran | Tambahkan rencana pengumpulan data, rubrik wawancara, catatan keputusan, template RACI sebelum dan sesudah, matriks prioritas, lembar hitung FTE baku dengan tabel skenario, kartu peran, dan lembar uji |
| Tabel jadwal pelaksanaan | Tidak berubah |

## 4. Riset dan status verifikasi
Pencarian web berjalan. Pembukaan halaman penuh gagal untuk `peraturan.go.id` dan `pasal.id` (proxy sesi menolak, kode 403), jadi isi di bawah berasal dari ringkasan hasil pencarian, bukan dari teks penuh sumbernya. Tim perlu membuka sumber sebelum dokumen dikirim.

| Rujukan di dokumen | Dipakai untuk | Status |
|---|---|---|
| [1] Permen PANRB 1/2020, [JDIH BPK](https://peraturan.bpk.go.id/Home/Details/143660/permen-pan-rb-no-1-tahun-2020) | Urutan kerja ABK | Judul dan kedudukan terkonfirmasi lewat pencarian. Lampiran belum dibaca |
| [2] [Kementerian PANRB, langkah penyusunan Anjab dan ABK](https://www.menpan.go.id/site/berita-terkini/langkah-penyusunan-anjab-dan-abk-dalam-penentuan-jabatan) | Tahapan: tim, analisis jabatan, pengumpulan data, pengolahan, verifikasi dan validasi | Dari ringkasan pencarian, halaman belum dibuka |
| [3] Wiggins dan McTighe 2005, Understanding by Design | Backward design | Buku dikenal, diketik dari pengetahuan umum. Cocokkan penerbit dan edisi |
| [4] Kolb 1984, Experiential Learning | Siklus belajar mengalami, merenungkan, memahami konsep, menerapkan | Buku dikenal, diketik dari pengetahuan umum. Cocokkan data terbitan |
| [5] [Improving Stylised Working Time Estimates with Time Diary Data](https://www.ncbi.nlm.nih.gov/pmc/articles/PMC6609582/) | Catatan harian sebagai pembanding perkiraan umum jam kerja | Dari ringkasan pencarian. Nama penulis belum ada, lengkapi |
| [6] [Harms dkk. 2019, validasi diary HETUS](https://pmc.ncbi.nlm.nih.gov/articles/PMC6546612) | Validasi catatan harian bertumpu pada beberapa penanda | Dari ringkasan pencarian |
| [7] Anderson dan Krathwohl 2001, A Taxonomy for Learning, Teaching, and Assessing | Taksonomi Bloom revisi | Buku dikenal, diketik dari pengetahuan umum. Cocokkan data terbitan |

Catatan penting soal rujukan:
0. Rujukan UF/IFAS (EDIS) untuk Kolb dikeluarkan dari v3 karena judulnya tidak terlihat dan pernyataan "instruksi seminimal mungkin" tidak lagi dipakai.
1. Permen PANRB 1/2020 mengatur instansi pemerintah, sedangkan PIM adalah BUMN. Dokumen klien menyebutnya sebagai rujukan urutan kerja, dan ambang FTE tetap mengikuti studi PIM sebelumnya. Jangan menyebutnya sebagai kewajiban bagi PIM.
2. Temuan yang tidak dimuat di dokumen klien karena belum terverifikasi: ringkasan satu jurnal 2025 menyebut kelonggaran 30 persen pada Permen PANRB 1/2020 dan Peraturan BKN 9/2022, serta 25 persen pada Permendagri 12/2008. Penjelasan Kementerian PANRB menyebut beban minimal 1.250 jam per tahun per jabatan. Keduanya perlu dicek ke lampiran resmi sebelum dipakai untuk menetapkan jam kerja efektif PIM.
3. Pola "70-20-10" dan "piramida belajar" tidak dipakai karena bukti kuatnya lemah menurut hasil pencarian. Porsi 25/54/21 adalah rancangan kami, bukan angka dari penelitian.
4. Cara membuat ulang dokumen: `NODE_PATH=$(npm root -g) node _build/build_kurikulum.js`. Data modul ada di konstanta `HARI` di skrip itu.

## 5. Varian dua hari (proposal v5)
Berkas: `proposal/v5 Proposal ... Pelatihan 2 Hari - PT Pupuk Iskandar Muda.docx` (31 halaman, desain sama dengan v3 dan v4). Skrip: `_build/susun_proposal_v5_2hari.py`. v4 tetap sebagai varian tiga hari. Per `ATURAN-PROPOSAL.md` bagian 4 dan 8, klien menerima satu varian saja, jadi pilih salah satu sebelum dikirim.

**Rancangan.** 16 JPL (12 jam efektif), delapan modul 90 menit, empat modul per hari, 30 peserta, dua fasilitator, jalur kasus yang sama. Pemadatan dari rancangan 12 modul:
| Modul dua hari | Berasal dari modul tiga hari |
|---|---|
| 1 Gambaran ABK dan menghitung FTE | 1 dan 3 |
| 2 Daftar pekerjaan dan rencana pengumpulan data | 2 dan 4 |
| 3 Teknik dan praktik wawancara | 5 dan 6 |
| 4 Mengisi lembar kerja dan mencocokkan data | 7 |
| 5 Memeriksa mutu data | 8 |
| 6 Hitung ulang FTE, heatmap, dan akar masalah | 9 dan 10 (tanpa panel manajer) |
| 7 Tindak lanjut, uji RACI, dan panel manajer | 11 dan panel dari 10 |
| 8 Rencana 30-60-90 dan penutupan | 12 |

Porsi rancangan 23 persen penjelasan, 57 persen praktik, 20 persen review, dihitung skrip dari bobot waktu per langkah (jumlah 720 menit). Bagian 7.3 di dokumen masih menulis sekitar 25/55/20. Selisihnya ringan dan tidak diubah.

**Penyebutan durasi yang diganti di dokumen (12 tempat):** rekomendasi di ringkasan eksekutif, kalimat pembuka dan tabel 7.3, judul 7.4, level 1 dan rangkaian evaluasi di 7.5, kalimat penyesuaian luaran, jadwal pelaksanaan, komponen jasa pelatihan, ketentuan biaya, dan catatan praktik pengumpulan data. Tidak ada sisa "tiga hari", "3 hari", "24 JPL", atau "Hari 3" di dokumen.

**Perlu diputuskan tim**
1. **Harga.** Komponen pelatihan tetap Rp40.000.000, karena hitungan internalnya belum ada dan saya tidak mengarang potongan. Dua hari mengurangi honor dan perjalanan fasilitator, tetapi biaya persiapan kasus, kartu peran, dan lembar uji tetap sama. Tim menentukan angka akhir.
2. **Bagian yang dipadatkan.** Praktik wawancara tinggal tiga putaran singkat, dan modul 6 serta 7 paling padat. Bila fasilitator merasa kurang, tambahkan klinik daring pada hari ke-14 selain hari ke-30.
3. **Target Kirkpatrick.** Post-test minimal 75 dan peningkatan 20 persen tidak diubah. Dengan waktu lebih singkat, target level 2 sedikit lebih berat. Bagian 7.5 belum dipelajari ulang untuk dua hari selain nomor hari.
4. **Alasan menyarankan dua hari.** Peserta lebih mudah dilepas dari pekerjaan, biaya perjalanan fasilitator lebih rendah, dan keterampilan inti tetap tercakup. Kekurangannya: pendalaman pindah ke klinik D+30.
5. **Bagian 7.3 dan 7.6** tidak diubah selain angka durasi. Istilah Inggris di 7.3 (hidden work, annualisasi) masih ada.
