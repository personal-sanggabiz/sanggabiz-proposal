# CATATAN TIM: Memakai Proposal Siap Pakai

File: `proposal/Proposal Workshop Pembelajaran Berdiferensiasi - SLB dan Inklusif Balikpapan.docx` (portrait A4, 15 halaman termasuk sampul, surat pengantar, dan lampiran).

## 1. Isian untuk tim (stabilo kuning di dokumen)
Semua bagian ini ditandai `[ISI TIM: ...]`. Tidak ada bagian lain yang kosong.
| Halaman | Isian |
|---|---|
| Surat pengantar | Nomor surat, nama dan jabatan penerima, instansi penerima, tanda tangan, nama dan jabatan penandatangan, kontak |
| Bab 2.1 | Nama hotel terpilih |
| Bab 2.2 | Tanggal pelaksanaan |
| Bab 6 | Data lengkap Pemateri 1 dan Pemateri 2 (nama, gelar, jabatan, keahlian, pendidikan, pengalaman, foto) |
| Bab 7 | Kontak penghubung |

Cara mencari semua isian: di Word, tekan Ctrl+F lalu cari `ISI TIM`.

## 2. Yang perlu dicek sebelum dikirim
Angka dan klaim di bawah berasal dari ringkasan hasil pencarian web, sebagian belum dibuka di sumber aslinya karena akses jaringan terbatas. Buka tautan di Lampiran C dan cocokkan.
1. 40.164 satuan pendidikan dengan peserta didik berkebutuhan khusus, 5.956 (14,83 persen) punya guru pembimbing khusus, per Desember 2023 (rujukan 1).
2. SLB Negeri Balikpapan: berdiri 1984, 401 siswa (219 SD, 96 SMP, 86 SMA). Tahun berita belum dipastikan (rujukan 8).
3. 11 SLB di bawah Pemprov Kaltim (rujukan 7).
4. Nomor Peraturan Wali Kota tentang pendidikan inklusif (disebut Nomor 3 Tahun 2013). Cek teks resmi di JDIH Balikpapan (rujukan 9).
5. Pernyataan Komisi IV DPRD Kaltim tentang kurangnya guru PLB (rujukan 6).
6. Klaim Sanggabiz di bab 1.5: 400+ proyek, 500+ pakar dan trainer, 6.000+ mitra, pelatihan Rumah BUMN di lebih dari sembilan kota termasuk Samarinda, daftar klien. Cocokkan dengan data terbaru perusahaan.
7. Foto di bab 1.5 diambil dari arsip compro. Pastikan boleh dipakai untuk klien ini.
8. Harga hotel: minta penawaran tertulis (Quest, Aston, Swiss-Belhotel) dengan permintaan yang sama, lalu perbarui bab 5 bila berbeda. Lihat `07-RISET-TEMPAT-HARGA.md`.
9. Biaya manajemen 10 persen dan PPN 11 persen adalah asumsi. Sesuaikan dengan kebijakan dan status pajak Sanggabiz.
10. Honor pemateri dan fasilitator adalah perkiraan. Sesuaikan setelah pemateri ditetapkan.
11. Target indikator keberhasilan (naik 20 persen, 80 persen menghasilkan draf, kepuasan 4 dari 5, 75 persen bergabung grup) adalah usulan. Ubah bila perlu.
12. Kuota peserta seimbang antara SLB dan sekolah inklusif adalah usulan. Jumlah guru SLB di Balikpapan belum diketahui, jadi 80 orang perlu dicek terhadap data Dapodik.

## 3. Mengubah angka anggaran
Semua angka ada di bagian `paket` pada `_build/build_proposal.js`. Ubah, lalu jalankan:
```
cd slb-kaltim/_build
NODE_PATH=$(npm root -g) node build_proposal.js
```
Atau edit langsung di Word. Tabel anggaran tidak terhubung rumus, jadi ubah subtotal dan total secara manual bila mengedit di Word.

## 4. Font
Dokumen memakai Montserrat. Pasang font dari `../tanoto/fonts/` di komputer yang membuka dokumen. Tanpa font itu, Word memakai font pengganti dan tata letak bisa bergeser sedikit.

## 5. Sebelum dikirim
- Jalankan `python3 tools/cek-gaya.py slb-kaltim/proposal` dari akar repo.
- Hapus semua stabilo kuning dan simpan sebagai PDF untuk dikirim.
