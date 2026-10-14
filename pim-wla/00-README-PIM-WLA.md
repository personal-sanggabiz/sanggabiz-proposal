# Proposal ABK Fungsi Umum dan Pelatihan, PT Pupuk Iskandar Muda

> Aturan: `../ATURAN-PENULISAN.md`, `../ATURAN-PROPOSAL.md`, `../FORMAT-STANDAR-PROPOSAL.md`. Cek dengan `python3 ../tools/cek-gaya.py pim-wla/proposal`.
>
> Proyek terpisah dari `../tanoto/` dan `../slb-kaltim/`. Jangan mencampur aturan, aset, atau klaim.

Dibuat: 14 Oktober 2026.

## Isi folder
| Berkas | Fungsi |
|---|---|
| `proposal/v2 Proposal ABK Fungsi Umum dan Pelatihan - PT Pupuk Iskandar Muda.docx` | Proposal siap pakai (A4, 29 halaman) |
| `proposal/v3 Proposal Analisis Beban Kerja Fungsi Umum dan RACI Enterprise - PT Pupuk Iskandar Muda.docx` | Dasar upgrade berikutnya: format dan isi sama dengan v1 (29 halaman), hanya nama dan logo Levner diganti Sanggabiz. Dibuat dengan `_build/rebrand_v1_ke_sanggabiz.py` |
| `proposal/v4 Proposal Analisis Beban Kerja Fungsi Umum dan RACI Enterprise - PT Pupuk Iskandar Muda.docx` | Proposal lengkap 31 halaman: desain dan isi sama dengan v3, hanya bagian 7.4 yang diganti kurikulum baru (12 modul, tabel enam kolom). Dibuat dengan `_build/susun_proposal_v4.py`. v3 tetap disimpan sebagai arsip |
| `proposal/v5 Proposal Analisis Beban Kerja Fungsi Umum dan RACI Enterprise Pelatihan 2 Hari - PT Pupuk Iskandar Muda.docx` | Varian pelatihan dua hari (16 JPL, delapan modul): proposal lengkap 31 halaman, desain sama dengan v3 dan v4. Dibuat dengan `_build/susun_proposal_v5_2hari.py`. v4 (tiga hari) tetap berlaku sebagai varian lain |
| `proposal/v1 Proposal Kurikulum Pelatihan ABK 2 Hari - PT Pupuk Iskandar Muda.pptx` | Deck proposal kurikulum dua hari (12 slide, gaya compro Sanggabiz: navy, kuning, teal, font Arial). Dibuat dengan `_build/build_kurikulum_deck.js`. PDF pratinjau di `preview-deck/` |
| `kurikulum/v3 Kurikulum Pelatihan Tiga Hari ABK - PT Pupuk Iskandar Muda.docx` dan `.pdf` | Lampiran siap kirim yang menggantikan bagian 7.4 pada v3: 12 modul, 24 JPL, satu kasus berjalan, tabel enam kolom tanpa menit per langkah, bahasa awam, tanpa sampul. v1 dan v2 disimpan sebagai arsip, skripnya `_build/build_kurikulum_v1.js` dan `_build/build_kurikulum_v2.js`. Skrip terbaru `_build/build_kurikulum.js` |
| `09-CATATAN-KURIKULUM.md` | Catatan tim untuk kurikulum: perubahan dari v3, keputusan yang perlu dikonfirmasi, status verifikasi rujukan |
| `08-CATATAN-TIM.md` | Mulai dari sini: perubahan dari v1, keputusan yang perlu dikonfirmasi, isian tim, daftar cek klaim |
| `sumber/v1 ...docx` | Dokumen asli yang diterima (atas nama Levner), disimpan sebagai arsip v1 |
| `preview/` | PDF dan gambar tiap halaman v2 |
| `preview-v4/` | PDF pratinjau v4 |
| `preview-v5/` | PDF pratinjau v5 (dua hari) |
| `preview-v3/` | PDF pratinjau v3. Render LibreOffice menumpuk teks sampul pada gambar, di Word tampil normal |
| `_build/build_proposal.js` | Skrip pembuat v2. `NODE_PATH=$(npm root -g) node build_proposal.js` |
