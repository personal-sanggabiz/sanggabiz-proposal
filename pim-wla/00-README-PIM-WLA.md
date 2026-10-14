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
| `kurikulum/v1 Kurikulum Pelatihan Tiga Hari ABK - PT Pupuk Iskandar Muda.docx` dan `.pdf` | Usulan pengganti bagian 7.4 pada v3 (12 modul, 24 JPL, satu kasus berjalan, format standar tipe 1). Dibuat dengan `_build/build_kurikulum.js` |
| `09-CATATAN-KURIKULUM.md` | Catatan tim untuk kurikulum: perubahan dari v3, keputusan yang perlu dikonfirmasi, status verifikasi rujukan |
| `08-CATATAN-TIM.md` | Mulai dari sini: perubahan dari v1, keputusan yang perlu dikonfirmasi, isian tim, daftar cek klaim |
| `sumber/v1 ...docx` | Dokumen asli yang diterima (atas nama Levner), disimpan sebagai arsip v1 |
| `preview/` | PDF dan gambar tiap halaman v2 |
| `preview-v3/` | PDF pratinjau v3. Render LibreOffice menumpuk teks sampul pada gambar, di Word tampil normal |
| `_build/build_proposal.js` | Skrip pembuat v2. `NODE_PATH=$(npm root -g) node build_proposal.js` |
