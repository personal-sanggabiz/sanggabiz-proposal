# Panduan untuk Claude di repo ini

Repo ini berisi proposal dan dokumen kerja Sanggabiz.

1. Sebelum menulis apa pun (dokumen, README, laporan, pesan commit, deskripsi PR, jawaban di chat), baca `ATURAN-PENULISAN.md` dan ikuti. Aturan itu berlaku juga untuk balasan Claude kepada pengguna.
2. Untuk proposal, baca juga `ATURAN-PROPOSAL.md` (metode, silabus, anggaran, tempat, bahasa) dan `FORMAT-STANDAR-PROPOSAL.md` (bentuk halaman standar).
3. Sebelum commit, jalankan `python3 tools/cek-gaya.py` pada file yang ditulis atau diubah. Perbaiki semua temuan, kecuali di kutipan dan salinan sumber.
4. Setiap proyek proposal punya foldernya sendiri (`tanoto/`, `slb-kaltim/`). Jangan mencampur aturan atau aset antarproyek.
5. Kutipan chat dan dokumen sumber ditulis apa adanya. Tafsiran dipisah dan diberi label.
6. Bagian yang harus diisi tim ditandai `[ISI TIM: ...]` dengan stabilo kuning di dokumen.
7. Versi lama dokumen tidak dihapus saat membuat versi baru (lihat `ATURAN-PROPOSAL.md` bagian 8).
8. Dokumen untuk klien tidak memuat catatan atau kata internal, dan ditulis agar terbaca orang awam (lihat `ATURAN-PROPOSAL.md` bagian 6 dan 9). Catatan internal masuk ke catatan tim.
