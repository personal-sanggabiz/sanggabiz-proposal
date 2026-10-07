# Assets — Bukti Kerja untuk Proposal Tanoto PPG

**Status semua dokumen di folder ini: DRAF/CONTOH, belum final.** Setiap file (Pedoman, Rubrik, Panduan FGD) punya keterangan status di sampul, header tiap halaman, dan footer: masih memerlukan masukan dan penyempurnaan lebih lanjut dari tim sebelum dipakai final, dan akan diselaraskan dengan Pedoman Pendampingan 2026 resmi begitu diterima dari Direktorat PPG saat kick-off.

Aset ini **sengaja tidak dimasukkan ke PPTX/PDF proposal utama**. Tujuannya menunjukkan kemampuan lewat hasil kerja nyata, bukan slide yang menjelaskan konsep — kamu yang memutuskan di mana dan bagaimana memakainya (disisipkan manual ke deck, dilampirkan terpisah, ditunjukkan langsung saat presentasi 5 Okt, dsb).

Semua data LPTK di dalamnya **ilustrasi/contoh** — diberi label jelas di tiap aset, konsisten dengan cara compro Sanggabiz sendiri menandai dashboard contoh dengan stempel "SAMPLE". Tidak ada klaim faktual soal kondisi LPTK tertentu.

## 1 — Dashboard Kesiapan Akreditasi (interaktif) — ada 2 versi
`1-dashboard/` — detail lengkap di `1-dashboard/README.md`
- **v2 (`index-v2.html`) — versi terbaru, "One Program, One Integrated Dashboard".** Master-detail: sidebar 12 LPTK + Ringkasan Program di kiri/atas, panel detail per LPTK mencakup **9 kategori**: baseline, hasil pendampingan, RTL Tracker (multi-item + PIC + tenggat), kemajuan akreditasi (baseline vs terkini + delta), pemetaan PPL 6 dimensi, Evidence Register, praktik baik, hambatan, rekomendasi. Gambar jadi: `dashboard-v2-ringkasan.png` dan `dashboard-v2-detail.png` (contoh Universitas Negeri Medan, semua 9 kategori dalam satu gambar).
- **v1 (`index.html`) — versi awal, single-screen ringkas.** Tabel 12 LPTK × 5 elemen + filter klaster + grafik rata-rata + status RTL ringkas (1 status, bukan daftar item). Tetap disimpan sebagai alternatif tampilan cepat. Gambar: `dashboard-screenshot.png` (3200×2012).
- Cara pakai saat presentasi live: jalankan `python -m http.server` di folder `assets`, buka `localhost:PORT/1-dashboard/index-v2.html` (v2, supaya klik-detail berfungsi) atau `index.html` (v1).

## 2 — Visual Sebelum vs Sesudah
`2-sebelum-sesudah/`
- **index.html** — sumber HTML (bisa diedit teksnya kalau perlu).
- **sebelum-sesudah.png** — gambar jadi 3200×1800 (rasio 16:9, pas ditempel sebagai slide penuh), berisi 5 baris perbandingan kondisi "sebelum" (bukti berserakan) vs "sesudah" (Evidence Register + Gap-to-Action Tracker jalan).

## 3 — Pedoman dan Instrumen Kerja Nyata
`3-instrumen/`
- **Pedoman-Pendampingan-Teknis.docx / .pdf** (8 halaman) — dokumen payung: metodologi 5 fase Evidence-Based Continuous Quality Improvement dipetakan ke RFP C.4, tata kelola dan batas peran, prosedur kerja langkah demi langkah tiap fase, indikator status RTL (rencana/berjalan/selesai/terbukti dengan definisi objektif), standar mutu luaran, daftar instrumen dan alat kerja (menaungi dua dokumen di bawah plus dashboard), dan daftar istilah (TF, DTPS, CPL, SPMI/PPEPP, dll). Ini yang menjawab Key Deliverable 1 di slide metodologi v2 ("Pedoman dan modul pendampingan teknis").
- **Rubrik-Observasi-PPL.docx / .pdf** — instrumen observasi 6 dimensi PPL (standar, peran, pendampingan, observasi, umpan balik, penilaian), skala 1–4, siap dipakai Koordinator Lapangan saat kunjungan. Produk Fase 02 Design, dipakai di Fase 03 Deliver.
- **Panduan-FGD-Pemetaan-PPL.docx / .pdf** — panduan FGD 4 kelompok informan (DPL, guru pamong, mahasiswa PPG, pimpinan prodi), pertanyaan disusun untuk langsung menjawab Temuan 2 & 3 Supervisi 2025 (pembelajaran berpusat mahasiswa; mutu praktikum dan umpan balik). Produk Fase 02 Design, dipakai di Fase 03 Deliver.
- **build_pedoman.py, build_instrumen.py** — skrip pembuat (python-docx); jalankan ulang kalau perlu mengubah isi lewat kode, atau edit langsung file .docx-nya. Convert ke PDF pakai LibreOffice (`soffice --headless --convert-to pdf nama.docx`).
- Ini persis luaran yang dijanjikan di Laporan Awal ("pedoman dan modul pendampingan", "instrumen FGD terkait PPL") — jadi bukan konsep baru, tapi bukti bahwa luaran itu sudah bisa dibuat, bukan sekadar janji.

## 4 — Slide 8 v2: Metodologi (Evidence-Based Continuous Quality Improvement)
`4-slide8-metodologi-v2/`
- **Slide 8 v2 - Metodologi (Evidence-Based CQI).pptx** — satu slide berdiri sendiri, siap diimpor manual ke Google Slides (bukan pengganti otomatis slide 8 di deck lokal — deck lokal tetap versi PDCA/PPEPP 4-fase).
- **Slide 8 v2 - Metodologi (Evidence-Based CQI).png** — versi gambar 3200×1800.
- Isi: 5 fase (Diagnose–Design–Deliver–Monitor & Verify–Evaluate & Improve, dipetakan 1:1 ke RFP C.4) + 9 Key Deliverables. Detail dasar metodologi dan riwayat keputusan ada di `4-slide8-metodologi-v2/README.md`.

## 5 — Slide "Pedoman dan Instrumen" (pasangan kartu Strategi Keberlanjutan)
`5-slide-pedoman-instrumen/`
- **Slide - Pedoman dan Instrumen.pptx** — satu slide berdiri sendiri, siap diimpor manual ke Google Slides. Pasangan dari slide "Repositori dan dashboard" yang sudah kamu buat (kartu Strategi Keberlanjutan yang sama).
- **Slide - Pedoman dan Instrumen.png** — versi gambar 3200×1800.
- Isi: 3 kartu (Pedoman Pendampingan Teknis, Rubrik Observasi PPL, Panduan FGD Pemetaan PPL), tiap kartu menampilkan cuplikan halaman isi dokumen (bukan sampul) + tombol "Buka di Google Drive" (hyperlink aktif) + teks link lengkap di bawahnya (juga hyperlink aktif, jadi kelihatan tanpa perlu klik dan tetap bisa langsung dibuka).
- Link yang dipakai (dari Google Drive milik user, bukan yang di-generate Claude):
  - Pedoman Pendampingan: https://drive.google.com/file/d/1ha5-qChCFAWESvPBs_YCIk1wztmZecn7/view
  - Rubrik Observasi: https://drive.google.com/file/d/1IoTR5Qby_PAVQTZDJweBhtjcmWByYDcr/view
  - Panduan FGD: https://drive.google.com/file/d/1QONJRSzPWMIEqa3gu9igzXVSQa41UhnV/view

## 6 — Playbook Akreditasi-Ready (contoh ringkas)
`6-playbook-akreditasi-ready/`
- **Playbook-Akreditasi-Ready.docx / .pdf** (4 halaman) — menjawab janji kartu "Replikasi ke LPTK lain" di Strategi Keberlanjutan: "Playbook ringkas Akreditasi-Ready dan rekomendasi tahapan untuk 172 LPTK." **Beda tujuan dari Pedoman Pendampingan Teknis** (yang itu SOP internal tim untuk 12 LPTK) — playbook ini produk diseminasi untuk 172 LPTK secara nasional, ditulis sebagai panduan mandiri (self-service), bukan prosedur kerja tim.
- Isi (sengaja dipersingkat, cuma menunjukkan struktur): tentang playbook dan cara pakai, 5 langkah menuju Akreditasi Unggul, Checklist Kesiapan Akreditasi Unggul (asesmen mandiri per elemen), pelajaran dari program 2026, langkah selanjutnya, referensi.
- **build_playbook.py** — skrip pembuat. Status draf/contoh ditulis eksplisit di sampul, header, dan footer, sama seperti dokumen lain.
- Versi lengkap (bukan versi ringkas 4 halaman ini) baru disusun di Fase 05 Evaluate & Improve program, setelah pembelajaran dari 12 LPTK benar-benar terkumpul — dokumen ini cuma menunjukkan bentuknya nanti.

## 7 — Slide "Replikasi ke LPTK Lain" (pasangan kartu Strategi Keberlanjutan ke-4)
`7-slide-playbook/`
- **Slide - Replikasi ke LPTK Lain.pptx** — satu slide berdiri sendiri, siap diimpor manual ke Google Slides. Melengkapi 3 slide Strategi Keberlanjutan lainnya (Repositori dan dashboard; Pedoman dan instrumen).
- **Slide - Replikasi ke LPTK Lain.png** — versi gambar 3200×1800.
- Isi: 2 cuplikan halaman Playbook Akreditasi-Ready (Tentang & 5 Langkah; Checklist Kesiapan) + 3 chip fakta (4 halaman, contoh ringkas, untuk 172 LPTK) + tombol dan teks link Google Drive (hyperlink aktif) ke Playbook di `6-playbook-akreditasi-ready/`.

Dengan ini, **3 dari 4 kartu Strategi Keberlanjutan sudah punya slide bukti**: Repositori dan dashboard (dibuat user sendiri + Netlify), Pedoman dan instrumen (folder 5), Replikasi ke LPTK lain (folder 7). Yang belum: **Komunitas praktik antar-LPTK**.

## Rekomendasi pemakaian
- **Dashboard + sebelum/sesudah**: cocok disisipkan di bagian "Pendekatan teknis" atau ditunjukkan sebagai demo hidup saat presentasi Zoom 5 Oktober — jauh lebih berkesan dibuka langsung di browser daripada screenshot statis.
- **Instrumen**: cocok jadi Lampiran Instrumen di proposal teknis (di luar 9 poin wajib Lampiran 1), atau ditunjukkan sepintas saat presentasi sebagai bukti "sudah ada, tinggal dipakai."
- Kalau salah satu mau dipublikasikan sebagai link (misal dashboard, biar bisa diklik langsung saat Zoom tanpa share-screen file lokal), bilang saja — bisa dibuatkan artifact terpisah dengan sekali klik.
