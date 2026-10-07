# Dashboard Kesiapan Akreditasi — 12 LPTK

Dua versi ada di folder ini. **v1 tidak dihapus/ditimpa** — v2 adalah pengembangan terpisah.

## v2 — Dashboard Terintegrasi (`index-v2.html`) — versi terbaru, dipakai duluan
"One Program, One Integrated Dashboard": satu sistem yang mencakup **9 kategori sekaligus**, per LPTK:
1. **Baseline** — skor kondisi awal (Supervisi 2025)
2. **Hasil pendampingan** — ringkasan naratif tiap kunjungan
3. **RTL** — RTL Tracker (Gap-to-Action): 3 item per LPTK, dengan PIC, tenggat, status
4. **Kemajuan akreditasi** — baseline vs skor terkini, dengan indikator delta (▲)
5. **Pemetaan PPL** — 6 dimensi (standar, peran, pendampingan, observasi, umpan balik, penilaian), sesuai Rubrik Observasi PPL di folder `3-instrumen/`
6. **Bukti kemajuan** — Evidence Register per LPTK (status Terverifikasi/Diklaim/Belum ada)
7. **Praktik baik**
8. **Hambatan**
9. **Rekomendasi** (dengan tag prioritas)

**Arsitektur**: master-detail. Sidebar kiri = daftar 12 LPTK + filter klaster + tombol "Ringkasan Program". Panel kanan = Ringkasan Program (agregat 12 LPTK, mirip tampilan v1) ATAU detail satu LPTK (6 bagian di atas), tergantung yang diklik.

**File:**
- `index-v2.html` — interaktif, jalankan lewat server lokal (`python -m http.server` di folder `assets`, buka `localhost:PORT/1-dashboard/index-v2.html`) supaya klik-detail berfungsi. Dobel-klik langsung juga jalan tapi sebagian browser membatasi file lokal.
- `dashboard-v2-ringkasan.png` — ekspor gambar tampilan Ringkasan Program (3200×1860).
- `dashboard-v2-detail.png` — ekspor gambar tampilan detail salah satu LPTK, contoh: Universitas Negeri Medan (3200×2900, memuat semua 9 kategori dalam satu gambar).

**Cara screenshot LPTK lain**: buka `index-v2.html#detail=NAMA%20LPTK&full=1` (nama LPTK di-URL-encode, `&full=1` mematikan scroll internal supaya semua konten kelihatan dalam satu gambar penuh) lalu screenshot / render lewat headless Chrome:
```bash
chrome.exe --headless --disable-gpu --hide-scrollbars --force-device-scale-factor=2 --window-size=1600,1450 --screenshot="output.png" "http://localhost:PORT/1-dashboard/index-v2.html#detail=Universitas%20Riau&full=1"
```

## v1 — Dashboard ringkas (`index.html`) — versi awal, single-screen
Satu layar: tabel 12 LPTK × 5 elemen (skor 1–4) + filter klaster + grafik rata-rata + status RTL ringkas (1 status per LPTK, bukan daftar item). Cocok untuk tampilan cepat/overview saja — tidak mencakup PPL 6-dimensi, Evidence Register, praktik baik/hambatan/rekomendasi.
- `index.html`, `dashboard-screenshot.png` (3200×2012).

## Semua data ilustrasi
Skor, status RTL, isi Evidence Register, praktik baik, hambatan, dan rekomendasi di **kedua versi** adalah **contoh/ilustrasi** yang dihasilkan dari kombinasi pola umum (bukan temuan riil terhadap LPTK tertentu) — diberi label jelas di tiap halaman. Diisi ulang dengan data sebenarnya begitu Laporan Awal dan kunjungan pertama berjalan.
