# ATURAN GAYA SANGGABIZ — MODE 2: PRESENTASI (LANDSCAPE, 16:9)

Berlaku untuk **proposal berbentuk slide**: proposal teknis yang dipresentasikan (Zoom atau tatap muka), deck penawaran, company profile, laporan ringkas berbentuk PPT. Bentuknya seperti PowerPoint, diekspor ke PDF untuk dikirim.

Sumber aturan ini:

| Sumber | Dipakai untuk |
|---|---|
| `reference/Company Profile - Training Portfolio Sanggabiz 2026 (2).pptx` (39 slide, 20 × 11,25 in) | Sumber utama gaya: warna, font, komponen, pola tata letak |
| `tanoto/01-ANALISIS-GAYA-Compro-Sanggabiz.md` | Pengukuran warna dan font dari XML compro (diverifikasi ulang dengan render sampel: sampul, kartu 3 kolom, studi kasus, penutup) |
| `tanoto/_build/lib.py` dan `build_deck.py` | Skala ukuran yang sudah dipakai di deck proposal Tanoto 13,33 × 7,5 in |

Mode 2 **gelap** (navy + kuning, font Montserrat). Mode 1 (dokumen) **terang** dengan font Nunito + Oswald. Jangan dicampur dalam satu proposal. Yang boleh dipakai bersama: logo, foto kegiatan, angka portofolio, dan data kontak.

---

## 0. Kapan memakai mode ini

| Kebutuhan | Pakai |
|---|---|
| Dipresentasikan di depan penilai atau klien | **Mode 2** |
| Klien meminta PPT atau PDF landscape | **Mode 2** |
| Dibaca sendiri, teks panjang, banyak tabel dan sitasi | Mode 1 (Dokumen) |
| RFP/TOR menentukan struktur | Struktur klien menang. Gaya visual tetap Mode 2. |

Proposal yang dinilai bukan pitch compro. Compro dibuat untuk memukau. Proposal dibuat untuk **dinilai dan ditelusuri**. Karena itu aturan di bawah lebih tenang dan lebih terbaca daripada compro (bagian 9).

---

## 1. Kesan yang dibangun

Korporat-modern, gelap, percaya diri, padat data: latar navy pekat, aksen kuning emas, teks putih besar dan tebal, kartu bersudut membulat, foto kegiatan sungguhan, angka besar. Pesan: *tim yang sudah sering turun ke lapangan dan hasilnya terukur.*

Tidak ada: warna pastel, gradasi ramai, ornamen dekoratif, clip art berwarna, serif di isi (serif hanya di slide pembatas foto).

---

## 2. Ukuran slide

| Hal | Nilai |
|---|---|
| Rasio | **16:9** |
| Proposal | **13,333 × 7,5 in** (PPT standar). Dipakai agar hasil PDF ringan dan ukuran teks mudah dihitung. |
| Compro (acuan saja) | 20 × 11,25 in. Ukuran font compro ×0,667 = ukuran di slide 13,33 in. |
| Margin kiri-kanan | **0,65 in**, lebar konten 12,03 in |
| Judul mulai | y ≈ 0,42 in (kicker) dan 0,72 in (judul) |
| Footer | y ≈ 7,06 in: logo kecil kiri, judul dokumen dan nomor RFP di sebelahnya (9 pt abu biru), nomor halaman kuning kanan |
| Zona aman | Tidak ada teks di luar margin. Logo boleh masuk margin. |

---

## 3. Warna

| Peran | Hex | Dipakai untuk |
|---|---|---|
| **Navy terang** | `#0B396D` | Awal gradasi latar |
| **Navy gelap** | `#092647` | Akhir gradasi latar, teks di atas kuning (`#0B2748` setara) |
| **Navy panel** | `#0B3361` | Panel isi kanan studi kasus |
| **Navy dalam** | `#0D2A4E` | Panel pesan kunci |
| **Kuning emas** | `#FFC00C` | Tab judul, kicker, huruf besar, nomor halaman, bingkai pesan kunci |
| **Teal kartu** | `#2F7E9D` | Isi kartu 3 kolom, kotak kerangka |
| **Putih** | `#FFFFFF` | Semua teks di atas navy, ikon garis, bingkai kartu |
| **Abu biru** | `#B0C0CF` | Teks sekunder, footer |
| **Biru aksi** | `#2446BC` / `#0071CE` | Hemat: kotak penekanan, tautan |
| **Tinta** | `#1D1D1F` | Teks di slide berlatar putih |

**Latar**: gradasi linear vertikal **`#0B396D` (atas/kiri) → `#092647` (bawah/kanan)**, sudut ±160°. Overlay foto: `#092647` opasitas ±75%.

**Proporsi**: ±70% navy, ±15% putih, ±8% kuning, ±5% teal, sisanya foto. Kuning selalu **sedikit dan bertujuan**: label, kata kunci, nomor. Jangan jadi latar besar kecuali satu panel sorotan.

**Pasangan kontras yang sah**

| Teks | Latar |
|---|---|
| Putih | Navy |
| Navy `#0B2748` | Kuning `#FFC00C` (tab, pil) |
| Kuning | Navy (kicker, huruf besar) |
| Putih | Teal `#2F7E9D` (isi kartu) |
| Tinta `#1D1D1F` | Putih (slide pembatas) |

**Warna klien**: jangan dipakai sebagai warna utama. Logo klien muncul di atas **kapsul atau panel putih** supaya warnanya tidak bentrok. (Contoh: hijau Tanoto Foundation tidak dipakai sebagai warna slide.)

---

## 4. Tipografi

### 4.1 Font

| Peran | Font | Catatan |
|---|---|---|
| Judul, sub-judul, angka | **Montserrat** ExtraBold / Bold / SemiBold | Pengganti Mont (font asli compro, tertanam). Mont dipakai bila berlisensi. |
| Isi, bullet, tabel | **Montserrat** Medium / Regular | |
| Judul slide pembatas foto | **Playfair Display** Bold (serif), kapital | Hanya di slide pembatas berlatar putih |
| Aksen sangat kecil | Poppins, Nunito | Jarang. Boleh ditinggalkan. |

File font ada di `tanoto/fonts/` (Montserrat Regular, Medium, SemiBold, Bold, ExtraBold). **Pasang sebelum membuka PPTX atau mengekspor PDF**, kalau tidak PDF akan berganti font.

### 4.2 Skala ukuran (di slide 13,33 in)

Angka ini adalah yang sudah dipakai di deck Tanoto. Ukuran compro ditaruh di kolom kanan sebagai acuan.

| Elemen | Ukuran | Gaya | Compro (20 in) |
|---|---|---|---|
| Judul sampul | 38–44 pt | ExtraBold, putih, kiri | 50–56 pt |
| **Judul slide** | **28 pt** | ExtraBold, putih, Title Case, kiri | 30–40 pt |
| **Kicker** (di atas judul) | **11,5 pt** | Bold, **KAPITAL, kuning**, jarak huruf +120 | — |
| Sub-judul / label bagian | 15–16 pt | Bold, kuning | 24–28 pt |
| **Isi / paragraf** | **13,5–14 pt** | Medium, putih, spasi 1,1, jarak paragraf 9 pt | 12–19 pt |
| Judul kartu / tab | 13–13,5 pt | Bold | — |
| **Tabel** | **11–12 pt** (batas bawah 10,8 pt) | Header bold | — |
| Angka statistik | 40–60 pt | ExtraBold, putih | 30–50 pt |
| Catatan sumber | 10–11 pt | Regular, abu biru | — |
| Footer | 9 pt | Regular, abu biru | — |
| Nomor halaman | 10 pt | Bold, kuning, kanan | — |
| Judul slide pembatas (serif) | 40–44 pt | Playfair Bold, kapital, tinta | ±50 pt |

**Batas keras**: isi ≥ 12 pt, tabel ≥ 10,8 pt, tidak ada teks isi di bawah 9 pt. Compro memakai teks yang jauh lebih kecil. Itu tidak boleh ditiru di proposal yang dibaca penilai.

### 4.3 Hierarki dan perataan

- **Tiga level per slide**: judul → sub-judul (kapital atau kuning) → isi.
- **Paragraf panjang rata kiri.** Compro memakai justify, tetapi di slide teks padat justify menghasilkan celah antar kata. Justify hanya untuk blok pendek di kartu sempit yang sudah diuji.
- Judul rata kiri. Rata tengah hanya di slide angka besar, penutup, dan pembatas.
- Bullet titik kecil putih. Tidak lebih dari 5 bullet per kartu, tidak lebih dari 12 baris per slide.
- Kapital: kicker, tab, dan sub-judul daftar ("LIST OF MATERIALS:"). Judul slide Title Case.

---

## 5. Komponen

| Komponen | Spesifikasi |
|---|---|
| **Kartu** | Sudut membulat besar (radius ±0,28 in di slide 13,33 in; ±24 px pada 1280 px), isi **teal `#2F7E9D`** atau navy tengah, **bingkai putih 2–2,25 pt**. Dipakai untuk 3 kolom. |
| **Tab judul** | Kapsul **kuning `#FFC00C`**, teks navy `#0B2748` tebal 13 pt, menempel di tepi atas kartu. Varian: kapsul biru terang dengan teks putih untuk label slide ("Our Services", "Framework"). |
| **Pil kontak/tag** | Kapsul garis putih, isi transparan: situs, Instagram, email. |
| **Panel pesan kunci** | Navy dalam `#0D2A4E`, **bingkai kuning 1,5 pt**, sudut membulat, satu atau dua kalimat putih di dalamnya. Dipakai di dasar slide sebagai kesimpulan. |
| **Ikon** | **Garis putih monolinier**, seragam, tanpa isi (garis ±2 px). Ikon kuning boleh untuk penekanan tunggal. Tidak ada ikon berwarna-warni. |
| **Angka besar** | Ikon garis putih + angka besar putih + label kecil tebal di bawah. Grid 3×2 (slide "Sanggabiz in Numbers"). Format `6.000+`. |
| **Huruf besar mnemonik** | Huruf kuning besar (P-L-A-N) di atas label putih, dalam urutan horizontal 4 langkah. |
| **Logo** | `sanggabiz` putih. Kiri atas (sampul, kartu) atau kiri bawah (studi kasus). **Watermark** logo besar di panel kanan, ±5% opasitas. Berkas: `tanoto/_build/assets/logo_white.png`. |
| **Judul vertikal** | Teks diputar 90° di strip kiri, putih atau navy ("DIGITAL TRAINING PROJECT", "OUR ONGOING PROJECT"). |
| **Foto** | **Foto kegiatan sungguhan**. Sampul memakai potongan foto miring dengan sudut membulat. Foto lain persegi atau membulat, kadang dengan overlay navy 75%. |
| **Tabel** | Header **kuning** dengan teks navy bold. Baris isi bergantian navy panel `#0B3361` dan navy dalam `#0D2A4E`, teks putih, kolom pertama bold. Garis teal 0,75 pt. Rata kiri, angka rata kanan atau tengah. |

---

## 6. Delapan pola tata letak

Compro memakai pola ini berulang. Tiap slide proposal harus jatuh ke salah satu pola, bukan dibuat dari nol.

| # | Pola | Struktur | Kapan dipakai |
|---|---|---|---|
| **A** | **Sampul** | Judul besar putih kiri. Logo kiri atas. **Kolase foto miring** di kanan. Pil statistik di bawah kiri ("100+ Great Organisations…"). Untuk proposal: tambah baris "Disiapkan untuk: [klien]" (logo klien di kapsul putih), nomor RFP, tanggal. | Slide 1 |
| **B** | **Penjelasan + diagram** | Kiri: kotak/diagram atau kartu teal. Kanan: judul tebal + paragraf + pil tag. Logo di pojok. | Ringkasan, pendekatan, kerangka, data/dashboard |
| **C** | **Tiga kolom kartu** | Judul di atas. 3 kartu teal, masing-masing dengan **tab kuning**, 4–5 baris berikon. | Peran, kapabilitas, risiko, tim |
| **D** | **Grid angka** | Judul tengah. 6 ikon + angka besar dalam 3×2. Atau 4 langkah horizontal bergaya P-L-A-N. | Konteks, profil Sanggabiz, linimasa |
| **E** | **Studi kasus** (paling sering di compro) | **Kiri**: foto kegiatan penuh + overlay navy, judul "Our Training Portfolio / [Klien]", empat baris berikon: **Tanggal · Jenis · Jumlah peserta · Topik**. **Kanan**: panel navy gelap, logo klien kanan atas, **"LIST OF MATERIALS:"** (kapital) + bullet, logo Sanggabiz kiri bawah. | Pengalaman relevan, deskripsi per tujuan |
| **F** | **Pembatas galeri foto** | Latar **putih**, strip navy kiri dengan logo diputar, kolase foto sudut kotak, **judul serif hitam besar**, logo klien di bawah. | Pemisah antar bab, bukti visual |
| **G** | **Kutipan / testimoni** | 5 kartu dengan tab kuning (nama klien), isi kutipan kecil, logo klien di atas. Baris bawah panel putih berisi logo. | Rekam jejak, referensi |
| **H** | **Penutup** | Judul putih tebal ("Let's Get on the Road with"), logo besar, **3 pil kontak**, alamat kecil. | Slide terakhir |

**Ritme**: slide teks dan slide foto berselang-seling. Setiap 3–5 slide isi ada satu slide bukti visual (foto, tangkapan layar, angka).

---

## 7. Kerangka isi proposal (default)

Urutan ini bisa dipangkas. Kalau ada RFP, RFP menentukan urutan dan istilah, lalu pola dipilih dari bagian 6. Jumlah slide: **20–24**.

| # | Slide | Isi | Pola |
|---|---|---|---|
| 1 | Sampul | Judul proposal, nomor RFP, klien, tanggal | A |
| 2 | Ringkasan Eksekutif | 3–4 paragraf. Kegiatan, hasil, elemen inovatif, nilai/dana. Satu slide. | B (teks kiri, angka kunci kanan) |
| 3 | Pemahaman konteks | Angka kunci masalah klien | D + B |
| 4 | Peran dan pembagian tugas | Klien · mitra lain · Sanggabiz | C |
| 5 | Kerangka pendekatan | **Define – Design – Deliver – Discuss** (kerangka 4 kata Sanggabiz), dipetakan ke fase proyek | B |
| 6–9 | Deskripsi per tujuan | Satu slide per tujuan: kegiatan, indikator, bukti | E |
| 10+ | Metodologi, data, rute, risiko | Sesuai kebutuhan | B, C |
| 14 | Linimasa dan deliverables | Fase, tanggal, luaran | D (4 langkah horizontal) |
| 16 | Tata kelola | Organogram, ToR, alur koordinasi | B |
| 17 | Keberlanjutan | | B |
| 18 | Anggaran ringkas dan termin | Angka besar + tabel | D + tabel |
| 19 | Profil Sanggabiz | Angka besar yang relevan saja | D |
| 20–22 | **Pengalaman relevan** | 3 studi kasus terpilih | **E** |
| 23 | Tim dan CV ringkas | | C |
| 24 | Penutup | Kontak | H |

**Slide pembuka bab (opsional)** memakai pola F.

**Formula pengalaman relevan** (pola E): *kapan · siapa · berapa peserta · apa topiknya · apa isi materinya.* Contoh: "April 2026 / Individual from SME / 110 Participants / Financial", disusul tiga bullet materi.

---

## 8. Gaya tulis (suara)

- **Pendek, tegas, bernomor.** "6.000+", "400+", "1000+ participants", "28+ Gov Institution".
- **Satu ide per slide.** Judul slide adalah klaim atau jawaban, bukan label kosong. Mis. "Hasil dapat ditelusuri sampai ke bukti" lebih baik dari "Metodologi".
- **Kerangka bernama dan mudah diingat**: Define–Design–Deliver–Discuss; P-L-A-N (Prepare, Learn, Apply, Navigate). Compro suka akronim dan urutan.
- **Judul bagian mengikuti istilah RFP persis**, dengan **kicker kuning kecil** sebagai penanda kriteria penilaian (mis. `PENDEKATAN TEKNIS · 50%`). Memudahkan penilai menemukan bagiannya.
- **Campuran Inggris–Indonesia** boleh bila sesuai kosa kata klien (Timeline, Governance, Steering Committee, Deliverables). Selebihnya bahasa Indonesia formal.
- Klaim kuat boleh bila ada angka dan bukti ("menghemat hingga 80% biaya tenaga kerja"). Tanpa angka, jangan.
- **Bukti sosial**: logo klien, kutipan klien bernama, hanya dengan izin.
- Opsional: satu baris "**Mengapa ini penting bagi [klien]**" di akhir bab. Jangan berlebihan.
- Format: Rupiah `Rp510.000.000` atau `Rp510 jt`, tanggal `7 Desember 2026`, persen `50%`.

---

## 9. Perbedaan dengan compro (dibuat lebih tenang untuk dinilai)

| Compro | Proposal |
|---|---|
| Teks 12–14 pt di slide 20 in (≈8–9 pt di 13,33 in) | Isi ≥ 13,5 pt, batas keras 12 pt |
| Justify di mana-mana | Rata kiri untuk paragraf panjang |
| Pitch: klaim dan logo | Metodologi: langkah, bukti, penanggung jawab, tanggal |
| Kuning di banyak tempat | Kuning hanya untuk label dan kata kunci |
| Banyak foto | Foto hanya yang boleh dipublikasikan dan relevan |
| Slide padat | Satu ide per slide, maksimum 12 baris |

---

## 10. Aturan khusus

1. **Identitas tetap Sanggabiz** (navy + kuning + Montserrat). Warna klien tidak menggantikan.
2. **Logo klien hanya di sampul** (di panel putih) dan di studi kasus yang relevan, dengan izin.
3. **Foto asli kegiatan** hanya untuk portofolio yang boleh dipublikasikan. Yang belum jelas izinnya ditulis ringkas tanpa foto atau logo.
4. **Semua angka portofolio sama** dengan compro dan bisa dibuktikan.
5. **Ikon garis putih seragam.** Gambar dan diagram di kartu teal atau navy.
6. **Tidak ada bukti yang dibuat-buat.** Dashboard atau tangkapan layar contoh diberi label "ilustrasi".
7. **Sumber ditulis** di bawah slide yang memakai data luar (10–11 pt, abu biru).

---

## 11. Catatan teknis pembuatan

- **Jangan memakai compro sebagai template langsung.** Compro adalah hasil konversi dari desain lain (bentuk *freeform*, font tertanam). Bangun deck baru dari token di bagian 3–5. Pakai `tanoto/_build/lib.py` sebagai titik mulai: sudah memuat token warna, `blank()` (latar gradasi), `card()`, `tab()`, `pill()`, `photo()`, `icon()`, `table()`, `header()`, `footer()` (teks footer di `lib.py` masih tertulis untuk Tanoto, ganti untuk klien lain).
- **Keluaran**: PPTX 16:9 → **PDF**. PPTX tetap disimpan sebagai versi editable. Deck yang sama dipakai untuk presentasi.
- **Aset yang bisa dipakai ulang**: logo, foto kegiatan (di `tanoto/_build/assets`), ikon garis (`_build/assets/icons`), diagram Framework.
- **Render untuk cek**: ubah ke PDF, lihat tiap slide. Cari teks yang meluber, tabel yang terpotong, foto buram.

---

## 12. Checklist sebelum kirim

- [ ] 16:9 (13,333 × 7,5 in). Margin 0,65 in.
- [ ] Montserrat terpasang; PDF tidak berganti font.
- [ ] Latar gradasi navy. Warna sesuai token. Kuning hanya sebagai aksen.
- [ ] Setiap slide jatuh ke pola A–H.
- [ ] Judul slide 28 pt, isi ≥ 13,5 pt (batas keras 12), tabel ≥ 10,8, tidak ada teks < 9 pt.
- [ ] Satu ide per slide, ≤ 12 baris, ≤ 5 bullet per kartu.
- [ ] Kicker kuning memuat kriteria penilaian RFP.
- [ ] Ikon garis seragam. Tidak ada clip art.
- [ ] Angka portofolio sama dengan compro. Sumber ditulis.
- [ ] Logo dan foto klien dipakai dengan izin.
- [ ] Nomor halaman kuning di kanan bawah, footer abu 9 pt.
- [ ] Ekspor PDF dan cek tiap slide. Simpan PPTX.
- [ ] Nama file: `Proposal Teknis - Sanggabiz - [Klien] [Program].pdf`.
