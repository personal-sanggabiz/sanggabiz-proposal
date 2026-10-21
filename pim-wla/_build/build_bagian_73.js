// Membuat berkas bagian 7.3 (pelatihan satu hari) untuk ditempel manual ke proposal PIM.
// Jalankan: NODE_PATH=$(npm root -g) node build_bagian_73.js
// Gaya mengikuti "Edited Proposal WLA dan Training Sanggabiz PIM 2026.docx": Arial, judul 173A5E, tabel berkepala navy.
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
} = require("docx");

const NAVY = "173A5E", FIRST = "EEF4F8", LINE = "D9D9D9";
const FONT = "Arial";
const W = 9590; // lebar isi (letter, margin 1325)
const OUT = path.join(__dirname, "..", "proposal", "v1 Bagian 7.3 Pelatihan ABK 1 Hari - PT Pupuk Iskandar Muda.docx");

const tr = (text, o = {}) => new TextRun({ text, font: FONT, color: "000000", size: 20, ...o });
const para = (text, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { after: o.after ?? 140, before: o.before ?? 0, line: 276 },
  keepNext: o.keepNext,
  children: Array.isArray(text) ? text : [tr(text, o.run)],
});
const h2 = (text) => new Paragraph({
  keepNext: true, spacing: { before: 200, after: 140, line: 269 },
  children: [tr(text, { bold: true, color: NAVY, size: 28 })],
});
const h3 = (text) => new Paragraph({
  keepNext: true, spacing: { before: 200, after: 120, line: 269 },
  children: [tr(text, { bold: true, color: NAVY, size: 24 })],
});
const focus = (text) => new Paragraph({
  keepNext: true, spacing: { after: 120, line: 276 },
  children: [tr("Fokus: ", { bold: true, color: NAVY }), tr(text)],
});

const border = { style: BorderStyle.SINGLE, size: 6, color: LINE };
const borders = { top: border, bottom: border, left: border, right: border };

// sel: string, atau array baris. Baris berawalan "~" ditebalkan.
function cell(content, width, o = {}) {
  const lines = Array.isArray(content) ? content : [content];
  return new TableCell({
    width: { size: width, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined,
    margins: { top: 90, bottom: 90, left: 100, right: 100 },
    children: lines.map((l, k) => {
      const base = { size: o.size || 16, color: o.color || "000000" };
      const m = l.startsWith("~") && l.match(/^~([^:]+:)(.*)$/);
      return new Paragraph({
        alignment: o.align || AlignmentType.LEFT,
        spacing: { after: m && k < lines.length - 1 ? 60 : 0, line: 252 },
        children: m ? [tr(m[1], { ...base, bold: true }), tr(m[2], base)] : [tr(l, { ...base, bold: o.bold })],
      });
    }),
  });
}
function table(widths, head, rows, o = {}) {
  const size = o.size || 16;
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths, alignment: AlignmentType.CENTER,
    rows: [
      new TableRow({ cantSplit: true, tableHeader: true,
        children: head.map((h, i) => cell(h, widths[i], { fill: NAVY, bold: true, color: "FFFFFF", align: AlignmentType.CENTER, size })) }),
      ...rows.map((r) => new TableRow({ cantSplit: true,
        children: r.map((c, i) => cell(c, widths[i], { size, fill: i === 0 && !o.noFirstFill ? FIRST : undefined })) })),
    ],
  });
}
const gap = () => new Paragraph({ spacing: { after: 120 }, children: [] });

const body = [];

// ---------------- 7.3 ----------------
body.push(h2("7.3 Pelatihan Transfer Kemampuan Analisis Beban Kerja"));
body.push(para("Pelatihan satu hari menggunakan satu kasus yang sama dari awal hingga akhir. Hasil setiap modul menjadi bahan modul berikutnya sehingga peserta mengalami alur lengkap dalam satu hari: merencanakan data, mewawancarai, mengisi lembar kerja, menguji mutu data, menghitung FTE, membaca heatmap, menyusun tindakan, dan membuat rencana penerapan. Kasus dibatasi pada dua posisi supaya setiap langkah dikerjakan sampai selesai. Latihan tambahan dibawa ke online implementation clinic pada D+30."));
body.push(table([2400, 7190], ["Elemen", "Desain Pelatihan"], [
  ["Durasi", "1 hari, 4 modul, total 8 JPL (1 JPL = 45 menit, atau 6 jam efektif). Sesi registrasi dan pembukaan berada di luar 8 JPL."],
  ["Pendekatan", "Studi kasus terpadu, penjelasan singkat, demonstrasi, latihan hitung, bermain peran, peer review, dan rencana penerapan."],
  ["Peserta", "Maksimal 30 peserta dalam 5 atau 6 kelompok. Peserta adalah tim yang akan terlibat dalam pengumpulan data, wawancara, penghitungan, validasi, dan tindak lanjut analisis beban kerja."],
  ["Bukti belajar", "Pre-test dan post-test, panduan wawancara, lembar kerja, FTE, heatmap, pohon akar masalah, RACI, serta rencana 30-60-90 hari."],
], { size: 18 }));

// ---------------- 7.3.2 ----------------
body.push(h3("7.3.2 Satu kasus empat modul"));
body.push(para("Hasil satu modul menjadi bahan modul berikutnya.", { keepNext: true }));
body.push(table([900, 3900, 1700, 3090], ["Modul", "Materi", "Waktu (WIB)", "Output"], [
  ["1", "Gambaran ABK, FTE, dan daftar pekerjaan", "09.00-10.30", "FTE awal, daftar pekerjaan, dan rencana data"],
  ["2", "Wawancara dan lembar kerja", "10.45-12.15", "Panduan wawancara dan lembar kerja dua posisi"],
  ["3", "Mutu data, hitung ulang FTE, dan heatmap", "13.15-14.45", "FTE akhir, heatmap, dan pohon akar masalah"],
  ["4", "Tindak lanjut, uji RACI, dan rencana 30-60-90 hari", "15.00-16.30", "Matriks tindakan dan rencana unit"],
]));

// ---------------- 7.3.3 ----------------
body.push(h3("7.3.3 Jadwal harian"));
body.push(table([1900, 5890, 1800], ["Waktu (WIB)", "Kegiatan", "JPL"], [
  ["08.30-09.00", "Registrasi dan pembukaan", "Di luar 8 JPL"],
  ["09.00-10.30", "Modul 1. Gambaran ABK, FTE, dan daftar pekerjaan", "2"],
  ["10.30-10.45", "Istirahat", ""],
  ["10.45-12.15", "Modul 2. Wawancara dan lembar kerja", "2"],
  ["12.15-13.15", "Makan siang dan ibadah", ""],
  ["13.15-14.45", "Modul 3. Memeriksa mutu data, menghitung ulang FTE, dan membaca heatmap", "2"],
  ["14.45-15.00", "Istirahat", ""],
  ["15.00-16.30", "Modul 4. Tindak lanjut, uji RACI, rencana 30-60-90 hari, post-test, dan penutupan", "2"],
  ["Total", "", "8"],
]));
body.push(para("Catatan: pre-test dikerjakan secara daring sebelum hari pelatihan. Jam dapat disesuaikan dengan jam kerja PIM.", { before: 120 }));

// ---------------- 7.3.4 ----------------
body.push(h3("7.3.4 Materi dan Silabus"));
body.push(focus("memahami ukuran beban kerja, mengumpulkan dan menguji data dari kasus nyata, lalu mengubah hasilnya menjadi tindakan dan rencana unit."));
body.push(table([2000, 2900, 2790, 1900], ["Topik", "Outline", "Tujuan", "Format"], [
  ["1. Gambaran ABK, FTE, dan daftar pekerjaan", [
    "Hasil akhir ABK: satu lembar hitung", "Enam langkah ABK", "FTE dan empat kategori beban",
    "Pekerjaan, hasil kerja, dan frekuensi", "Pekerjaan di luar uraian jabatan",
    "Empat sumber data, dua sumber untuk pekerjaan terbesar"],
   "Peserta mampu menjelaskan langkah ABK, menghitung FTE satu posisi, dan menyusun daftar pekerjaan beserta rencana pengumpulan datanya.",
   "Tebakan, penjelasan, latihan hitung, latihan formulir"],
  ["2. Wawancara dan lembar kerja", [
    "Susunan wawancara dan pertanyaan lanjutan", "Kekeliruan umum dalam wawancara", "Meminta contoh nyata dan angka",
    "Bermain peran tiga orang", "Dari catatan wawancara ke lembar kerja",
    "Catatan kerja harian dan data transaksi sebagai pembanding", "Mencatat selisih antarsumber"],
   "Peserta mampu menyusun panduan wawancara, mewawancarai dengan pertanyaan netral, dan mengisi lembar kerja ABK lalu mencocokkannya dengan sumber lain.",
   "Peragaan, bermain peran, latihan data kelompok"],
  ["3. Memeriksa mutu data, menghitung ulang FTE, dan membaca heatmap", [
    "Kelengkapan dan pekerjaan ganda", "Satuan dan periode tidak cocok", "Angka janggal dan hitungan ganda",
    "Uji kewajaran jam per orang", "Hitung ulang FTE dengan data teruji", "Tiga pilihan jam kerja efektif",
    "Membaca heatmap", "Pohon akar masalah"],
   "Peserta mampu memeriksa mutu data dengan daftar periksa, menghitung FTE dari data teruji, dan menelusuri penyebab beban tinggi.",
   "Latihan periksa silang, contoh, latihan hitung, diskusi"],
  ["4. Tindak lanjut, uji RACI, dan rencana 30-60-90 hari", [
    "Pilihan tindak lanjut sesuai akar masalah", "Urutan prioritas", "Uji RACI sebelum dan sesudah",
    "Menguji hasil di depan rekan yang berperan sebagai manajer", "Rencana 30-60-90 hari", "Online implementation clinic D+30"],
   "Peserta mampu memilih tindak lanjut, menguji dampaknya pada RACI, dan menyusun rencana unit yang terukur.",
   "Latihan kelompok, tinjauan silang, latihan, post-test"],
]));

// ---------------- 7.3.5 ----------------
body.push(h3("7.3.5 Aktivitas dan Output"));
body.push(para("Angka menit di dalam kurung adalah waktu tiap langkah. Jumlah menit satu modul sama dengan 90 menit.", { keepNext: true }));
const act = (steps, alat, indikator, output) => [
  steps,
  [`~Alat atau media: ${alat}`, `~Indikator capaian: ${indikator}`, `~Output konkret: ${output}`],
];
const rowsAct = [
  ["1. Gambaran ABK, FTE, dan daftar pekerjaan", ...act([
    "1. Peserta menebak FTE satu posisi tanpa alat hitung dan menulisnya di kartu (8 menit).",
    "2. Fasilitator menunjukkan sebaran tebakan, lalu menguraikan lembar hitung yang sudah jadi menjadi enam langkah (17 menit).",
    "3. Peserta menghitung FTE satu posisi dari kasus kelompok (25 menit).",
    "4. Fasilitator menjelaskan kolom formulir dan empat sumber data. Tiap peserta mengisi minimal 8 pekerjaan untuk satu posisi (25 menit).",
    "5. Kelompok memilih sumber data untuk pekerjaan terbesar, lalu saling menilai rencananya (15 menit)."],
    "Flipchart, kartu tebakan, lembar hitung FTE, paket kasus, formulir daftar pekerjaan, formulir rencana data",
    "Minimal 80% peserta menghitung FTE dengan selisih paling banyak 0,10 dari kunci. Minimal 80% peserta mengisi 8 pekerjaan. Pekerjaan terbesar punya minimal dua sumber data.",
    "FTE awal, daftar pekerjaan, dan rencana pengumpulan data")],
  ["2. Wawancara dan lembar kerja", ...act([
    "1. Fasilitator memperagakan wawancara yang menggiring dan yang netral. Peserta mencatat perbedaannya (12 menit).",
    "2. Pasangan menyusun panduan wawancara untuk satu posisi (13 menit).",
    "3. Kelompok tiga orang berlatih tiga putaran dan bergantian sebagai pewawancara, narasumber, dan pengamat (30 menit).",
    "4. Fasilitator membagikan paket data: catatan wawancara, catatan harian, dan cuplikan register. Kelompok mengisi lembar kerja dua posisi dan mencatat selisih antarsumber (30 menit).",
    "5. Fasilitator memeriksa daftar selisih tiap kelompok dan memberi umpan balik (5 menit)."],
    "Lembar pengamatan, kartu peran, paket data, lembar kerja ABK",
    "Panduan memuat minimal 10 pertanyaan terbuka. Minimal 75% peserta menggali 2 dari 3 fakta tersembunyi. Tiap kelompok mencatat minimal 3 selisih beserta alasan memilih sumber.",
    "Panduan wawancara, catatan wawancara, lembar kerja dua posisi, dan daftar selisih data")],
  ["3. Memeriksa mutu data, menghitung ulang FTE, dan membaca heatmap", ...act([
    "1. Fasilitator menjelaskan daftar periksa enam butir (10 menit).",
    "2. Kelompok mencari enam kesalahan sisipan di lembar kelompok lain. Pemilik lembar memutuskan menerima atau menolak tiap temuan (25 menit).",
    "3. Fasilitator memberi contoh: beban 1.600 jam dengan tiga pilihan jam kerja efektif menghasilkan FTE 1,28, 1,14, dan 1,07 (10 menit).",
    "4. Kelompok menghitung ulang FTE dua posisi pada tiga pilihan asumsi dan menyusun heatmap (25 menit).",
    "5. Kelompok menelusuri akar masalah satu posisi yang padat atau berlebih dengan pohon akar masalah (20 menit)."],
    "Lembar latihan, daftar periksa, lembar hitung FTE, heatmap",
    "Tiap kelompok menemukan minimal 5 dari 6 kesalahan. Tiap posisi punya FTE pada tiga pilihan asumsi. Tiap pohon akar masalah memuat minimal dua penyebab.",
    "Daftar masalah data yang diputuskan, FTE akhir, heatmap, dan pohon akar masalah")],
  ["4. Tindak lanjut, uji RACI, dan rencana 30-60-90 hari", ...act([
    "1. Fasilitator menjelaskan pencocokan akar masalah dengan pilihan tindakan dan urutan prioritasnya (10 menit).",
    "2. Kelompok memilih dua temuan, membuat potongan RACI sebelum dan sesudah, lalu menghitung ulang FTE penerima peran (25 menit).",
    "3. Kelompok bertukar hasil. Peserta kelompok lain berperan sebagai manajer dan menguji dengan kartu pertanyaan (15 menit).",
    "4. Fasilitator menjelaskan cara menulis target yang terukur. Tiap peserta menyusun rencana 30-60-90 hari untuk unitnya (15 menit).",
    "5. Peserta mengerjakan post-test (15 menit).",
    "6. Peserta mengisi evaluasi akhir dan formulir Level 4, lalu menerima sertifikat (10 menit)."],
    "Formulir RACI, matriks prioritas, kartu pertanyaan, formulir rencana, soal post-test, formulir Level 4",
    "Tiap tindakan punya penanggung jawab, perkiraan perubahan FTE, dan hasil uji RACI. Skor post-test minimal 75 dan naik rata-rata minimal 20% dari nilai awal.",
    "Matriks tindak lanjut, RACI sebelum dan sesudah, rencana 30-60-90 hari, hasil post-test, dan komitmen penerapan")],
];
body.push(new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: [2000, 4600, 2990], alignment: AlignmentType.CENTER,
  rows: [
    new TableRow({ cantSplit: true, tableHeader: true,
      children: ["Topik", "Aktivitas Pembelajaran", "Alat, Indikator, dan Output"].map((h, i) => cell(h, [2000, 4600, 2990][i], { fill: NAVY, bold: true, color: "FFFFFF", align: AlignmentType.CENTER })) }),
    ...rowsAct.map((r) => new TableRow({ cantSplit: true,
      children: r.map((c, i) => cell(c, [2000, 4600, 2990][i], { fill: i === 0 ? FIRST : undefined })) })),
  ],
}));
body.push(para("Catatan: angka pada contoh Modul 3 adalah contoh latihan dan bukan angka PIM.", { before: 120 }));

// ---------------- 7.3.6 ----------------
body.push(h3("7.3.6 Output dan evaluasi pembelajaran"));
body.push(table([1700, 7890], ["Tahap", "Output Peserta"], [
  ["Modul 1", "FTE awal, daftar pekerjaan, dan rencana pengumpulan data"],
  ["Modul 2", "Panduan wawancara, catatan wawancara, lembar kerja dua posisi, dan daftar selisih data"],
  ["Modul 3", "Daftar masalah data yang diputuskan, FTE akhir, heatmap, dan pohon akar masalah"],
  ["Modul 4", "Matriks tindak lanjut, RACI sebelum dan sesudah, rencana 30-60-90 hari, hasil post-test, dan komitmen peserta"],
]));
body.push(gap());
body.push(table([1500, 2800, 3490, 1800], ["Level", "Instrumen", "Target", "Waktu"], [
  ["1. Reaksi", "Evaluasi akhir", "Rata-rata minimal 4,2 dari 5 dan 85% peserta menilai baik", "Akhir pelatihan"],
  ["2. Pembelajaran", "Pre-test, post-test, latihan FTE, dan praktik wawancara", "Post-test minimal 75 dan kenaikan rata-rata minimal 20%", "Sebelum dan akhir pelatihan"],
  ["3. Perilaku", "Daftar cek atasan, bukti penggunaan, dan online implementation clinic", "Minimal 75% peserta sasaran memakai hasil pelatihan pada satu aktivitas ABK", "D+30"],
  ["4. Hasil", "Formulir satu halaman yang dipantau PIM", "Kelengkapan data minimal 90%, isu berulang turun, dan tindak lanjut memiliki PIC serta target waktu", "Dipantau PIM"],
]));
body.push(para("PIM menerima lembar hitung FTE baku, seluruh formulir, lembar pengamatan, kunci kasus, materi pelatihan, pre-test dan post-test, e-sertifikat, laporan pelatihan, serta formulir evaluasi Level 4. Kasus, jadwal, dan komposisi peserta dapat dikunci bersama PIM sebelum pelaksanaan.", { before: 140 }));

const doc = new Document({
  creator: "Sanggabiz", title: "Bagian 7.3 Pelatihan ABK 1 Hari - PT Pupuk Iskandar Muda",
  styles: { default: { document: { run: { font: FONT, size: 20 } } } },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1181, bottom: 1123, left: 1325, right: 1325 } } },
    children: body,
  }],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("OK", OUT); });
