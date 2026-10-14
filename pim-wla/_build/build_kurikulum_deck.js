// Membuat deck proposal kurikulum pelatihan ABK dua hari (gaya compro Sanggabiz: navy, kuning, teal).
// Jalankan: NODE_PATH=<folder node_modules berisi pptxgenjs> node build_kurikulum_deck.js
// Font Arial dipakai di seluruh deck agar tampil sama di semua komputer.
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");

const A = (f) => path.join(__dirname, "assets", f);
const I = (n, c) => A(`icons/${n}_${c}.png`);
const OUT = path.join(__dirname, "..", "proposal", "v1 Proposal Kurikulum Pelatihan ABK 2 Hari - PT Pupuk Iskandar Muda.pptx");

const THEME = {
  name: "Sanggabiz",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "1D1D1F", lt1: "FFFFFF", dk2: "0B396D", lt2: "EEF3F9",
    accent1: "0B396D", accent2: "FFC00C", accent3: "2F7E9D", accent4: "092647", accent5: "B0C0CF", accent6: "15908F",
    hlink: "2446BC", folHlink: "727171",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 inci
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Proposal Kurikulum Pelatihan Transfer Kemampuan Analisis Beban Kerja";
pres.author = "Sanggabiz";
pres.company = "Sanggabiz";
const C = pres.SchemeColor;
const R = pres.ShapeType;

// ---------- layout ----------
pres.defineSlideMaster({
  title: "DARK",
  background: { path: A("bg_navy.png") },
  objects: [
    { image: { x: 0.65, y: 0.45, w: 1.9, h: 0.377, path: A("logo_white.png"), objectName: "Logo Sanggabiz" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.65, y: 1.0, w: 12, h: 0.9, fontSize: 32, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "LIGHT",
  background: { color: C.background1 },
  objects: [
    { image: { x: 10.5, y: 0.38, w: 1.75, h: 0.347, path: A("logo_navy.png"), objectName: "Logo Sanggabiz" } },
    { image: { x: 12.3, y: 0.2, w: 0.62, h: 0.614, path: A("logo_pim.png"), objectName: "Logo PIM" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.8, h: 0.85, fontSize: 26, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: "Sanggabiz  |  Proposal Kurikulum Pelatihan ABK  |  PT Pupuk Iskandar Muda", options: { x: 0.5, y: 7.08, w: 9, h: 0.28, fontSize: 9, color: C.text1, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.08, w: 0.7, h: 0.28, fontSize: 9, color: C.text1, align: "right" },
});
pres.defineSlideMaster({ title: "COVER", background: { path: A("bg_navy.png") }, objects: [] });

// ---------- penolong ----------
function pill(s, text, x, y, w, h, o = {}) {
  s.addText(text, {
    shape: R.roundRect, rectRadius: h / 2, x, y, w, h, isTextBox: true,
    fill: o.fill ? { color: o.fill } : undefined, line: o.line ? { color: o.line, width: 1.5 } : undefined,
    color: o.color || C.text2, bold: o.bold !== false, fontSize: o.size || 12, align: "center", valign: "middle", margin: 0,
    charSpacing: o.charSpacing, objectName: o.name || "Label",
  });
}
function card(s, x, y, w, h, name) {
  s.addShape(R.roundRect, { x, y, w, h, rectRadius: 0.2, fill: { color: C.accent3 }, line: { color: C.background1, width: 2 }, objectName: name || "Kartu" });
}
const bullets = (items, o = {}) => items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 2, ...o } }));
const steps = (items) => items.map((t, i) => ({ text: `${i + 1}. ${t}`, options: { breakLine: i < items.length - 1, paraSpaceAfter: 3 } }));

// ---------- isi ----------
const HARI1 = [
  { no: 1, topik: "Gambaran ABK dan menghitung FTE", format: "Tebakan, penjelasan, latihan hitung",
    outline: ["Hasil akhir ABK: satu lembar hitung", "Enam langkah ABK", "Beban kerja, kinerja, jumlah pegawai", "FTE dan empat kategori beban"],
    tujuan: "Peserta mampu menjelaskan langkah ABK dan menghitung FTE satu posisi.",
    akt: ["Peserta menebak FTE satu posisi tanpa alat hitung.", "Fasilitator menunjukkan sebaran tebakan, lalu menguraikan lembar hitung yang sudah jadi menjadi enam langkah.", "Peserta menghitung FTE satu posisi dari kasus kelompok dan mendaftar data yang perlu dibuktikan."],
    alat: "Flipchart, lembar hitung FTE, paket kasus",
    ukur: "Minimal 80% peserta menghitung FTE dengan selisih paling banyak 0,10 dari kunci.",
    out: "FTE awal dan daftar data yang perlu dibuktikan" },
  { no: 2, topik: "Daftar pekerjaan dan rencana pengumpulan data", format: "Penjelasan, latihan formulir",
    outline: ["Pekerjaan, hasil kerja, dan frekuensi", "Pekerjaan di luar uraian jabatan", "Empat sumber data", "Dua sumber untuk pekerjaan terbesar"],
    tujuan: "Peserta mampu menyusun daftar pekerjaan satu posisi dan rencana pengumpulan data yang dapat diperiksa.",
    akt: ["Fasilitator menjelaskan kolom formulir dan empat sumber data.", "Tiap peserta mengisi minimal 8 pekerjaan untuk satu posisi.", "Kelompok memilih sumber data, lalu saling menilai rencananya."],
    alat: "Formulir daftar pekerjaan, formulir rencana data",
    ukur: "Minimal 80% peserta mengisi 8 pekerjaan. Pekerjaan terbesar punya minimal dua sumber data.",
    out: "Daftar pekerjaan dan rencana pengumpulan data" },
  { no: 3, topik: "Teknik dan praktik wawancara", format: "Peragaan, bermain peran",
    outline: ["Susunan wawancara dan pertanyaan lanjutan", "Kekeliruan umum dalam wawancara", "Meminta contoh nyata dan angka", "Bermain peran tiga orang"],
    tujuan: "Peserta mampu menyusun panduan wawancara dan mewawancarai dengan pertanyaan netral.",
    akt: ["Fasilitator memperagakan wawancara yang menggiring dan yang netral.", "Pasangan menyusun panduan wawancara untuk satu posisi.", "Kelompok tiga orang berlatih tiga putaran dan bergantian peran."],
    alat: "Lembar pengamatan, kartu peran",
    ukur: "Panduan memuat minimal 10 pertanyaan terbuka. Minimal 75% peserta menggali 2 dari 3 fakta tersembunyi.",
    out: "Panduan wawancara dan catatan wawancara" },
  { no: 4, topik: "Mengisi lembar kerja dan mencocokkan data", format: "Latihan data kelompok",
    outline: ["Dari catatan wawancara ke lembar kerja", "Catatan kerja harian", "Data transaksi sebagai pembanding", "Mencatat selisih antarsumber"],
    tujuan: "Peserta mampu mengisi lembar kerja ABK dan mencocokkannya dengan sumber lain.",
    akt: ["Fasilitator membagikan paket data: catatan wawancara, catatan harian, dan cuplikan register.", "Kelompok mengisi lembar kerja tiga posisi dan mencatat selisih antarsumber.", "Fasilitator memeriksa daftar masalah data tiap kelompok."],
    alat: "Paket data, lembar kerja ABK",
    ukur: "Tiap kelompok mencatat minimal 3 selisih beserta alasan memilih sumber.",
    out: "Lembar kerja tiga posisi dan daftar masalah data" },
];
const HARI2 = [
  { no: 5, topik: "Memeriksa mutu data", format: "Latihan periksa silang",
    outline: ["Kelengkapan dan pekerjaan ganda", "Satuan dan periode tidak cocok", "Angka janggal dan hitungan ganda", "Uji kewajaran jam per orang"],
    tujuan: "Peserta mampu memeriksa mutu data dengan daftar periksa dan memutuskan perbaikannya.",
    akt: ["Fasilitator menjelaskan daftar periksa enam butir.", "Kelompok mencari enam kesalahan sisipan, lalu memeriksa lembar kelompok lain.", "Pemilik lembar memutuskan menerima atau menolak tiap temuan."],
    alat: "Lembar latihan, daftar periksa",
    ukur: "Tiap kelompok menemukan minimal 5 dari 6 kesalahan.",
    out: "Daftar masalah data yang diputuskan dan lembar kerja yang diperbaiki" },
  { no: 6, topik: "Hitung ulang FTE, heatmap, dan akar masalah", format: "Contoh, latihan hitung, diskusi",
    outline: ["Hitung ulang FTE dengan data teruji", "Tiga pilihan jam kerja efektif", "Membaca heatmap", "Pohon akar masalah"],
    tujuan: "Peserta mampu menghitung FTE dari data teruji dan menelusuri penyebab beban tinggi.",
    akt: ["Fasilitator memberi contoh: beban 1.600 jam dengan tiga pilihan jam kerja efektif menghasilkan FTE 1,28, 1,14, dan 1,07.", "Kelompok menghitung ulang FTE tiga posisi dan menyusun heatmap.", "Kelompok menelusuri akar masalah satu posisi yang padat atau berlebih."],
    alat: "Lembar hitung FTE, heatmap",
    ukur: "Tiap posisi punya FTE pada tiga pilihan asumsi. Tiap pohon akar masalah memuat minimal dua penyebab.",
    out: "FTE akhir, heatmap, dan pohon akar masalah" },
  { no: 7, topik: "Tindak lanjut, uji RACI, dan panel manajer", format: "Latihan kelompok, simulasi panel",
    outline: ["Pilihan tindak lanjut", "Uji RACI sebelum dan sesudah", "Urutan prioritas", "Mempertahankan hasil di depan manajer"],
    tujuan: "Peserta mampu memilih tindak lanjut, menguji dampaknya pada RACI, dan mempertahankan hasilnya.",
    akt: ["Fasilitator menjelaskan pencocokan akar masalah dengan tindakan.", "Kelompok memilih tiga temuan, membuat potongan RACI sebelum dan sesudah, dan menghitung ulang FTE penerima peran.", "Peserta kelompok lain berperan sebagai manajer dan menguji hasilnya."],
    alat: "Formulir RACI, matriks prioritas, kartu pertanyaan panel",
    ukur: "Tiap tindakan punya penanggung jawab, perkiraan perubahan FTE, dan hasil uji RACI.",
    out: "Matriks tindak lanjut, RACI sebelum dan sesudah, dan catatan keputusan" },
  { no: 8, topik: "Rencana 30-60-90 hari dan penutupan", format: "Latihan, post-test",
    outline: ["Rencana 30-60-90 hari", "Online implementation clinic D+30", "Komitmen penerapan"],
    tujuan: "Peserta mampu menyusun rencana unit dan menunjukkan penguasaan lewat post-test.",
    akt: ["Fasilitator menjelaskan cara menulis target yang terukur.", "Tiap peserta menyusun rencana 30-60-90 hari untuk unitnya.", "Peserta mengerjakan post-test, mengisi evaluasi akhir, dan menerima sertifikat serta formulir Level 4."],
    alat: "Formulir rencana, soal post-test, formulir Level 4",
    ukur: "Skor post-test minimal 75 dan naik rata-rata minimal 20% dari nilai awal.",
    out: "Rencana 30-60-90 hari, hasil post-test, dan komitmen" },
];

// ---------- tabel ----------
const BORDER = { type: "solid", pt: 0.75, color: "FFFFFF" };
const head = (t, extra = {}) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.accent1 }, align: "center", valign: "middle", fontSize: 12, border: BORDER, ...extra } });

function tabelSilabus(s, judul, fokus, mods) {
  const rows = [[head("Hari dan fokus"), head("Topik"), head("Outline materi"), head("Tujuan pembelajaran"), head("Format materi")]];
  mods.forEach((m, i) => {
    const fill = { color: i % 2 ? C.background1 : C.background2 };
    const first = i === 0 ? [{
      text: [
        { text: judul, options: { bold: true, color: C.accent6, fontSize: 14, breakLine: true } },
        { text: " ", options: { fontSize: 6, breakLine: true } },
        { text: "Fokus:", options: { bold: true, color: C.text1, fontSize: 11, breakLine: true } },
        { text: fokus, options: { italic: true, color: C.text1, fontSize: 11 } },
      ],
      options: { rowspan: mods.length, valign: "middle", fill: { color: C.background2 }, border: BORDER, margin: [0.1, 0.12, 0.1, 0.12] },
    }] : [];
    rows.push([
      ...first,
      { text: `${m.no}. ${m.topik}`, options: { bold: true, fontSize: 12, color: C.text1, valign: "middle", fill, border: BORDER } },
      { text: bullets(m.outline, { fontSize: 11, color: C.text1 }), options: { valign: "middle", fill, border: BORDER } },
      { text: m.tujuan, options: { fontSize: 11, color: C.text1, valign: "middle", fill, border: BORDER } },
      { text: m.format, options: { bold: true, fontSize: 11, color: C.accent6, align: "center", valign: "middle", fill, border: BORDER } },
    ]);
  });
  s.addTable(rows, { x: 0.5, y: 1.25, w: 12.3, colW: [2.2, 2.3, 3.6, 2.9, 1.3], rowH: [0.42, 1.25, 1.25, 1.25, 1.25].slice(0, mods.length + 1), margin: [0.06, 0.1, 0.06, 0.1], fontFace: "Arial", objectName: "Tabel silabus" });
}

function tabelAktivitas(s, mods) {
  const rows = [[head("Topik"), head("Aktivitas pembelajaran"), head("Alat atau media"), head("Indikator capaian"), head("Output konkret")]];
  mods.forEach((m, i) => {
    const fill = { color: i % 2 ? C.background1 : C.background2 };
    rows.push([
      { text: `${m.no}. ${m.topik}`, options: { bold: true, fontSize: 11, color: C.text1, valign: "middle", fill, border: BORDER } },
      { text: steps(m.akt), options: { fontSize: 10.5, color: C.text1, valign: "middle", fill, border: BORDER } },
      { text: m.alat, options: { fontSize: 10.5, color: C.text1, valign: "middle", fill, border: BORDER } },
      { text: m.ukur, options: { fontSize: 10.5, color: C.text1, valign: "middle", fill, border: BORDER } },
      { text: m.out, options: { bold: true, fontSize: 10.5, color: C.accent6, valign: "middle", fill, border: BORDER } },
    ]);
  });
  s.addTable(rows, { x: 0.5, y: 1.25, w: 12.3, colW: [1.9, 4.9, 1.6, 2.3, 1.6], rowH: [0.42, 1.25, 1.25, 1.25, 1.25].slice(0, mods.length + 1), margin: [0.06, 0.1, 0.06, 0.1], fontFace: "Arial", objectName: "Tabel aktivitas" });
}

async function build() {
  // ===== bagian 1: gambaran =====
  pres.addSection({ title: "Gambaran" });

  // 1. sampul
  let s = pres.addSlide({ masterName: "COVER", sectionTitle: "Gambaran" });
  s.addShape(R.ellipse, { x: 8.6, y: 3.1, w: 6.2, h: 6.2, fill: { color: C.accent3, transparency: 80 }, line: { type: "none" }, objectName: "Dekorasi lingkaran besar" });
  s.addShape(R.ellipse, { x: 11.55, y: 5.55, w: 1.25, h: 1.25, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Dekorasi lingkaran kuning" });
  s.addImage({ path: A("logo_white.png"), x: 0.65, y: 0.55, w: 2.4, h: 0.475, objectName: "Logo Sanggabiz" });
  pill(s, "PROPOSAL KURIKULUM", 0.65, 1.95, 2.9, 0.42, { fill: C.accent2, size: 12, charSpacing: 2, name: "Label proposal" });
  s.addText([{ text: "Pelatihan Transfer Kemampuan", options: { breakLine: true } }, { text: "Analisis Beban Kerja" }], { x: 0.65, y: 2.5, w: 8.6, h: 1.6, fontSize: 38, bold: true, color: C.background1, valign: "top", margin: 0, isTextBox: true, objectName: "Judul sampul" });
  s.addText("Dua hari di Medan untuk PT Pupuk Iskandar Muda", { x: 0.65, y: 4.3, w: 8, h: 0.45, fontSize: 18, color: C.accent5, margin: 0, isTextBox: true, objectName: "Subjudul sampul" });
  ["2 hari, 16 JPL", "8 modul", "30 peserta"].forEach((t, i) => pill(s, t, 0.65 + i * 2.4, 5.55, 2.2, 0.5, { line: C.background1, color: C.background1, size: 14, name: `Ringkasan ${i + 1}` }));
  s.addText("Diajukan oleh Sanggabiz   |   Oktober 2026", { x: 0.65, y: 6.7, w: 8, h: 0.35, fontSize: 14, color: C.background1, margin: 0, isTextBox: true, objectName: "Penyusun dan bulan" });
  s.addShape(R.roundRect, { x: 9.1, y: 0.5, w: 3.6, h: 1.15, rectRadius: 0.575, fill: { color: C.background1 }, line: { type: "none" }, objectName: "Kapsul klien" });
  s.addImage({ path: A("logo_pim.png"), x: 9.3, y: 0.62, w: 0.9, h: 0.89, objectName: "Logo PIM" });
  s.addText([{ text: "Disiapkan untuk", options: { fontSize: 11, color: C.text1, breakLine: true } }, { text: "PT Pupuk Iskandar Muda", options: { fontSize: 13, bold: true, color: C.text2 } }], { x: 10.3, y: 0.62, w: 2.3, h: 0.9, valign: "middle", margin: 0, isTextBox: true, objectName: "Penerima" });

  // 2. ringkasan
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Gambaran" });
  s.addText("Dua hari untuk menguasai siklus ABK", { placeholder: "title" });
  s.addText("Peserta mengerjakan satu kasus dari awal sampai akhir: menyusun daftar pekerjaan, mengumpulkan dan memeriksa data, menghitung FTE (beban kerja dalam satuan pegawai penuh waktu), membaca hasilnya, lalu menyusun rencana 30-60-90 hari untuk unitnya.", { x: 0.65, y: 2.2, w: 4.7, h: 3.6, fontSize: 16, color: C.background1, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true, objectName: "Uraian ringkasan" });
  const stat = [["2", "hari tatap muka di Medan", "FiMapPin"], ["16", "JPL, setara 12 jam efektif", "FiClock"], ["8", "modul, masing-masing 2 JPL", "FiLayers"], ["30", "peserta, maksimal", "FiUsers"], ["5", "kelompok dengan satu kasus bersama", "FiTarget"], ["2", "fasilitator: utama dan pendamping", "FiAward"]];
  stat.forEach(([n, l, ic], i) => {
    const x = 5.9 + (i % 3) * 2.45, y = 2.0 + Math.floor(i / 3) * 2.35;
    s.addImage({ path: I(ic, "w"), x, y, w: 0.45, h: 0.45, objectName: `Ikon ${l}` });
    s.addText(n, { x, y: y + 0.5, w: 2.2, h: 0.85, fontSize: 44, bold: true, color: C.background1, margin: 0, valign: "middle", isTextBox: true, objectName: `Angka ${l}` });
    s.addText(l, { x, y: y + 1.35, w: 2.2, h: 0.7, fontSize: 14, bold: true, color: C.accent2, margin: 0, valign: "top", isTextBox: true, objectName: `Label ${l}` });
  });

  // 3. dasar rancangan
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Gambaran" });
  s.addText("Enam pertimbangan di balik kurikulum", { placeholder: "title" });
  const dasar = [
    ["Urutan kerja ABK", "Peserta merencanakan dan mengumpulkan data dulu, baru menghitung, memeriksa, dan menindaklanjuti."],
    ["Mulai dari hasil akhir", "Peserta melihat lembar hitung yang sudah jadi, lalu menguraikannya mundur menjadi langkah kerja."],
    ["Mengalami dulu", "Peserta menebak FTE sebelum menerima rumus, lalu belajar dari sebaran tebakan sendiri."],
    ["Satu kasus utuh", "Hasil satu modul menjadi bahan modul berikutnya, dari daftar pekerjaan sampai rencana tindak."],
    ["FTE dihitung dua kali", "Pertama dari perkiraan, kedua dari data yang sudah diperiksa, agar terlihat pengaruh mutu data."],
    ["Data dari banyak sumber", "Hasil wawancara dicocokkan dengan catatan harian dan data transaksi."],
  ];
  dasar.forEach(([t, d], i) => {
    const x = 0.65 + (i % 3) * 4.11, y = 2.3 + Math.floor(i / 3) * 2.15;
    card(s, x, y, 3.81, 1.85, `Kartu ${t}`);
    pill(s, t.toUpperCase(), x + 0.25, y - 0.2, Math.min(3.3, 0.5 + t.length * 0.17), 0.4, { fill: C.accent2, size: 11, name: `Tab ${t}` });
    s.addText(d, { x: x + 0.25, y: y + 0.28, w: 3.31, h: 1.45, fontSize: 14, color: C.background1, valign: "middle", margin: 0, isTextBox: true, objectName: `Isi ${t}` });
  });
  s.addText("Rujukan: Permen PANRB Nomor 1 Tahun 2020, Wiggins dan McTighe (2005), Kolb (1984), Anderson dan Krathwohl (2001), Harms dkk. (2019).", { x: 0.65, y: 6.75, w: 12, h: 0.3, fontSize: 10, color: C.accent5, margin: 0, isTextBox: true, objectName: "Rujukan" });

  // 4. alur kasus
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Gambaran" });
  s.addText("Satu kasus, delapan modul", { placeholder: "title" });
  s.addText("Hasil satu modul menjadi bahan modul berikutnya.", { x: 0.65, y: 1.95, w: 9, h: 0.35, fontSize: 16, color: C.accent5, margin: 0, isTextBox: true, objectName: "Pengantar alur" });
  const alur = [
    [1, "Gambaran ABK dan hitung FTE", "FTE awal", "FiTarget"], [2, "Daftar pekerjaan dan rencana data", "Rencana data", "FiList"],
    [3, "Teknik dan praktik wawancara", "Panduan wawancara", "FiMessageSquare"], [4, "Isi lembar kerja dan cocokkan data", "Lembar kerja", "FiFileText"],
    [5, "Periksa mutu data", "Data teruji", "FiSearch"], [6, "Hitung ulang FTE dan baca heatmap", "FTE akhir dan heatmap", "FiTrendingUp"],
    [7, "Tindak lanjut dan uji RACI", "Matriks tindakan", "FiRepeat"], [8, "Rencana 30-60-90 hari", "Rencana unit", "FiCheckCircle"],
  ];
  [["HARI 1", 2.85], ["HARI 2", 5.25]].forEach(([lab, y]) => pill(s, lab, 0.65, y - 0.4, 1.3, 0.32, { fill: C.accent2, size: 11, name: `Label ${lab}` }));
  alur.forEach(([n, t, o, ic], i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = 0.65 + col * 3.1, y = 2.85 + row * 2.4;
    card(s, x, y, 2.7, 1.75, `Modul ${n}`);
    s.addShape(R.ellipse, { x: x + 0.15, y: y + 0.15, w: 0.5, h: 0.5, fill: { color: C.accent2 }, line: { type: "none" }, objectName: `Nomor ${n}` });
    s.addText(String(n), { x: x + 0.15, y: y + 0.15, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Angka ${n}` });
    s.addImage({ path: I(ic, "w"), x: x + 2.05, y: y + 0.17, w: 0.45, h: 0.45, objectName: `Ikon modul ${n}` });
    s.addText(t, { x: x + 0.2, y: y + 0.68, w: 2.3, h: 0.55, fontSize: 14, bold: true, color: C.background1, valign: "top", margin: 0, isTextBox: true, objectName: `Judul modul ${n}` });
    s.addShape(R.roundRect, { x: x + 0.2, y: y + 1.28, w: 2.3, h: 0.33, rectRadius: 0.165, fill: { color: C.accent4 }, line: { type: "none" }, objectName: `Kapsul hasil ${n}` });
    s.addText(o, { x: x + 0.2, y: y + 1.28, w: 2.3, h: 0.33, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `Hasil modul ${n}` });
    if (col < 3) s.addShape(R.rightArrow, { x: x + 2.76, y: y + 0.75, w: 0.28, h: 0.25, fill: { color: C.background1 }, line: { type: "none" }, objectName: `Panah ${n}` });
  });

  // 5. jadwal
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Gambaran" });
  s.addText("Jadwal harian", { placeholder: "title" });
  const jt = (t, o = {}) => ({ text: t, options: { fontSize: 13, color: C.text1, valign: "middle", border: BORDER, ...o } });
  const wkt = [["08.30 - 09.00", "Registrasi dan pembukaan", "Ulasan Hari 1 oleh peserta"], ["09.00 - 10.30", "Modul 1. Gambaran ABK dan menghitung FTE", "Modul 5. Memeriksa mutu data"], ["10.45 - 12.15", "Modul 2. Daftar pekerjaan dan rencana pengumpulan data", "Modul 6. Hitung ulang FTE, heatmap, dan akar masalah"], ["13.15 - 14.45", "Modul 3. Teknik dan praktik wawancara", "Modul 7. Tindak lanjut, uji RACI, dan panel manajer"], ["15.00 - 16.30", "Modul 4. Mengisi lembar kerja dan mencocokkan data", "Modul 8. Rencana 30-60-90 hari dan penutupan"]];
  const rowsJ = [[head("Waktu (WIB)", { fontSize: 13 }), head("Hari 1", { fontSize: 13 }), head("Hari 2", { fontSize: 13 })]];
  wkt.forEach((r, i) => { const fill = { color: i % 2 ? C.background1 : C.background2 }; rowsJ.push([jt(r[0], { bold: true, align: "center", fill }), jt(r[1], { fill }), jt(r[2], { fill })]); });
  s.addTable(rowsJ, { x: 0.5, y: 1.4, w: 12.3, colW: [2.1, 5.1, 5.1], rowH: [0.45, 0.62, 0.62, 0.62, 0.62, 0.62], margin: [0.05, 0.12, 0.05, 0.12], fontFace: "Arial", objectName: "Tabel jadwal" });
  s.addText("Istirahat 15 menit pada pagi dan sore hari, makan siang pukul 12.15 sampai 13.15 WIB. Sesi pukul 08.30 sampai 09.00 berada di luar 16 JPL. Pre-test dikerjakan sebelum Hari 1. Jam dapat disesuaikan dengan jam kerja PIM.", { x: 0.5, y: 5.5, w: 12.3, h: 0.8, fontSize: 13, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: "Catatan jadwal" });

  // ===== bagian 2: silabus =====
  pres.addSection({ title: "Silabus" });
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Silabus" });
  s.addText([{ text: "Materi dan Silabus  ", options: { color: C.text2 } }, { text: "Hari 1", options: { color: C.accent6 } }], { placeholder: "title" });
  tabelSilabus(s, "Hari 1: Merencanakan dan Mengumpulkan Data", "Memahami ukuran beban kerja, merencanakan, dan mengumpulkan data dari kasus nyata.", HARI1);
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Silabus" });
  s.addText([{ text: "Aktivitas dan Output  ", options: { color: C.text2 } }, { text: "Hari 1", options: { color: C.accent6 } }], { placeholder: "title" });
  tabelAktivitas(s, HARI1);
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Silabus" });
  s.addText([{ text: "Materi dan Silabus  ", options: { color: C.text2 } }, { text: "Hari 2", options: { color: C.accent6 } }], { placeholder: "title" });
  tabelSilabus(s, "Hari 2: Memeriksa dan Menindaklanjuti", "Memeriksa data, membaca hasil, dan menindaklanjuti hingga rencana unit.", HARI2);
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Silabus" });
  s.addText([{ text: "Aktivitas dan Output  ", options: { color: C.text2 } }, { text: "Hari 2", options: { color: C.accent6 } }], { placeholder: "title" });
  tabelAktivitas(s, HARI2);
  s.addText("Angka pada Modul 6 adalah contoh latihan dan bukan angka PIM.", { x: 0.5, y: 6.72, w: 9, h: 0.28, fontSize: 10, color: C.text1, margin: 0, isTextBox: true, objectName: "Catatan contoh latihan" });

  // ===== bagian 3: hasil dan evaluasi =====
  pres.addSection({ title: "Hasil dan evaluasi" });
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Hasil dan evaluasi" });
  s.addText("Yang dibawa pulang peserta", { placeholder: "title" });
  const hasil = [
    [["FTE awal dan daftar data yang perlu dibuktikan", "FiTarget"], ["Daftar pekerjaan dan rencana pengumpulan data", "FiList"], ["Panduan wawancara dan catatan wawancara", "FiMessageSquare"], ["Lembar kerja tiga posisi", "FiFileText"]],
    [["Daftar masalah data yang sudah diputuskan", "FiSearch"], ["FTE akhir, heatmap, dan pohon akar masalah", "FiTrendingUp"], ["Matriks tindak lanjut dan RACI sebelum dan sesudah", "FiRepeat"], ["Rencana 30-60-90 hari dan komitmen peserta", "FiCheckCircle"]],
  ];
  hasil.forEach((rowItems, r) => {
    const y = 2.45 + r * 2.05;
    pill(s, r ? "HARI 2" : "HARI 1", 0.65, y - 0.42, 1.3, 0.32, { fill: C.accent2, size: 11, name: `Label hasil hari ${r + 1}` });
    rowItems.forEach(([t, ic], c) => {
      const x = 0.65 + c * 3.1;
      card(s, x, y, 2.7, 1.45, `Hasil ${r + 1}.${c + 1}`);
      s.addImage({ path: I(ic, "w"), x: x + 0.2, y: y + 0.18, w: 0.4, h: 0.4, objectName: `Ikon hasil ${r + 1}.${c + 1}` });
      s.addText(t, { x: x + 0.2, y: y + 0.65, w: 2.3, h: 0.8, fontSize: 14, bold: true, color: C.background1, valign: "top", margin: 0, isTextBox: true, objectName: `Teks hasil ${r + 1}.${c + 1}` });
    });
  });
  s.addShape(R.roundRect, { x: 0.65, y: 6.15, w: 12.0, h: 0.8, rectRadius: 0.2, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Panel perangkat" });
  s.addText("Diserahkan kepada PIM: lembar hitung FTE baku, seluruh formulir, lembar pengamatan, kunci kasus, dan formulir Level 4.", { x: 0.9, y: 6.15, w: 11.5, h: 0.8, fontSize: 14, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true, objectName: "Teks perangkat" });

  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Hasil dan evaluasi" });
  s.addText("Cara kami mengukur hasil belajar", { placeholder: "title" });
  const lv = [
    ["Level 1", "Reaksi", "Survei singkat tiap hari dan evaluasi akhir Hari 2.", "Rata-rata minimal 4,2 dari 5 dan 85% peserta menilai baik."],
    ["Level 2", "Pembelajaran", "Pre-test, post-test, latihan FTE, dan praktik wawancara.", "Post-test minimal 75 dan naik rata-rata 20% dari nilai awal."],
    ["Level 3", "Perilaku", "Daftar cek atasan, bukti penggunaan, dan online implementation clinic D+30.", "Minimal 75% peserta sasaran memakai hasil pelatihan pada satu aktivitas ABK."],
    ["Level 4", "Hasil", "Formulir satu halaman yang diserahkan di Hari 2 dan dipantau PIM.", "Kelengkapan data minimal 90%, isu berulang turun, rencana tindak punya penanggung jawab dan target waktu."],
  ];
  lv.forEach(([l, n, ins, tg], i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(R.roundRect, { x, y: 1.5, w: 2.9, h: 3.85, rectRadius: 0.15, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Kartu ${l}` });
    s.addShape(R.roundRect, { x, y: 1.5, w: 2.9, h: 0.9, rectRadius: 0.15, fill: { color: C.accent1 }, line: { type: "none" }, objectName: `Kepala ${l}` });
    s.addText([{ text: l, options: { fontSize: 12, color: C.accent2, bold: true, breakLine: true } }, { text: n, options: { fontSize: 18, color: C.background1, bold: true } }], { x: x + 0.2, y: 1.5, w: 2.5, h: 0.9, valign: "middle", margin: 0, isTextBox: true, objectName: `Judul ${l}` });
    s.addText([{ text: "Instrumen", options: { bold: true, color: C.accent6, fontSize: 13, breakLine: true } }, { text: ins, options: { fontSize: 14, color: C.text1, breakLine: true, paraSpaceAfter: 10 } }, { text: "Target", options: { bold: true, color: C.accent6, fontSize: 13, breakLine: true } }, { text: tg, options: { fontSize: 14, color: C.text1 } }], { x: x + 0.2, y: 2.55, w: 2.5, h: 2.7, valign: "top", margin: 0, isTextBox: true, objectName: `Isi ${l}` });
  });
  s.addText("Target Level 3 dan Level 4 adalah indikator bersama yang juga dipengaruhi data, dukungan atasan, dan tindak lanjut PIM. D+30 berarti 30 hari setelah pelatihan.", { x: 0.5, y: 5.65, w: 12.3, h: 0.6, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "Catatan evaluasi" });

  // ===== bagian 4: penutup =====
  pres.addSection({ title: "Penutup" });
  s = pres.addSlide({ masterName: "COVER", sectionTitle: "Penutup" });
  s.addShape(R.ellipse, { x: 8.8, y: 2.6, w: 6.4, h: 6.4, fill: { color: C.accent3, transparency: 80 }, line: { type: "none" }, objectName: "Dekorasi lingkaran besar" });
  s.addImage({ path: A("logo_white.png"), x: 0.65, y: 0.55, w: 2.4, h: 0.475, objectName: "Logo Sanggabiz" });
  s.addText("Mari mulai bersama Sanggabiz", { x: 0.65, y: 2.5, w: 11, h: 1.0, fontSize: 40, bold: true, color: C.background1, valign: "top", margin: 0, isTextBox: true, objectName: "Judul penutup" });
  s.addText("Kami siap menyesuaikan kasus, jadwal, dan peserta bersama PIM.", { x: 0.65, y: 3.6, w: 11, h: 0.5, fontSize: 18, color: C.accent5, margin: 0, isTextBox: true, objectName: "Ajakan" });
  [["FiPhone", "0877-7413-4288"], ["FiGlobe", "www.sanggabiz.com"], ["FiMail", "partnership@sanggabiz.com"]].forEach(([ic, t], i) => {
    const x = 0.65 + i * 4.0;
    s.addShape(R.roundRect, { x, y: 5.4, w: 3.7, h: 0.6, rectRadius: 0.3, fill: { type: "none" }, line: { color: C.background1, width: 1.5 }, objectName: `Kapsul kontak ${i + 1}` });
    s.addImage({ path: I(ic, "w"), x: x + 0.22, y: 5.55, w: 0.3, h: 0.3, objectName: `Ikon kontak ${i + 1}` });
    s.addText(t, { x: x + 0.65, y: 5.4, w: 2.95, h: 0.6, fontSize: 14, bold: true, color: C.background1, valign: "middle", margin: 0, isTextBox: true, objectName: `Kontak ${i + 1}` });
  });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ditulis:", OUT);
}
build().catch((e) => { console.error(e); process.exit(1); });
