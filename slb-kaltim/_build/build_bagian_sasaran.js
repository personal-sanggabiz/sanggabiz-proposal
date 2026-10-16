// v4: pengganti sub-bagian 2.3 "Peserta dan Cara Rekrutmen" (docx) dengan kuota per sekolah dari data nyata.
// Versi lama: build_bagian_sasaran_v1.js sampai _v3.js.
// Data seluruh SD dan SMP dibaca dari ../data/sd-smp-balikpapan.psv.
// Jalankan: NODE_PATH=<folder node_modules berisi docx> node build_bagian_sasaran.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Footer,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageNumber, LevelFormat, VerticalAlign, TabStopType,
} = require("docx");

const NAVY = "0B396D", INK = "1D1D1F", MUTED = "5C5F66", GOLD = "FFC00C", LINE = "C9D2DE", WARN = "9A3412";
const FONT = "Montserrat";
const W = 9638;

// ---------- data seluruh SD dan SMP (sumber: Sekolah Kita, semester 2026/2027 Gasal) ----------
const rows = fs.readFileSync(path.join(__dirname, "..", "data", "sd-smp-balikpapan.psv"), "utf8")
  .trim().split("\n").slice(1).map((l) => l.split("|"))
  .map(([npsn, nama, bp, st, kec, guru, siswa]) => ({ npsn, nama, bp, st, kec, guru: +guru, siswa: +siswa }));
const rp = (n) => n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
const sum = (a, f) => a.reduce((t, x) => t + f(x), 0);
const KEC = ["Selatan", "Timur", "Utara", "Tengah", "Barat", "Kota"];
const kecStat = KEC.map((k) => {
  const r = rows.filter((x) => x.kec === k);
  const sd = r.filter((x) => x.bp === "SD"), smp = r.filter((x) => x.bp === "SMP");
  return { k, sd: sd.length, smp: smp.length, guru: sum(r, (x) => x.guru), siswa: sum(r, (x) => x.siswa) };
});
const T = {
  n: rows.length, sd: rows.filter((x) => x.bp === "SD").length, smp: rows.filter((x) => x.bp === "SMP").length,
  neg: rows.filter((x) => x.st === "N").length, guru: sum(rows, (x) => x.guru), siswa: sum(rows, (x) => x.siswa),
};

// ---------- penolong teks (sama dengan build_proposal.js) ----------
function runs(text, base = {}) {
  const out = [];
  const re = /(\[ISI TIM:[^\]]*\])|(\*\*[^*]+\*\*)/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    if (m[1]) out.push(new TextRun({ text: m[1], ...base, bold: true, shading: { type: ShadingType.CLEAR, fill: "FFF200", color: "auto" }, color: INK }));
    else out.push(new TextRun({ text: m[2].slice(2, -2), ...base, bold: true }));
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...base }));
  return out;
}
const p = (text, o = {}) => new Paragraph({
  children: runs(text, o.run || {}), alignment: o.align || AlignmentType.LEFT,
  spacing: { after: o.after ?? 110, before: o.before ?? 0, line: o.line ?? 290 }, keepNext: o.keepNext, indent: o.indent,
});
const h2 = (text) => new Paragraph({ heading: "Heading2", children: [new TextRun({ text })], keepNext: true });
const h3 = (text) => new Paragraph({ children: [new TextRun({ text, bold: true, size: 20, color: NAVY })], spacing: { before: 200, after: 90 }, keepNext: true });
const bullet = (text) => new Paragraph({ numbering: { reference: "bul", level: 0 }, children: runs(text), spacing: { after: 70, line: 290 } });
const note = (text) => new Paragraph({ children: runs(text, { size: 16, color: MUTED }), spacing: { before: 60, after: 160, line: 260 } });
const step = (text) => p(text, { indent: { left: 360, hanging: 360 } });
const refP = (n, t) => new Paragraph({
  spacing: { after: 70, line: 260 }, indent: { left: 420, hanging: 420 }, tabStops: [{ type: TabStopType.LEFT, position: 420 }],
  children: [new TextRun({ text: `[${n}]\t${t}`, size: 16 })],
});
const pageBreak = () => new Paragraph({ children: [], pageBreakBefore: true, spacing: { after: 0 } });

// ---------- penolong tabel ----------
const thin = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: thin, bottom: thin, left: thin, right: thin };
function cell(content, width, o = {}) {
  const lines = Array.isArray(content) ? content : [content];
  return new TableCell({
    width: { size: width, type: WidthType.DXA }, borders,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    verticalAlign: VerticalAlign.TOP, margins: { top: 70, bottom: 70, left: 100, right: 100 },
    children: lines.map((t) => new Paragraph({
      alignment: o.align || AlignmentType.LEFT, spacing: { after: 30, line: 260 },
      children: runs(t, { size: o.size || 17, bold: o.bold, color: o.color || INK }),
    })),
  });
}
// o.right = indeks kolom rata kanan, o.total = baris terakhir diberi latar
function table(widths, header, body, o = {}) {
  const right = o.right || [];
  const head = new TableRow({ tableHeader: true, cantSplit: true, children: header.map((t, i) => cell(t, widths[i], { fill: NAVY, color: "FFFFFF", bold: true, size: 16, align: right.includes(i) ? AlignmentType.RIGHT : undefined })) });
  const trs = body.map((r, ri) => new TableRow({
    cantSplit: true,
    children: r.map((c, i) => cell(c, widths[i], {
      fill: o.total && ri === body.length - 1 ? "DCE6F2" : (ri % 2 ? "F7F9FC" : undefined),
      align: right.includes(i) ? AlignmentType.RIGHT : undefined,
    })),
  }));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows: [head, ...trs] });
}

// ---------- halaman petunjuk (untuk tim, dihapus sebelum ditempel) ----------
const petunjuk = [
  new Paragraph({
    children: [new TextRun({ text: "PETUNJUK PENEMPELAN (hapus halaman ini sebelum ditempel)", bold: true, size: 20, color: WARN })],
    spacing: { after: 140 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: WARN, space: 6 } },
  }),
  h2("Di mana menaruhnya"),
  step("1. **Ganti seluruh sub-bagian 2.3 Peserta dan Cara Rekrutmen** di Bab 2 dengan isi halaman berikutnya. Dokumen ini menggantikan v3. Tabel jadwal H-21 sampai Hari-H dari proposal v2 sudah dibawa ke 2.3.7."),
  step("2. **Tambahkan rujukan baru [16] sampai [28]** di akhir Lampiran C (halaman terakhir dokumen ini). Proposal v2 berhenti di rujukan [15]."),
  step("3. **Perbaiki satu kalimat di Bab 1.1.** Ganti \"Peraturan Wali Kota Nomor 3 Tahun 2013 tentang penyelenggaraan pendidikan inklusif [9]\" menjadi \"Peraturan Wali Kota Nomor 9 Tahun 2018 tentang Penyelenggaraan Pendidikan Inklusif [18]\"."),
  step("4. **Angka siswa SLB Negeri Balikpapan di Bab 1.1 dan Gambar 2** (401 siswa, dari berita ANTARA) lebih lama daripada data Kemendikdasmen semester 2026/2027 Gasal (381 siswa). Ganti bila ingin angka seragam."),
  step("5. **Lampiran D** (file terpisah) memuat data seluruh 274 SD dan SMP di Balikpapan. Tempel setelah Lampiran C bila klien meminta data lengkap."),
  step("6. **Surat Pengantar.** Instansi penerima yang tepat ada dua: Kepala Cabang Dinas Pendidikan Wilayah I Provinsi Kalimantan Timur dan Kepala Dinas Pendidikan dan Kebudayaan Kota Balikpapan. Kontaknya ada di 2.3.6."),
  h2("Yang berubah dari v3"),
  p("Kuota sekarang per sekolah dengan aturan yang bisa dihitung ulang, memakai data guru dan siswa nyata. Kategori kini bernama A sampai D dengan penjelasan di 2.3.1, dan kolom Kuota di tabel sekolah diganti Guru diundang. SDN 010 Balikpapan Selatan dikeluarkan karena berita yang menyebutnya tidak menyebut kecamatan dan SDN 010 ada di keenam kecamatan. Penggantinya SDN 008 Balikpapan Kota. SMP Negeri 4 dan SD IT Luqman Al-Hakim ditambahkan. Delapan SMA dan SMK inklusif kini lengkap dengan alamat, kontak, guru, dan siswa."),
  h2("Isian yang tersisa"),
  p("Dua tanda kuning di 2.3.3: sekolah inklusif di Kecamatan Balikpapan Timur dan Balikpapan Utara. Namanya ada di Keputusan Wali Kota yang belum terbit di sumber publik."),
  h2("Cek sebelum dikirim"),
  p("Telepon semua sekolah, karena kolom telepon dan faks SLB di data sumber berisi nomor yang sama. Konfirmasi status inklusif SDN 008 Balikpapan Kota, SMP Negeri 4, dan SMP Negeri 5 ke Disdikbud Kota, karena sumbernya berita 2019 dan jurnal 2023. Cek apakah Juknis SPMB 2026 mengubah daftar SMA dan SMK. Jumlah guru dan siswa berasal dari Sekolah Kita semester 2026/2027 Gasal dan bisa berubah. Kuota adalah usulan, ubah bila klien menetapkan lain."),
];

// ---------- isi pengganti 2.3 ----------
const isi = [
  pageBreak(),
  h2("2.3 Peserta dan Cara Rekrutmen"),
  p("Peserta adalah 80 guru SLB dan sekolah inklusif di Balikpapan, dari semua jenjang dan semua jenis kekhususan. Peserta tidak membayar, karena biaya ditanggung pemberi dana. Sanggabiz yang mencari dan mengonfirmasi peserta."),
  p(`Data yang kami pakai berasal dari sumber resmi. Balikpapan memiliki empat SLB dengan 80 guru dan 682 siswa. Kota ini juga memiliki ${T.sd} SD dan ${T.smp} SMP yang memuat ${rp(T.guru)} guru dan ${rp(T.siswa)} siswa, serta delapan SMA dan SMK negeri penyelenggara pendidikan inklusif dengan 465 guru dan 9.468 siswa [16, 21]. Guru SLB saja tidak mencukupi untuk 80 peserta, sehingga kami juga mengundang guru dari sekolah inklusif. Data diakses pada 15 Oktober 2026.`),

  h3("2.3.1 Kategori sekolah sasaran dan jumlah guru yang diundang"),
  p("Kami mengelompokkan sekolah sasaran ke empat kategori menurut jenis sekolah dan instansi pembinanya. **Kategori A** adalah SLB, yaitu sekolah khusus yang seluruh muridnya berkebutuhan khusus. **Kategori B dan C** adalah sekolah inklusif, yaitu sekolah biasa yang menerima murid berkebutuhan khusus bersama murid lain. Kategori B memuat SD dan SMP, kategori C memuat SMA dan SMK. **Kategori D** bukan penyedia peserta. Isinya kelompok guru, komunitas orang tua, dan mitra yang membantu menyebarkan undangan."),
  p("Kolom Guru diundang menyebut jumlah guru yang kami undang dari tiap sekolah. Angka itu sekaligus menjadi kuota peserta, dan jumlahnya 80 orang."),
  p("Pendidikan khusus dan pendidikan menengah menjadi urusan pemerintah provinsi. SD dan SMP menjadi urusan pemerintah kota [20]. Karena itu undangan kami berjalan lewat dua instansi: Cabang Dinas Pendidikan Wilayah I yang mewakili Dinas Pendidikan dan Kebudayaan Provinsi Kalimantan Timur di Balikpapan, dan Dinas Pendidikan dan Kebudayaan Kota Balikpapan."),
  table([1300, 1700, 800, 1050, 800, 900, 3088],
    ["Kategori", "Jenis sekolah", "Sekolah", "Guru tercatat", "Siswa", "Guru diundang", "Aturan penentuan jumlah guru diundang"],
    [
      ["**A. SLB**", "Sekolah khusus", "4", "80", "682", "40", "Sekitar separuh guru tiap SLB, dibulatkan. SLB Negeri 22, Tunas Bangsa 7, Dharma Kencana 5, Santo Fransiskus Assisi 6."],
      ["**B. SD dan SMP inklusif**", "Sekolah inklusif", "6 kecamatan", "156 di 5 sekolah tercatat", "3.585", "24", "Empat guru per kecamatan, karena Peraturan Wali Kota mewajibkan minimal satu sekolah inklusif di tiap kecamatan [18]. Satu tim terdiri atas guru kelas atau guru mata pelajaran dan GPK."],
      ["**C. SMA dan SMK inklusif**", "Sekolah inklusif", "8", "465", "9.468", "16", "Dua guru per sekolah, yaitu guru mata pelajaran dan GPK atau guru bimbingan konseling."],
      ["**D. Jaringan pendukung**", "KKG, komunitas, mitra", "-", "-", "-", "Tanpa kuota", "Membantu menyebarkan undangan dan menjadi mitra (2.3.5)."],
      ["**Total peserta**", "", "", "", "", "**80**", ""],
    ], { right: [2, 3, 4, 5], total: true }),
  note("Jumlah guru diundang adalah usulan Sanggabiz dan dapat disesuaikan dengan permintaan klien. Sekolah menunjuk sendiri gurunya. Jumlah guru dan siswa: portal Sekolah Kita, Kemendikdasmen, semester 2026/2027 Gasal [21]."),

  h3("2.3.2 Kategori A: empat SLB di Kota Balikpapan"),
  p("Referensi Kemendikdasmen mencatat empat SLB di Balikpapan, dua di Kecamatan Balikpapan Selatan dan dua di Kecamatan Balikpapan Kota [16]. Seluruhnya kami undang. Jumlah guru yang diundang dari tiap sekolah sekitar separuh dari gurunya, agar kegiatan belajar di sekolah tetap berjalan."),
  table([2100, 2300, 2300, 700, 700, 700, 838],
    ["Sekolah", "Alamat", "Kontak", "Guru", "Siswa", "Rombel", "Guru diundang"],
    [
      [["**SLB Negeri Balikpapan**", "NPSN 30401525, akreditasi A", "Negeri, TKLB sampai SMLB"], "Jl. Letkol Pol. HM. Asnawi Arbain RT 14, Kel. Sepinggan Raya, Kec. Balikpapan Selatan", ["Telepon 0542-8879984", "slbn.balikpapan@yahoo.com", "slbnbalikpapan.sch.id"], "44", "381", "44", "22"],
      [["**SLB Tunas Bangsa**", "NPSN 30402883, akreditasi B", "Swasta, SDLB sampai SMLB"], "Jl. Kutilang RSS Damai III RT 28, Kel. Gunung Bahagia, Kec. Balikpapan Selatan", ["Telepon 0542-8870057 *", "tunasbangsa831@yahoo.co.id", "slbtunasbangsabalikpapan.sch.id"], "13", "85", "13", "7"],
      [["**SLB Dharma Kencana Balikpapan**", "NPSN 30402882, akreditasi B", "Swasta, SDLB sampai SMLB"], "Jl. Pembangunan No. 30 Blok A, Kel. Telaga Sari, Kec. Balikpapan Kota", ["Telepon 0542-410406", "slbdk@yahoo.co.id"], "10", "111", "20", "5"],
      [["**Sekolah Khusus Santo Fransiskus Assisi**", "NPSN 69965143, akreditasi B", "Swasta, TKLB sampai SMLB"], "Jl. Mayor (Pol) Zaenal Arifin No. 40 RT 18, Kel. Damai, Kec. Balikpapan Kota", ["Telepon 0542-8794310", "skhfransiskusassisi@gmail.com"], "13", "105", "14", "6"],
      ["**Jumlah empat SLB**", "", "", "**80**", "**682**", "**91**", "**40**"],
    ], { right: [3, 4, 5, 6], total: true }),
  note("* Data sumber mencantumkan nomor 8870057 tanpa kode area. Kami menambahkan kode area Balikpapan (0542). Sumber: Referensi Kemendikdasmen [16] dan Sekolah Kita [21]. Rombel adalah rombongan belajar."),

  h3("2.3.3 Kategori B: SD dan SMP inklusif"),
  p("Peraturan Wali Kota Balikpapan Nomor 9 Tahun 2018 mewajibkan setiap kecamatan menyelenggarakan paling sedikit satu pendidikan inklusif pada jenjang SD dan SMP, dan menetapkan sekolahnya dengan Keputusan Wali Kota (Pasal 4 ayat 3 dan ayat 5) [18]. Pasal 12 sampai 14 membagi pendidik inklusif menjadi guru kelas atau guru mata pelajaran dan GPK. Kami mengundang keduanya, karena modul ajar berdiferensiasi disusun guru kelas bersama GPK."),
  p("Daftar resmi Keputusan Wali Kota belum terbit di sumber publik. Sebagai gantinya, kami memakai sekolah yang disebut pejabat Dinas Pendidikan dan Kebudayaan Kota Balikpapan atau penelitian yang terbit. Kepala Dinas menyebut SDN 008 Balikpapan Kota, SMP Negeri 5, dan SMA Negeri 4 pada 12 Juli 2019 [26], lalu SDN 08, SMP Negeri 4, dan SMP Negeri 5 pada 15 Juli 2019 [27]. Penelitian tahun 2023 menyebut jalur inklusi pada awalnya diterapkan di SDN 08, SMPN 4, dan SMPN 5, serta mengkaji praktik inklusi di SD IT Luqman Al-Hakim dan SD IT Assalaam [28]. SDN 003 Balikpapan Kota menggelar sosialisasi sekolah inklusi pada 19 November 2024 [23]."),
  table([2100, 2000, 2000, 700, 800, 700, 1338],
    ["Sekolah", "Alamat", "Kontak", "Guru", "Siswa", "Guru diundang", "Kecamatan dan dasar"],
    [
      [["**SD Negeri 008 Balikpapan Kota**", "NPSN 30402860, akreditasi A"], "Jl. Jend. Sudirman RT 22 No. 03", ["Telepon 0542-410201", "sdn.008balikpapankota@gmail.com"], "15", "302", "2", ["Balikpapan Kota", "Kepala Disdikbud 2019 [26, 27]"]],
      [["**SD Negeri 003 Balikpapan Kota**", "NPSN 30401549, akreditasi A"], "Jl. Wiluyo Puspoyudo", ["Telepon belum tercatat", "sdn003_balsel@yahoo.co.id"], "29", "639", "2", ["Balikpapan Kota", "Sosialisasi inklusi 2024 [23]"]],
      [["**SMP Negeri 5 Balikpapan**", "NPSN 30401474, akreditasi A"], "Jl. Prona II No. 104 RT 24", ["Telepon 0542-764142", "balikpapansmpn5@gmail.com", "smp5bpn.sch.id"], "49", "1.287", "4", ["Balikpapan Selatan", "Semua sumber [26, 27, 28]"]],
      [["**SMP Negeri 4 Balikpapan**", "NPSN 30401475, akreditasi A"], "Jl. Bukit Pelajar I, Sidodadi", ["Telepon 0542-734042", "smp4_bpp@yahoo.com"], "50", "1.195", "4", ["Balikpapan Barat", "Disdikbud 2019, jurnal 2023 [27, 28]"]],
      [["**SD IT Luqman Al-Hakim**", "NPSN 30402848, akreditasi A, swasta"], "Jl. Sultan Alauddin RT 01 No. 26, Karang Bugis Dalam", ["Telepon 0542-441930", "sditluqmanalhakimbpp@gmail.com", "luqmanalhakim-bpp.sch.id"], "13", "162", "4", ["Balikpapan Tengah", "Penelitian inklusi 2023 [28]"]],
      [["**[ISI TIM: sekolah inklusif Kec. Balikpapan Timur]**"], "Menunggu Keputusan Wali Kota", "Menunggu", "-", "-", "4", ["Balikpapan Timur", "Perwali 9/2018 [18]"]],
      [["**[ISI TIM: sekolah inklusif Kec. Balikpapan Utara]**"], "Menunggu Keputusan Wali Kota. SD IT Assalaam dikaji penelitian 2023 [28]", "Menunggu", "-", "-", "4", ["Balikpapan Utara", "Perwali 9/2018 [18]"]],
      ["**Jumlah**", "", "", "**156**", "**3.585**", "**24**", "Guru dan siswa dari 5 sekolah tercatat"],
    ], { right: [3, 4, 5], total: true }),
  note("Berita IDN Times 10 Februari 2020 menyebut SDN 010 tanpa kecamatan [22]. SDN 010 ada di keenam kecamatan, sehingga kami belum memasukkannya sebelum dinas mengonfirmasi. Sumber kontak, guru, dan siswa: Sekolah Kita [21]."),
  p("Jumlah sekolah di tiap kecamatan menunjukkan besarnya kelompok guru yang dapat dijangkau bila klien menghendaki sekolah inklusif lain. Data lengkap tiap sekolah ada di Lampiran D."),
  table([2300, 1100, 1100, 1700, 1700, 1738],
    ["Kecamatan", "SD", "SMP", "Guru SD dan SMP", "Siswa SD dan SMP", "SLB"],
    [
      ...kecStat.map((s) => [`Balikpapan ${s.k}`, String(s.sd), String(s.smp), rp(s.guru), rp(s.siswa), s.k === "Selatan" || s.k === "Kota" ? "2" : "0"]),
      ["**Jumlah**", `**${T.sd}**`, `**${T.smp}**`, `**${rp(T.guru)}**`, `**${rp(T.siswa)}**`, "**4**"],
    ], { right: [1, 2, 3, 4, 5], total: true }),
  note(`Jumlah sekolah negeri dan swasta, tidak termasuk madrasah dan sekolah kerja sama. Dari ${T.n} sekolah itu, ${T.neg} berstatus negeri. Sumber: Sekolah Kita [21], dicocokkan dengan Referensi Kemendikdasmen yang mencatat 198 SD dan 81 SMP [16]. Selisih tiga SD dan dua SMP berasal dari sekolah yang belum muncul di Sekolah Kita atau tercatat sebagai sekolah kerja sama.`),

  h3("2.3.4 Kategori C: SMA dan SMK inklusif"),
  p("Delapan sekolah negeri berikut tercantum sebagai penyelenggara pendidikan inklusif pada Lampiran 9 Petunjuk Teknis SPMB SMA, SMK, dan SLB Kota Balikpapan Tahun 2025 [17]. Penetapannya mengacu pada Surat Keputusan Kepala Dinas Pendidikan dan Kebudayaan Provinsi Kalimantan Timur Nomor 400.3/12060/DISDIKBUD.V/2023. Setiap sekolah mengirim dua guru."),
  table([2000, 2300, 2100, 800, 800, 700, 938],
    ["Sekolah", "Alamat", "Kontak", "Guru", "Siswa", "Guru diundang", "Kecamatan"],
    [
      [["**SMA Negeri 1 Balikpapan**", "NPSN 30401491, akreditasi A"], "Jl. Kapten Piere Tendean RT 30 No. 63", ["Telepon 0542-422742", "sman1bpp@gmail.com"], "57", "1.369", "2", "Kota"],
      [["**SMA Negeri 2 Balikpapan**", "NPSN 30401490, akreditasi A"], "Jl. Soekarno Hatta, Strat IV", ["Telepon 0542-424686", "school@sman2bpp.sch.id"], "51", "1.052", "2", "Utara"],
      [["**SMA Negeri 4 Balikpapan**", "NPSN 30401488, akreditasi A"], "Jl. Sepinggan Baru III No. 36", ["Telepon 0542-780044", "sman4bpn@gmail.com"], "47", "1.186", "2", "Selatan"],
      [["**SMA Negeri 7 Balikpapan**", "NPSN 30401504, akreditasi A"], "Jl. Mulawarman", ["Telepon 0542-790025", "sma7balikpapan@gmail.com"], "53", "1.089", "2", "Timur"],
      [["**SMA Negeri 8 Balikpapan**", "NPSN 30402920, akreditasi A"], "Jl. AMD Gn. IV RT 14", ["Telepon 0542-7584515", "info@sman8bpp.sch.id"], "45", "869", "2", "Barat"],
      [["**SMK Negeri 4 Balikpapan**", "NPSN 30401508, akreditasi A"], "Jl. Belibis 3 No. 01", ["Telepon 0542-873890", "school@smkn4balikpapan.sch.id"], "67", "1.260", "2", "Selatan"],
      [["**SMK Negeri 5 Balikpapan**", "NPSN 30404599, akreditasi B"], "Jl. Mulawarman No. 65", ["Telepon 0542-790767", "smkn5bpn@gmail.com"], "70", "1.229", "2", "Timur"],
      [["**SMK Negeri 6 Balikpapan**", "NPSN 30409918, akreditasi B"], "Jl. Soekarno Hatta KM 7 RT 61, Graha Indah", ["Telepon 0542-8830033", "sekolah@smkn6-bpn.sch.id"], "75", "1.414", "2", "Utara"],
      ["**Jumlah delapan sekolah**", "", "", "**465**", "**9.468**", "**16**", ""],
    ], { right: [3, 4, 5], total: true }),
  note("Sumber kontak, guru, dan siswa: Sekolah Kita, semester 2026/2027 Gasal [21]. Nomor telepon SMA Negeri 7 (790025) dan SMA Negeri 8 (7584515) tercatat tanpa kode area, sehingga kami menambahkan 0542."),

  h3("2.3.5 Kategori D: jaringan pendukung"),
  p("Kategori ini tidak memiliki kuota. Mereka membantu undangan dan menjadi mitra bila klien menghendaki."),
  bullet("**KKG PAI Balikpapan Kota.** Kelompok kerja guru ini menggelar bimbingan teknis strategi pendidikan inklusi dalam pembelajaran PAI di SDN 011 Balikpapan Kota pada 26 Juli 2025 [24]. Materinya mencakup pendekatan diferensiasi, sehingga anggotanya menjadi sasaran undangan yang sesuai."),
  bullet("**ABK Hebat Balikpapan.** Komunitas orang tua anak berkebutuhan khusus yang diketuai Nuri Cahyani menurut berita 2020. Data Dinas Kesehatan Balikpapan pada 2020 memperkirakan 400 sampai 500 anak berkebutuhan khusus di kota ini [22]. Komunitas dapat menyampaikan undangan ke sekolah anak mereka."),
  bullet("**Guru SLB Negeri Balikpapan sebagai narasumber lapangan.** Salah satu guru SLBN menjadi narasumber sosialisasi inklusi di SDN 003 Balikpapan Kota [23]. SLB sudah dilibatkan sebagai rujukan bagi sekolah inklusif, dan kami mengundang mereka dalam peran yang sama."),
  bullet("**Pembinaan provinsi.** Dinas Pendidikan dan Kebudayaan Provinsi Kalimantan Timur melatih 60 kepala sekolah dari 10 kabupaten dan kota dalam manajemen pendidikan inklusif di Balikpapan pada 20 sampai 22 Februari 2024 [25]. Peserta pelatihan itu berasal dari seluruh Kaltim, sehingga jaringan alumninya dapat membantu menyebarkan undangan."),

  h3("2.3.6 Kontak instansi pembina"),
  table([2300, 3100, 2438, 1800],
    ["Instansi", "Alamat", "Kontak", "Peran dalam program"],
    [
      [["**Cabang Dinas Pendidikan Wilayah I**", "Dinas Pendidikan dan Kebudayaan Provinsi Kalimantan Timur"], "Jl. A.W. Syahrani, Batu Ampar, Kec. Balikpapan Utara. Wilayah kerja: Kota Balikpapan dan Kabupaten Penajam Paser Utara", ["Surel cabdindik.kaltim1@gmail.com", "cabdin1.disdikbud.kaltimprov.go.id"], "Rekomendasi dan undangan untuk kategori A dan C"],
      [["**Dinas Pendidikan dan Kebudayaan Kota Balikpapan**"], "Jl. Ruhui Rahayu I, Kel. Sepinggan, Kec. Balikpapan Selatan", ["Telepon 0812-5591-5591", "Surel disdik@balikpapan.go.id"], "Rekomendasi, daftar sekolah, dan undangan untuk kategori B"],
    ]),
  note("Sumber: Petunjuk Teknis SPMB 2025 untuk Cabang Dinas [17] dan situs Dinas Pendidikan dan Kebudayaan Kota Balikpapan [19]."),

  h3("2.3.7 Jalur dan jadwal rekrutmen"),
  p("Jalur yang kami pakai:", { after: 60 }),
  bullet("surat permohonan rekomendasi dan data sekolah kepada Cabang Dinas Pendidikan Wilayah I dan Dinas Pendidikan dan Kebudayaan Kota Balikpapan, lalu surat undangan dari dinas melalui pemberi dana,"),
  bullet("undangan resmi kepada kepala sekolah di kategori A sampai C. Tiap sekolah menunjuk guru yang memegang kelas dengan murid berkebutuhan khusus, dan GPK diprioritaskan,"),
  bullet("kelompok kerja guru dan musyawarah kerja kepala sekolah (KKG dan MKKS), serta jaringan pada kategori D,"),
  bullet("jaringan mitra lokal dan kampus di Kalimantan Timur."),
  p("Calon peserta mengisi formulir pendaftaran daring. Kuota tiap sekolah dan tiap kategori ditutup bila penuh, dan sisa kuota dialihkan ke sekolah lain yang masih memiliki pendaftar. Jawaban formulir menjadi pemetaan awal untuk pemateri.", { before: 60 }),
  table([2200, 7438], ["Waktu", "Kegiatan rekrutmen"], [
    ["H-21", "Undangan dikirim ke dinas, KKG, MKKS, dan kepala sekolah. Formulir pendaftaran daring dibuka"],
    ["H-14", "Pendaftaran ditutup bila kuota penuh. Formulir mencatat jenjang, mata pelajaran, jenis kebutuhan murid di kelas, kendala utama, dan kebutuhan akses peserta"],
    ["H-7", "Konfirmasi kehadiran lewat WhatsApp. Peserta cadangan dihubungi bila ada yang berhalangan"],
    ["H-3", "Pengingat berisi lokasi, jam, dan petunjuk akses"],
    ["Hari-H", "Registrasi dan daftar hadir"],
  ]),

  pageBreak(),
  h2("Tambahan untuk Lampiran C. Rujukan"),
  p("Tempel setelah rujukan [15].", { run: { size: 16, color: MUTED } }),
  refP(16, "Kemendikdasmen. Referensi Data Pendidikan, halaman satuan pendidikan per NPSN 30401525, 30402883, 30402882, dan 69965143, daftar program SLB Kota Balikpapan, serta daftar satuan pendidikan dikdas per kecamatan. https://referensi.data.kemendikdasmen.go.id. Diakses 15 Oktober 2026."),
  refP(17, "Cabang Dinas Pendidikan Wilayah I Provinsi Kalimantan Timur (2025). Petunjuk Teknis SPMB SMA, SMK, dan SLB Kota Balikpapan Tahun Pelajaran 2025/2026, Lampiran 9. SK Nomor 400.3.8/370.2/Cabdisdikwil.1/2025, 22 April 2025. https://berkas.spmb.id/cabdinbalikpapan/juknis.250521095215.pdf"),
  refP(18, "Peraturan Wali Kota Balikpapan Nomor 9 Tahun 2018 tentang Penyelenggaraan Pendidikan Inklusif. https://peraturan.bpk.go.id/Download/88906/PERWALI%20NOMOR%209%202018.pdf"),
  refP(19, "Dinas Pendidikan dan Kebudayaan Kota Balikpapan. Halaman kontak. https://disdikbud.balikpapan.go.id. Diakses 15 Oktober 2026."),
  refP(20, "Undang-Undang Nomor 23 Tahun 2014 tentang Pemerintahan Daerah, Lampiran huruf A, pembagian urusan pemerintahan bidang pendidikan."),
  refP(21, "Kemendikdasmen. Sekolah Kita, profil seluruh SD, SMP, SLB, dan SMA atau SMK sasaran di Kota Balikpapan, semester 2026/2027 Gasal. https://sekolah.data.kemendikdasmen.go.id. Diakses 15 Oktober 2026."),
  refP(22, "IDN Times Kaltim (10 Februari 2020). Kurang Tenaga Pengajar, Pemkot Balikpapan Akan Bangun Sekolah Inklusi. https://kaltim.idntimes.com/news/kaltim/melani-indra-hapsari/kurang-tenaga-pengajar-pemkot-balikpapan-akan-bangun-sekolah-inklusi"),
  refP(23, "Balpos (23 November 2024). Sosialisasi Sekolah Inklusi di SDN 003 Balikpapan Kota: Wujudkan Pendidikan Ramah ABK. https://www.balpos.com/pendidikan/1795344660/sosialisasi-sekolah-inklusi-di-sdn-003-balikpapan-kota-wujudkan-pendidikan-ramah-abk"),
  refP(24, "KKG PAI Balikpapan (27 Juli 2025). Bimtek Pendidikan Inklusi: Wujudkan Pembelajaran PAI yang Ramah dan Adaptif. http://www.kkgpaibalikpapan.com/2025/07/bimtek-pendidikan-inklusi-wujudkan.html"),
  refP(25, "Dinas Pendidikan dan Kebudayaan Provinsi Kalimantan Timur (28 Februari 2024). Bidang PPK Gelar Manajemen Pengelolaan Pendidikan Inklusif untuk Kepala Sekolah. https://web.disdikbud.kaltimprov.go.id/berita/detail/bidang-ppk-gelar-manajemen-pengelolaan-pendidikan-inklusif-untuk-kepala-sekolah-Ae8qI"),
  refP(26, "KPFM Balikpapan (12 Juli 2019). Kelas Inklusi di Setiap Kecamatan. https://kapefm.com/kelas-inklusi-di-setiap-kecamatan/"),
  refP(27, "IDN Times Kaltim (15 Juli 2019). Pada 2020, Balikpapan Bakal Terapkan Sekolah Inklusi Tiap Kecamatan. https://kaltim.idntimes.com/news/kalimantan-timur/pada-2020-balikpapan-bakal-terapkan-sekolah-inklusi-tiap-kecamatan-00-bgzwd-6ljst2"),
  refP(28, "Julianti, K., Zurqoni, dan Zamroni (2023). Model Pendidikan Inklusi pada Mata Pelajaran Pendidikan Agama Islam: Studi Multisitus SDIT Luqman Al-Hakim dan SDIT Assalaam. Jurnal Kolaboratif Sains, 6(10), 1371-1372. https://jurnal.unismuhpalu.ac.id/index.php/JKS/article/download/4296/3238"),
];

const footer = new Footer({
  children: [new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: LINE, space: 6 } },
    tabStops: [{ type: TabStopType.RIGHT, position: W }],
    children: [
      new TextRun({ text: "Sanggabiz  |  Workshop Pembelajaran Berdiferensiasi untuk Guru SLB dan Sekolah Inklusif", size: 14, color: MUTED }),
      new TextRun({ text: "\t", size: 14 }),
      new TextRun({ children: ["Halaman ", PageNumber.CURRENT], size: 14, color: MUTED }),
    ],
  })],
});

const doc = new Document({
  creator: "Sanggabiz", title: "Pengganti Bab 2.3 Peserta dan Cara Rekrutmen",
  styles: {
    default: { document: { run: { font: FONT, size: 19, color: INK } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: NAVY },
        paragraph: { spacing: { before: 0, after: 200 }, outlineLevel: 0, border: { bottom: { style: BorderStyle.SINGLE, size: 14, color: GOLD, space: 6 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 22, bold: true, color: NAVY },
        paragraph: { spacing: { before: 240, after: 110 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [{ reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 260 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: footer }, children: [...petunjuk, ...isi],
  }],
});

const outDir = path.join(__dirname, "..", "bagian-sasaran");
fs.mkdirSync(outDir, { recursive: true });
const out = path.join(outDir, "v4 Pengganti Bab 2.3 Peserta dan Sekolah Sasaran - SLB dan Inklusif Balikpapan.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ditulis:", out); console.log("SD/SMP:", T, kecStat); });
