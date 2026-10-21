// Membuat versi ringkas Bab 2.3: satu kalimat dan satu tabel sekolah sasaran (docx), untuk ditempel manual ke proposal utama.
// Data sama dengan build_bagian_sasaran.js (v4). Jalankan: NODE_PATH=<folder node_modules berisi docx> node build_tabel_sekolah.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Footer,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign, TabStopType,
} = require("docx");

const NAVY = "0B396D", INK = "1D1D1F", MUTED = "5C5F66", GOLD = "FFC00C", LINE = "C9D2DE";
const FONT = "Montserrat";
const W = 9638;

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
const p = (text, o = {}) => new Paragraph({ children: runs(text, o.run || {}), spacing: { after: o.after ?? 110, before: o.before ?? 0, line: o.line ?? 290 }, keepNext: o.keepNext });
const note = (text) => new Paragraph({ children: runs(text, { size: 15, color: MUTED }), spacing: { before: 60, after: 80, line: 250 } });
const h2 = (text) => new Paragraph({ heading: "Heading2", children: [new TextRun({ text })], keepNext: true });

const thin = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: thin, bottom: thin, left: thin, right: thin };
const SZ = 15;
function cell(content, width, o = {}) {
  const lines = Array.isArray(content) ? content : [content];
  return new TableCell({
    width: { size: width, type: WidthType.DXA }, borders, columnSpan: o.span,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    verticalAlign: VerticalAlign.TOP, margins: { top: 50, bottom: 50, left: 70, right: 70 },
    children: lines.map((t) => new Paragraph({
      alignment: o.right ? AlignmentType.RIGHT : AlignmentType.LEFT, spacing: { after: 20, line: 235 },
      children: runs(t, { size: o.size || SZ, bold: o.bold, color: o.color || INK }),
    })),
  });
}

const widths = [380, 2300, 850, 1900, 2000, 560, 640, 1008];
const header = ["No", "Sekolah", "Kecamatan", "Alamat", "Kontak", "Guru", "Siswa", "Guru diundang"];
const right = [5, 6, 7];
let no = 0;
const row = (nm, kec, alamat, kontak, guru, siswa, undang) => {
  no += 1;
  const cs = [String(no), nm, kec, alamat, kontak, guru, siswa, undang];
  return new TableRow({ cantSplit: true, children: cs.map((c, i) => cell(c, widths[i], { right: right.includes(i), fill: no % 2 === 0 ? "F7F9FC" : undefined })) });
};
const group = (text) => new TableRow({ cantSplit: true, children: [cell(text, W, { span: widths.length, fill: "DCE6F2", bold: true })] });
const sub = (label, g, s, u) => new TableRow({ cantSplit: true, children: [
  cell("", widths[0], { fill: "EEF3F9" }), cell(label, widths[1] + widths[2] + widths[3] + widths[4], { span: 4, fill: "EEF3F9", bold: true }),
  cell(g, widths[5], { fill: "EEF3F9", bold: true, right: true }), cell(s, widths[6], { fill: "EEF3F9", bold: true, right: true }), cell(u, widths[7], { fill: "EEF3F9", bold: true, right: true }),
] });

const head = new TableRow({ tableHeader: true, cantSplit: true, children: header.map((t, i) => cell(t, widths[i], { fill: NAVY, color: "FFFFFF", bold: true, right: right.includes(i) })) });
const rows = [
  head,
  group("A. SLB (sekolah khusus)"),
  row(["**SLB Negeri Balikpapan**", "NPSN 30401525, negeri, akreditasi A"], "Selatan", "Jl. Letkol Pol. HM. Asnawi Arbain RT 14, Kel. Sepinggan Raya", ["0542-8879984", "slbn.balikpapan@yahoo.com"], "44", "381", "22"),
  row(["**SLB Tunas Bangsa**", "NPSN 30402883, swasta, akreditasi B"], "Selatan", "Jl. Kutilang RSS Damai III RT 28, Kel. Gunung Bahagia", ["0542-8870057 *", "tunasbangsa831@yahoo.co.id"], "13", "85", "7"),
  row(["**SLB Dharma Kencana Balikpapan**", "NPSN 30402882, swasta, akreditasi B"], "Kota", "Jl. Pembangunan No. 30 Blok A, Kel. Telaga Sari", ["0542-410406", "slbdk@yahoo.co.id"], "10", "111", "5"),
  row(["**Sekolah Khusus Santo Fransiskus Assisi**", "NPSN 69965143, swasta, akreditasi B"], "Kota", "Jl. Mayor (Pol) Zaenal Arifin No. 40 RT 18, Kel. Damai", ["0542-8794310", "skhfransiskusassisi@gmail.com"], "13", "105", "6"),
  sub("Jumlah kategori A", "80", "682", "40"),
  group("B. SD dan SMP inklusif (sekolah biasa yang menerima murid berkebutuhan khusus)"),
  row(["**SD Negeri 008 Balikpapan Kota**", "NPSN 30402860, negeri, akreditasi A"], "Kota", "Jl. Jend. Sudirman RT 22 No. 03", ["0542-410201", "sdn.008balikpapankota@gmail.com"], "15", "302", "2"),
  row(["**SD Negeri 003 Balikpapan Kota**", "NPSN 30401549, negeri, akreditasi A"], "Kota", "Jl. Wiluyo Puspoyudo", ["Telepon belum tercatat", "sdn003_balsel@yahoo.co.id"], "29", "639", "2"),
  row(["**SMP Negeri 5 Balikpapan**", "NPSN 30401474, negeri, akreditasi A"], "Selatan", "Jl. Prona II No. 104 RT 24", ["0542-764142", "balikpapansmpn5@gmail.com"], "49", "1.287", "4"),
  row(["**SMP Negeri 4 Balikpapan**", "NPSN 30401475, negeri, akreditasi A"], "Barat", "Jl. Bukit Pelajar I, Sidodadi", ["0542-734042", "smp4_bpp@yahoo.com"], "50", "1.195", "4"),
  row(["**SD IT Luqman Al-Hakim**", "NPSN 30402848, swasta, akreditasi A"], "Tengah", "Jl. Sultan Alauddin RT 01 No. 26, Karang Bugis Dalam", ["0542-441930", "sditluqmanalhakimbpp@gmail.com"], "13", "162", "4"),
  row(["**[ISI TIM: sekolah inklusif Kec. Balikpapan Timur]**"], "Timur", "Menunggu Keputusan Wali Kota", "Menunggu", "-", "-", "4"),
  row(["**[ISI TIM: sekolah inklusif Kec. Balikpapan Utara]**"], "Utara", "Menunggu Keputusan Wali Kota", "Menunggu", "-", "-", "4"),
  sub("Jumlah kategori B (guru dan siswa dari 5 sekolah tercatat)", "156", "3.585", "24"),
  group("C. SMA dan SMK inklusif"),
  row(["**SMA Negeri 1 Balikpapan**", "NPSN 30401491, negeri, akreditasi A"], "Kota", "Jl. Kapten Piere Tendean RT 30 No. 63", ["0542-422742", "sman1bpp@gmail.com"], "57", "1.369", "2"),
  row(["**SMA Negeri 2 Balikpapan**", "NPSN 30401490, negeri, akreditasi A"], "Utara", "Jl. Soekarno Hatta, Strat IV", ["0542-424686", "school@sman2bpp.sch.id"], "51", "1.052", "2"),
  row(["**SMA Negeri 4 Balikpapan**", "NPSN 30401488, negeri, akreditasi A"], "Selatan", "Jl. Sepinggan Baru III No. 36", ["0542-780044", "sman4bpn@gmail.com"], "47", "1.186", "2"),
  row(["**SMA Negeri 7 Balikpapan**", "NPSN 30401504, negeri, akreditasi A"], "Timur", "Jl. Mulawarman", ["0542-790025", "sma7balikpapan@gmail.com"], "53", "1.089", "2"),
  row(["**SMA Negeri 8 Balikpapan**", "NPSN 30402920, negeri, akreditasi A"], "Barat", "Jl. AMD Gn. IV RT 14", ["0542-7584515", "info@sman8bpp.sch.id"], "45", "869", "2"),
  row(["**SMK Negeri 4 Balikpapan**", "NPSN 30401508, negeri, akreditasi A"], "Selatan", "Jl. Belibis 3 No. 01", ["0542-873890", "school@smkn4balikpapan.sch.id"], "67", "1.260", "2"),
  row(["**SMK Negeri 5 Balikpapan**", "NPSN 30404599, negeri, akreditasi B"], "Timur", "Jl. Mulawarman No. 65", ["0542-790767", "smkn5bpn@gmail.com"], "70", "1.229", "2"),
  row(["**SMK Negeri 6 Balikpapan**", "NPSN 30409918, negeri, akreditasi B"], "Utara", "Jl. Soekarno Hatta KM 7 RT 61, Graha Indah", ["0542-8830033", "sekolah@smkn6-bpn.sch.id"], "75", "1.414", "2"),
  sub("Jumlah kategori C", "465", "9.468", "16"),
  new TableRow({ cantSplit: true, children: [
    cell("", widths[0], { fill: "DCE6F2" }), cell("Total peserta", widths[1] + widths[2] + widths[3] + widths[4], { span: 4, fill: "DCE6F2", bold: true }),
    cell("", widths[5], { fill: "DCE6F2" }), cell("", widths[6], { fill: "DCE6F2" }), cell("80", widths[7], { fill: "DCE6F2", bold: true, right: true }),
  ] }),
];
const tbl = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows });

const children = [
  h2("2.3 Peserta dan Cara Rekrutmen"),
  p("Peserta berjumlah 80 guru dari 4 SLB dan sekolah inklusif di Balikpapan, dengan rincian sekolah, kontak, dan jumlah guru yang diundang pada tabel berikut [16, 21]."),
  tbl,
  note("Guru diundang adalah jumlah guru yang kami undang dari tiap sekolah dan sekaligus menjadi kuota peserta. Jumlah itu usulan Sanggabiz: sekitar separuh guru tiap SLB, empat guru per kecamatan untuk SD dan SMP, dan dua guru per sekolah untuk SMA dan SMK."),
  note("* Data sumber mencantumkan nomor 8870057 tanpa kode area. Kami menambahkan kode area Balikpapan (0542). Nomor telepon SMA Negeri 7 dan SMA Negeri 8 juga ditambah kode area."),
  note("Sumber data sekolah, kontak, guru, dan siswa: Referensi Kemendikdasmen [16] dan Sekolah Kita, semester 2026/2027 Gasal [21], diakses 15 Oktober 2026. Daftar SMA dan SMK inklusif: Juknis SPMB 2025 [17]. Status inklusif kategori B: pemberitaan 2019, penelitian 2023, dan sosialisasi 2024 [23, 26, 27, 28]. Rujukan [16] sampai [28] ada di dokumen v4."),
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
  creator: "Sanggabiz", title: "Bab 2.3 versi ringkas: satu tabel sekolah sasaran",
  styles: {
    default: { document: { run: { font: FONT, size: 19, color: INK } } },
    paragraphStyles: [{ id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: FONT, size: 22, bold: true, color: NAVY }, paragraph: { spacing: { before: 0, after: 110 }, outlineLevel: 1 } }],
  },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } }, footers: { default: footer }, children }],
});
const outDir = path.join(__dirname, "..", "bagian-sasaran");
fs.mkdirSync(outDir, { recursive: true });
const out = path.join(outDir, "Bab 2.3 versi ringkas - satu tabel sekolah sasaran.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ditulis:", out, "baris data:", no); });
