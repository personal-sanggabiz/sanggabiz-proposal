// v3: hanya sekolah yang datanya lengkap dan statusnya jelas (4 SLB, SMP Negeri 5, 8 SMA dan SMK inklusif), target total 80.
// Versi lama: build_daftar_sekolah_v1.js (12 kolom), _v2.js (19 sekolah, termasuk yang belum terkonfirmasi). Data dari ../10-RISET-SEKOLAH-SASARAN-DAN-KONTAK.md dan build_bagian_sasaran.js (v4).
// Jalankan: NODE_PATH=/opt/node-tools/node_modules node build_daftar_sekolah.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Footer,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign, TabStopType,
} = require("docx");

const NAVY = "0B396D", INK = "1D1D1F", MUTED = "5C5F66", GOLD = "FFC00C", LINE = "C9D2DE";
const FONT = "Montserrat";
const W = 9638;

// [nama, npsn, jenjang, status dan akreditasi, kecamatan, guru tercatat (null bila belum ada), target peserta]
const GROUPS = [
  { label: "A. SLB. Target 40 peserta, sekitar separuh guru tiap SLB", rows: [
    ["SLB Negeri Balikpapan", "30401525", "SLB, TKLB sampai SMLB", "Negeri, A", "Balikpapan Selatan", 44, 22],
    ["SLB Tunas Bangsa", "30402883", "SLB, SDLB sampai SMLB", "Swasta, B", "Balikpapan Selatan", 13, 7],
    ["SLB Dharma Kencana Balikpapan", "30402882", "SLB, SDLB sampai SMLB", "Swasta, B", "Balikpapan Kota", 10, 5],
    ["Sekolah Khusus Santo Fransiskus Assisi", "69965143", "SLB, TKLB sampai SMLB", "Swasta, B", "Balikpapan Kota", 13, 6],
  ] },
  { label: "B. SMP inklusif. Target 8 peserta, dua tim guru", rows: [
    ["SMP Negeri 5 Balikpapan", "30401474", "SMP inklusif", "Negeri, A", "Balikpapan Selatan", 49, 8],
  ] },
  { label: "C. SMA dan SMK inklusif. Target 32 peserta, 4 per sekolah", rows: [
    ["SMA Negeri 1 Balikpapan", "30401491", "SMA inklusif", "Negeri, A", "Balikpapan Kota", 57, 4],
    ["SMA Negeri 2 Balikpapan", "30401490", "SMA inklusif", "Negeri, A", "Balikpapan Utara", 51, 4],
    ["SMA Negeri 4 Balikpapan", "30401488", "SMA inklusif", "Negeri, A", "Balikpapan Selatan", 47, 4],
    ["SMA Negeri 7 Balikpapan", "30401504", "SMA inklusif", "Negeri, A", "Balikpapan Timur", 53, 4],
    ["SMA Negeri 8 Balikpapan", "30402920", "SMA inklusif", "Negeri, A", "Balikpapan Barat", 45, 4],
    ["SMK Negeri 4 Balikpapan", "30401508", "SMK inklusif", "Negeri, A", "Balikpapan Selatan", 67, 4],
    ["SMK Negeri 5 Balikpapan", "30404599", "SMK inklusif", "Negeri, B", "Balikpapan Timur", 70, 4],
    ["SMK Negeri 6 Balikpapan", "30409918", "SMK inklusif", "Negeri, B", "Balikpapan Utara", 75, 4],
  ] },
];

const sumT = (rows) => rows.reduce((t, r) => t + r[6], 0);
const all = GROUPS.flatMap((g) => g.rows);
const TOTAL = sumT(all);
if (TOTAL !== 80) throw new Error("total target " + TOTAL + " bukan 80");
[40, 8, 32].forEach((n, i) => { if (sumT(GROUPS[i].rows) !== n) throw new Error("subtotal kelompok " + i + " salah"); });

// ---------- penolong ----------
const thin = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: thin, bottom: thin, left: thin, right: thin };
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
function cell(content, width, o = {}) {
  const lines = Array.isArray(content) ? content : [content];
  return new TableCell({
    width: { size: width, type: WidthType.DXA }, borders, columnSpan: o.span,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    verticalAlign: VerticalAlign.CENTER, margins: { top: 70, bottom: 70, left: 100, right: 100 },
    children: lines.map((t) => new Paragraph({
      alignment: o.align || AlignmentType.LEFT, spacing: { after: 20, line: 250 },
      children: runs(t, { size: o.size || 17, bold: o.bold, color: o.color || INK }),
    })),
  });
}

const widths = [450, 3150, 1450, 1000, 1500, 1038, 1050];
if (widths.reduce((a, b) => a + b, 0) !== W) throw new Error("lebar kolom salah");
const header = ["No", "Nama sekolah", "Jenjang", "Status, akreditasi", "Kecamatan", "Guru tercatat", "Target peserta"];
const RIGHT = [5, 6];

const headRow = new TableRow({
  tableHeader: true, cantSplit: true,
  children: header.map((t, i) => cell(t, widths[i], { fill: NAVY, color: "FFFFFF", bold: true, size: 16, align: RIGHT.includes(i) ? AlignmentType.RIGHT : undefined })),
});
const bodyRows = [];
let no = 0;
for (const g of GROUPS) {
  bodyRows.push(new TableRow({ cantSplit: true, children: [cell(g.label, W, { span: widths.length, fill: "DCE6F2", bold: true, color: NAVY, size: 17 })] }));
  g.rows.forEach((r, ri) => {
    no += 1;
    const [nama, npsn, jenjang, st, kec, guru, target] = r;
    const fill = ri % 2 ? "F7F9FC" : undefined;
    const namaCell = npsn === "-" ? nama : [`**${nama}**`, `NPSN ${npsn}`];
    const vals = [String(no), namaCell, jenjang, st, kec, guru === null ? "-" : String(guru), `**${target}**`];
    bodyRows.push(new TableRow({
      cantSplit: true,
      children: vals.map((v, i) => cell(v, widths[i], { fill, align: RIGHT.includes(i) ? AlignmentType.RIGHT : undefined })),
    }));
  });
}
const guruTotal = all.reduce((t, r) => t + (r[5] || 0), 0);
bodyRows.push(new TableRow({
  cantSplit: true,
  children: ["", `**Jumlah ${all.length} sekolah**`, "", "", "", `**${guruTotal}**`, `**${TOTAL}**`]
    .map((v, i) => cell(v, widths[i], { fill: "DCE6F2", align: RIGHT.includes(i) ? AlignmentType.RIGHT : undefined })),
}));
const tabel = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows: [headRow, ...bodyRows] });

const p = (text, o = {}) => new Paragraph({ children: runs(text, {}), spacing: { after: o.after ?? 100, line: 280 } });
const note = (text) => new Paragraph({ children: runs(text, { size: 16, color: MUTED }), spacing: { before: 40, after: 80, line: 255 } });

const children = [
  new Paragraph({ heading: "Heading1", children: [new TextRun({ text: "Sekolah Sasaran dan Target Peserta di Kota Balikpapan" })] }),
  p(`Kami mengundang ${TOTAL} guru dari ${all.length} sekolah: 4 SLB, 1 SMP inklusif, dan 8 SMA dan SMK inklusif. Semua sekolah di tabel punya NPSN, status, akreditasi, kecamatan, dan jumlah guru yang tercatat, serta dasar penetapan atau sumber yang jelas.`, { after: 140 }),
  tabel,
  note("Target peserta adalah usulan Sanggabiz dan dapat disesuaikan dengan permintaan klien. Sekolah menunjuk sendiri gurunya."),
  note("Aturan target. SLB: sekitar separuh guru tiap sekolah, supaya kegiatan belajar tetap berjalan. SMA dan SMK inklusif: 4 guru per sekolah, yaitu guru mata pelajaran dan guru pendamping khusus atau guru bimbingan konseling. SMP Negeri 5: 8 guru, dua tim, karena ini satu-satunya SMP yang disebut konsisten oleh pejabat dinas dan penelitian (2019 sampai 2023). Jumlahnya dapat dikurangi bila dinas menambah SD dan SMP inklusif."),
  note("Sekolah yang tidak dimasukkan. SD dan SMP inklusif lain belum masuk karena Keputusan Wali Kota yang menetapkannya belum terbit di sumber publik, dan sumber berita saling berbeda soal namanya. [ISI TIM: daftar SD dan SMP inklusif dari Disdikbud Kota Balikpapan, lalu tambahkan ke tabel dan kurangi target SMA, SMK, atau SMP Negeri 5]"),
  note("Sumber: guru tercatat dari portal Sekolah Kita, Kemendikdasmen, semester 2026/2027 Gasal. NPSN dan akreditasi dari Referensi Kemendikdasmen. SMA dan SMK inklusif dari Lampiran 9 Juknis SPMB SMA, SMK, dan SLB Kota Balikpapan 2025. Alamat dan kontak tiap sekolah ada di v1 daftar sekolah."),
];

const footer = new Footer({
  children: [new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: LINE, space: 6 } },
    tabStops: [{ type: TabStopType.RIGHT, position: W }],
    children: [
      new TextRun({ text: "Sanggabiz  |  Sekolah Sasaran dan Target Peserta Balikpapan", size: 14, color: MUTED }),
      new TextRun({ text: "\t", size: 14 }),
      new TextRun({ children: ["Halaman ", PageNumber.CURRENT], size: 14, color: MUTED }),
    ],
  })],
});

const doc = new Document({
  creator: "Sanggabiz", title: "Sekolah Sasaran dan Target Peserta di Kota Balikpapan",
  styles: {
    default: { document: { run: { font: FONT, size: 19, color: INK } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: NAVY },
        paragraph: { spacing: { before: 0, after: 200 }, outlineLevel: 0, border: { bottom: { style: BorderStyle.SINGLE, size: 14, color: GOLD, space: 6 } } } },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: footer }, children,
  }],
});

const outDir = path.join(__dirname, "..", "daftar-sekolah");
fs.mkdirSync(outDir, { recursive: true });
const out = path.join(outDir, "v3 Sekolah Sasaran dan Target Peserta SLB dan Inklusif Balikpapan.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ditulis:", out, "| sekolah:", all.length, "guru:", guruTotal, "target:", TOTAL); });
