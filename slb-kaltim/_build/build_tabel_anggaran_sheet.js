// Membuat proposal docx. Jalankan: NODE_PATH=$(npm root -g) node build_proposal.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, Footer,
  AlignmentType, BorderStyle, WidthType, ShadingType, HeightRule, PageNumber, LevelFormat,
  VerticalAlign, TabStopType,
} = require("docx");

const NAVY = "0B396D", GOLD = "FFC00C", INK = "1D1D1F", MUTED = "5C5F66", LIGHT = "EEF3F9", LINE = "C9D2DE";
const FONT = "Montserrat";
const W = 9638; // lebar isi A4 dengan margin 2 cm
const A = (f) => path.join(__dirname, "assets", f);

// ---------- penolong teks ----------
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
  spacing: { after: o.after ?? 110, before: o.before ?? 0, line: o.line ?? 290 },
  keepNext: o.keepNext, indent: o.indent, pageBreakBefore: o.pageBreakBefore,
});
const h1 = (text, pb = false) => new Paragraph({
  heading: "Heading1", children: [new TextRun({ text })], pageBreakBefore: pb, keepNext: true,
});
const h2 = (text) => new Paragraph({ heading: "Heading2", children: [new TextRun({ text })], keepNext: true });
const bullet = (text) => new Paragraph({
  numbering: { reference: "bul", level: 0 }, children: runs(text), spacing: { after: 70, line: 290 },
});
const caption = (text) => new Paragraph({
  alignment: AlignmentType.CENTER, children: [new TextRun({ text, size: 17, color: INK })], spacing: { after: 180 },
});
const h3 = (text) => new Paragraph({
  children: [new TextRun({ text, bold: true, size: 20, color: NAVY })], spacing: { before: 200, after: 90 }, keepNext: true,
});
const fmt = (n) => "Rp " + Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");

// ---------- penolong tabel ----------
const thin = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: thin, bottom: thin, left: thin, right: thin };
function cell(content, width, o = {}) {
  const lines = Array.isArray(content) ? content : [content];
  return new TableCell({
    width: { size: width, type: WidthType.DXA }, borders,
    columnSpan: o.span,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    verticalAlign: o.valign || VerticalAlign.TOP,
    margins: { top: 70, bottom: 70, left: 100, right: 100 },
    children: lines.map((t) => new Paragraph({
      alignment: o.align || AlignmentType.LEFT, spacing: { after: 30, line: 260 },
      children: runs(t, { size: o.size || 17, bold: o.bold, color: o.color || INK }),
    })),
  });
}
function table(widths, header, rows, o = {}) {
  const head = new TableRow({
    tableHeader: true, cantSplit: true,
    children: header.map((t, i) => cell(t, widths[i], { fill: NAVY, color: "FFFFFF", bold: true, size: 16 })),
  });
  const body = rows.map((r, ri) => new TableRow({
    cantSplit: !o.split,
    children: r.map((c, i) => {
      if (c && c.group) return cell(c.group, widths.reduce((a, b) => a + b, 0), { fill: "DCE6F2", bold: true, span: widths.length });
      const al = o.align && o.align[i];
      return cell(c && c.text !== undefined ? c.text : c, widths[i], {
        fill: c && typeof c.bg === "string" ? c.bg : (ri % 2 ? "F7F9FC" : undefined), align: al, size: o.size, bold: !!(c && typeof c === "object" && !Array.isArray(c) && c.bold === true),
      });
    }).slice(0, r[0] && r[0].group ? 1 : r.length),
  }));
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows: [head, ...body] });
}
const spacer = (n = 120) => new Paragraph({ spacing: { after: n }, children: [] });
const img = (file, w, h, type = "png") => new ImageRun({ type, data: fs.readFileSync(A(file)), transformation: { width: w, height: h } });
const centerImg = (file, w, h, type) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, keepNext: true, children: [img(file, w, h, type)] });


// ---------- tabel anggaran versi sheet (isi sesuai sheet, gaya sama dengan proposal v2) ----------
const W8 = [760, 2840, 500, 800, 500, 900, 1500, 1838];
const SUMW = 2840 + 500 + 800 + 500 + 900 + 1500; // lebar label ringkasan (kolom 2 sampai 7)
const GROUP_W = W8.slice(1).reduce((a, b) => a + b, 0);
const BANYAK = [
  { kode: "A", judul: "Program Implementation Personnel", baris: [
    ["A.1", "Project Manager", "1", "Person", "1", "days", 4000000],
    ["A.2", "Pemateri lokal, pakar pendidikan inklusif (Sesi 1 sampai 4), 4 jam tatap muka", "1", "Person", "1", "days", 10000000],
    ["A.3", "MC dan Moderator", "1", "Person", "1", "days", 1000000],
  ] },
  { kode: "B", judul: "Materi dan Equipment", baris: [
    ["B.1", "Modul/Materi", "1", "days", "1", "package", 3000000],
    ["B.2", "Perlengkapan peserta: ID card, pulpen, kartu kerja, E-Sertifikat", "1", "days", "80", "orang", 50000],
    ["B.3", "Snack", "1", "days", "85", "orang", 30000],
    ["B.4", "Makan Siang", "1", "days", "85", "orang", 50000],
  ] },
  { kode: "C", judul: "Local Transport", baris: [
    ["C.1", "Rent Cars", "1", "days", "1", "cars", 1500000],
  ] },
  { kode: "D", judul: "Evaluation dan Report", baris: [
    ["D.1", "Monitoring dan Reporting", "1", "pax", "1", "time", 2000000],
  ] },
];
const dot = (n) => Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
const C = (t, w, o = {}) => cell(t, w, { size: 16, ...o });
const rows = [];
rows.push(new TableRow({ tableHeader: true, cantSplit: true, children: [
  C("Code", W8[0], { fill: NAVY, color: "FFFFFF", bold: true, align: AlignmentType.CENTER }),
  C("Description", W8[1], { fill: NAVY, color: "FFFFFF", bold: true }),
  C("Quantity", W8[2] + W8[3], { fill: NAVY, color: "FFFFFF", bold: true, span: 2, align: AlignmentType.CENTER }),
  C("Frequency", W8[4] + W8[5], { fill: NAVY, color: "FFFFFF", bold: true, span: 2, align: AlignmentType.CENTER }),
  C("Price (IDR)", W8[6], { fill: NAVY, color: "FFFFFF", bold: true, align: AlignmentType.RIGHT }),
  C("Total (IDR)", W8[7], { fill: NAVY, color: "FFFFFF", bold: true, align: AlignmentType.RIGHT }),
] }));
const summaryRow = (label, val, fill, bold) => new TableRow({ cantSplit: true, children: [
  C("", W8[0], { fill }),
  C(label, SUMW, { fill, bold, span: 6, align: AlignmentType.RIGHT }),
  C(val, W8[7], { fill, bold, align: AlignmentType.RIGHT }),
] });
let total = 0;
BANYAK.forEach((g) => {
  rows.push(new TableRow({ cantSplit: true, children: [
    C(g.kode, W8[0], { fill: "DCE6F2", bold: true, align: AlignmentType.CENTER }),
    C(g.judul, GROUP_W, { fill: "DCE6F2", bold: true, span: 7 }),
  ] }));
  let sub = 0;
  g.baris.forEach((r, i) => {
    const jumlah = Number(r[2]) * Number(r[4]) * r[6]; sub += jumlah;
    const fill = i % 2 ? "F7F9FC" : undefined;
    rows.push(new TableRow({ cantSplit: true, children: [
      C(r[0], W8[0], { fill, align: AlignmentType.CENTER }),
      C(r[1], W8[1], { fill }),
      C(r[2], W8[2], { fill, align: AlignmentType.CENTER }),
      C(r[3], W8[3], { fill }),
      C(r[4], W8[4], { fill, align: AlignmentType.CENTER }),
      C(r[5], W8[5], { fill }),
      C(dot(r[6]), W8[6], { fill, align: AlignmentType.RIGHT }),
      C(dot(jumlah), W8[7], { fill, align: AlignmentType.RIGHT }),
    ] }));
  });
  rows.push(summaryRow(g.judul, dot(sub), "F2F5FA", true));
  total += sub;
});
const fee = total * 0.10;
const pph = (total + fee) * 0.005;
const grand = total + pph; // sesuai sheet: total + PPh (fee tidak ikut dijumlahkan)
rows.push(summaryRow("Total", dot(total), "F2F5FA", true));
rows.push(summaryRow("Management Fee 10%", dot(fee), "F2F5FA", false));
rows.push(summaryRow("PPH 0,5%", dot(pph), "F2F5FA", false));
rows.push(summaryRow("Total (include Pajak)", dot(grand), "DCE6F2", true));
const tabel = new Table({ width: { size: 9638, type: WidthType.DXA }, columnWidths: W8, rows });

const page = { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } };
const doc = new Document({
  creator: "Sanggabiz", title: "Rincian Anggaran",
  styles: {
    default: { document: { run: { font: FONT, size: 19, color: INK } } },
    paragraphStyles: [
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 22, bold: true, color: NAVY },
        paragraph: { spacing: { before: 0, after: 110 }, outlineLevel: 1 } },
    ],
  },
  sections: [{ properties: { page }, children: [
    h2("5.2 Rincian Anggaran"),
    tabel,
    new Paragraph({ spacing: { before: 160, after: 40 }, children: [new TextRun({ text: "Catatan *)", bold: true, size: 17 })] }),
    p("Anggaran di atas sudah termasuk pajak.", { after: 20, run: { size: 17 } }),
    p("Anggaran bersifat estimatif dan masih dapat dinegosiasikan sesuai kesepakatan bersama.", { run: { size: 17 } }),
  ] }],
});
const out = path.join(__dirname, "..", "proposal", "Tabel Anggaran 5.2 versi sheet (untuk disalin).docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ditulis:", out, "total", dot(total), "fee", dot(fee), "pph", dot(pph), "grand", dot(grand), "grand dengan fee", dot(total + fee + pph)); });
