// Membuat daftar perubahan v2 ke v3. Jalankan: python3 data_daftar_perubahan.py && NODE_PATH=$(npm root -g) node build_daftar_perubahan.js
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Footer, AlignmentType, BorderStyle, WidthType, ShadingType, PageNumber, PageOrientation, VerticalAlign } = require("docx");
const NAVY = "0B396D", GOLD = "FFC00C", INK = "1D1D1F", MUTED = "5C5F66", LINE = "C9D2DE", GROUP = "DCE6F2", ZEBRA = "F7F9FC";
const FONT = "Montserrat";
const data = JSON.parse(fs.readFileSync(path.join(__dirname, "daftar_perubahan.json"), "utf8"));
const COLS = [560, 1700, 4600, 4600, 2878, 800];
const W = COLS.reduce((a, b) => a + b, 0);
const thin = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: thin, bottom: thin, left: thin, right: thin };
const para = (t, o = {}) => new Paragraph({ spacing: { after: o.after ?? 50, line: 260 }, alignment: o.align, children: [new TextRun({ text: t, font: FONT, size: o.size || 16, bold: o.bold, color: o.color || INK })] });
const cell = (lines, w, o = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA }, borders, columnSpan: o.span,
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  verticalAlign: o.valign || VerticalAlign.TOP, margins: { top: 70, bottom: 50, left: 90, right: 90 },
  children: (Array.isArray(lines) ? lines : [lines]).map((t) => para(t, o)),
});
const head = new TableRow({ tableHeader: true, cantSplit: true, children: ["No", "Bagian", "Di v2 Anda", "Di v3", "Alasan", "Pakai?"].map((t, i) => cell(t, COLS[i], { fill: NAVY, bold: true, color: "FFFFFF", valign: VerticalAlign.CENTER, align: i === 0 || i === 5 ? AlignmentType.CENTER : undefined })) });
const rows = [head];
let n = 0;
data.forEach((g) => {
  rows.push(new TableRow({ cantSplit: true, children: [cell(g.title, W, { span: 6, fill: GROUP, bold: true, size: 17, color: NAVY })] }));
  g.rows.forEach((r, i) => {
    n++;
    const fill = i % 2 ? ZEBRA : undefined;
    rows.push(new TableRow({ cantSplit: true, children: [
      cell(String(n), COLS[0], { fill, align: AlignmentType.CENTER }),
      cell(r[0], COLS[1], { fill, bold: true }),
      cell(r[1], COLS[2], { fill }),
      cell(r[2], COLS[3], { fill }),
      cell(r[3] || "", COLS[4], { fill, color: MUTED }),
      cell("[  ]", COLS[5], { fill, align: AlignmentType.CENTER }),
    ] }));
  });
});
const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 18 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 850, bottom: 850, left: 850, right: 850 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Daftar perubahan v2 ke v3  |  Halaman ", font: FONT, size: 14, color: MUTED }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 14, color: MUTED })] })] }) },
    children: [
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Daftar Perubahan v2 ke v3", font: FONT, bold: true, size: 36, color: NAVY })] }),
      new Paragraph({ spacing: { after: 160 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: GOLD, space: 6 } }, children: [new TextRun({ text: "Workshop Pembelajaran Berdiferensiasi, SLB dan Inklusif Balikpapan", font: FONT, size: 20, color: MUTED })] }),
      para("Kolom Di v2 Anda diambil dari file v2 yang Anda unggah. Kolom Di v3 diambil dari v3 di repo. Teks di dalam sel ditulis apa adanya, bisa langsung disalin. Tanda tiga titik berarti bagian kalimat yang tidak berubah. Kolom Pakai diisi tim.", { size: 17, after: 140 }),
      new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: COLS, rows }),
    ],
  }],
});
Packer.toBuffer(doc).then((b) => {
  const out = path.join(__dirname, "..", "proposal", "Daftar Perubahan v2 ke v3 (untuk dibandingkan).docx");
  fs.writeFileSync(out, b); console.log("ditulis:", out);
});
