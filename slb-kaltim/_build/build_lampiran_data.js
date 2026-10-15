// Membuat "Lampiran D. Data SD dan SMP di Kota Balikpapan" (docx) dari ../data/sd-smp-balikpapan.psv.
// Jalankan: NODE_PATH=<folder node_modules berisi docx> node build_lampiran_data.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Footer,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign, TabStopType,
} = require("docx");

const NAVY = "0B396D", INK = "1D1D1F", MUTED = "5C5F66", GOLD = "FFC00C", LINE = "C9D2DE";
const FONT = "Montserrat";
const W = 9638;
const rp = (n) => n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
const sum = (a, f) => a.reduce((t, x) => t + f(x), 0);

const rows = fs.readFileSync(path.join(__dirname, "..", "data", "sd-smp-balikpapan.psv"), "utf8")
  .trim().split("\n").slice(1).map((l) => l.split("|"))
  .map(([npsn, nama, bp, st, kec, guru, siswa]) => ({ npsn, nama, bp, st: st === "N" ? "Negeri" : "Swasta", kec, guru: +guru, siswa: +siswa }));
const KEC = ["Selatan", "Timur", "Utara", "Tengah", "Barat", "Kota"];

const thin = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: thin, bottom: thin, left: thin, right: thin };
const cell = (t, w, o = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.TOP,
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  margins: { top: 40, bottom: 40, left: 80, right: 80 },
  children: [new Paragraph({ alignment: o.right ? AlignmentType.RIGHT : AlignmentType.LEFT, spacing: { after: 0, line: 240 }, children: [new TextRun({ text: String(t), size: 15, bold: o.bold, color: o.color || INK })] })],
});
const widths = [450, 1150, 4200, 700, 900, 800, 1438];
const right = [5, 6];
function kecTable(k) {
  const r = rows.filter((x) => x.kec === k).sort((a, b) => a.bp.localeCompare(b.bp) || a.st.localeCompare(b.st) || a.nama.localeCompare(b.nama));
  const head = new TableRow({ tableHeader: true, cantSplit: true, children: ["No", "NPSN", "Nama sekolah", "Jenjang", "Status", "Guru", "Siswa"].map((t, i) => cell(t, widths[i], { fill: NAVY, color: "FFFFFF", bold: true, right: right.includes(i) })) });
  const body = r.map((x, i) => new TableRow({ cantSplit: true, children: [i + 1, x.npsn, x.nama, x.bp, x.st, x.guru, rp(x.siswa)].map((t, j) => cell(t, widths[j], { fill: i % 2 ? "F7F9FC" : undefined, right: right.includes(j) })) }));
  const tot = new TableRow({ cantSplit: true, children: ["", "", `Jumlah Balikpapan ${k}: ${r.filter((x) => x.bp === "SD").length} SD dan ${r.filter((x) => x.bp === "SMP").length} SMP`, "", "", rp(sum(r, (x) => x.guru)), rp(sum(r, (x) => x.siswa))].map((t, j) => cell(t, widths[j], { fill: "DCE6F2", bold: true, right: right.includes(j) })) });
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows: [head, ...body, tot] });
}

const children = [
  new Paragraph({ heading: "Heading2", children: [new TextRun({ text: "Lampiran D. Data SD dan SMP di Kota Balikpapan" })], keepNext: true }),
  new Paragraph({ spacing: { after: 110, line: 290 }, children: [new TextRun({ text: `Lampiran ini memuat ${rows.length} SD dan SMP negeri dan swasta di enam kecamatan, dengan ${rp(sum(rows, (x) => x.guru))} guru dan ${rp(sum(rows, (x) => x.siswa))} siswa. Data diambil dari portal Sekolah Kita, Kemendikdasmen, semester 2026/2027 Gasal, diakses 15 Oktober 2026 [21]. Madrasah dan sekolah kerja sama tidak dimasukkan. Satuan pendidikan diurutkan menurut kecamatan, jenjang, status, lalu nama.`, size: 19 })] }),
  new Paragraph({ spacing: { after: 160, line: 260 }, children: [new TextRun({ text: "Angka nol atau jauh di bawah sekolah sejenis pada kolom guru atau siswa menunjukkan data yang belum lengkap di sumber, bukan sekolah tanpa guru atau siswa.", size: 16, color: MUTED })] }),
];
for (const k of KEC) {
  children.push(new Paragraph({ children: [new TextRun({ text: `Kecamatan Balikpapan ${k}`, bold: true, size: 20, color: NAVY })], spacing: { before: 220, after: 90 }, keepNext: true }));
  children.push(kecTable(k));
}
children.push(new Paragraph({ spacing: { before: 200, after: 70, line: 260 }, children: [new TextRun({ text: "[21] Kemendikdasmen. Sekolah Kita, profil satuan pendidikan Kota Balikpapan, semester 2026/2027 Gasal. https://sekolah.data.kemendikdasmen.go.id. Diakses 15 Oktober 2026.", size: 16 })] }));

const footer = new Footer({
  children: [new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: LINE, space: 6 } },
    tabStops: [{ type: TabStopType.RIGHT, position: W }],
    children: [
      new TextRun({ text: "Sanggabiz  |  Lampiran D. Data SD dan SMP di Kota Balikpapan", size: 14, color: MUTED }),
      new TextRun({ text: "\t", size: 14 }),
      new TextRun({ children: ["Halaman ", PageNumber.CURRENT], size: 14, color: MUTED }),
    ],
  })],
});
const doc = new Document({
  creator: "Sanggabiz", title: "Lampiran D. Data SD dan SMP di Kota Balikpapan",
  styles: {
    default: { document: { run: { font: FONT, size: 19, color: INK } } },
    paragraphStyles: [{ id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: FONT, size: 26, bold: true, color: NAVY },
      paragraph: { spacing: { before: 0, after: 160 }, outlineLevel: 1, border: { bottom: { style: BorderStyle.SINGLE, size: 14, color: GOLD, space: 6 } } } }],
  },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } }, footers: { default: footer }, children }],
});
const outDir = path.join(__dirname, "..", "bagian-sasaran");
fs.mkdirSync(outDir, { recursive: true });
const out = path.join(outDir, "v4 Lampiran D - Data SD dan SMP Balikpapan.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ditulis:", out); });
