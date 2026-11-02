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


// ---------- tabel anggaran untuk disalin (harga ke klien, sudah termasuk pajak) ----------
const ANGGARAN = [
  { group: "A. Pemateri dan pemandu acara" },
  ["1", "Pemateri lokal, pakar pendidikan inklusif (Sesi 1 sampai 4), 4 jam tatap muka", 4, "jam", 1500000],
  ["2", "MC dan moderator", 1, "orang", 750000],
  { group: "B. Tempat dan konsumsi" },
  ["3", "Paket meeting setengah hari di Quest Hotel Balikpapan, Orchid Room (ruang, proyektor dan tata suara, coffee break, makan siang)", 90, "orang", 148500],
  { group: "C. Materi dan bahan ajar" },
  ["4", "Penyusunan modul ringkas dan template modul ajar berdiferensiasi", 1, "paket", 6000000],
  ["5", "Bahan ajar interaktif: kuis dan lembar kerja digital", 1, "paket", 5500000],
  ["6", "Cetak modul ringkas dan template (hitam putih, sampul warna)", 90, "eksemplar", 15000],
  ["7", "Perlengkapan peserta: ID card, pulpen, kartu kerja", 90, "set", 15000],
  ["8", "Sertifikat peserta", 80, "lembar", 5000],
  { group: "D. Rekrutmen, pendampingan, dan koordinasi" },
  ["9", "Rekrutmen dan koordinasi peserta: undangan, formulir pendaftaran, konfirmasi, pengingat", 1, "paket", 7500000],
  ["10", "Pendamping kelompok dan panitia lokal", 2, "orang", 500000],
  ["11", "IO (Information Officer) dan koordinator kehadiran, H-7 sampai hari-H", 1, "orang", 750000],
  { group: "E. Evaluasi, dokumentasi, dan tindak lanjut" },
  ["12", "Pre-test, post-test, penilaian produk, dan laporan akhir", 1, "paket", 5000000],
  ["13", "Dokumentasi foto", 1, "paket", 800000],
  ["14", "Tindak lanjut empat minggu lewat grup WhatsApp", 1, "paket", 3000000],
  { group: "F. Lain-lain" },
  ["15", "Transportasi lokal dan kebutuhan hari-H", 1, "paket", 500000],
];
let TOTAL = 0;
const anggaranRows = ANGGARAN.map((r) => {
  if (r.group) return [r];
  const jumlah = r[2] * r[4]; TOTAL += jumlah;
  return [r[0], r[1], String(r[2]), r[3], fmt(r[4]), fmt(jumlah)];
});
const PER_PESERTA = Math.round(TOTAL / 80 / 1000) * 1000;
const widths = [500, 3900, 700, 1100, 1700, 1738];
const sum = (label, val) => [{ text: "", bg: "F2F5FA" }, { text: label, bold: true, bg: "F2F5FA" }, { text: "", bg: "F2F5FA" }, { text: "", bg: "F2F5FA" }, { text: "", bg: "F2F5FA" }, { text: val, bold: true, bg: "F2F5FA" }];
const tabel = table(widths, ["No", "Komponen", "Vol", "Satuan", "Harga satuan", "Jumlah"], [
  ...anggaranRows,
  sum("TOTAL (sudah termasuk pajak)", fmt(TOTAL)),
  sum("Per peserta (80 orang, dibulatkan)", fmt(PER_PESERTA)),
], { align: [AlignmentType.CENTER, AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.RIGHT, AlignmentType.RIGHT] });

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
  sections: [{ properties: { page }, children: [h2("5.2 Rincian Anggaran"), tabel] }],
});
const out = path.join(__dirname, "..", "proposal", "Tabel Anggaran 5.2 (untuk disalin).docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ditulis:", out, fmt(TOTAL), fmt(PER_PESERTA)); });
