// v1: daftar sekolah SLB dan sekolah inklusif Kota Balikpapan dalam satu tabel (docx, landscape A4).
// Data diambil dari ../10-RISET-SEKOLAH-SASARAN-DAN-KONTAK.md dan tabel di build_bagian_sasaran.js (v4).
// Jalankan: NODE_PATH=/opt/node-tools/node_modules node build_daftar_sekolah.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Footer,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign, TabStopType, PageOrientation,
} = require("docx");

const NAVY = "0B396D", INK = "1D1D1F", MUTED = "5C5F66", GOLD = "FFC00C", LINE = "C9D2DE";
const FONT = "Montserrat";
const MARGIN = 700;
const W = 16838 - 2 * MARGIN; // 15438

// ---------- data ----------
// [nama, [jenis, jenjang], npsn, [status, akreditasi], kecamatan, alamat, telepon, surel, guru, siswa, dasar]
const SLB = [
  ["SLB Negeri Balikpapan", ["SLB", "TKLB sampai SMLB"], "30401525", ["Negeri", "A"], "Balikpapan Selatan", "Jl. Letkol Pol. HM. Asnawi Arbain RT 14, Kel. Sepinggan Raya", "0542-8879984", "slbn.balikpapan@yahoo.com", 44, 381, "Referensi Kemendikdasmen"],
  ["SLB Tunas Bangsa", ["SLB", "SDLB sampai SMLB"], "30402883", ["Swasta", "B"], "Balikpapan Selatan", "Jl. Kutilang RSS Damai III RT 28, Kel. Gunung Bahagia", "0542-8870057 *", "tunasbangsa831@yahoo.co.id", 13, 85, "Referensi Kemendikdasmen"],
  ["SLB Dharma Kencana Balikpapan", ["SLB", "SDLB sampai SMLB"], "30402882", ["Swasta", "B"], "Balikpapan Kota", "Jl. Pembangunan No. 30 Blok A, Kel. Telaga Sari", "0542-410406", "slbdk@yahoo.co.id", 10, 111, "Referensi Kemendikdasmen"],
  ["Sekolah Khusus Santo Fransiskus Assisi", ["SLB", "TKLB sampai SMLB"], "69965143", ["Swasta", "B"], "Balikpapan Kota", "Jl. Mayor (Pol) Zaenal Arifin No. 40 RT 18, Kel. Damai", "0542-8794310", "skhfransiskusassisi@gmail.com", 13, 105, "Referensi Kemendikdasmen"],
];
const SDSMP = [
  ["SD Negeri 008 Balikpapan Kota", ["SD inklusif", "SD"], "30402860", ["Negeri", "A"], "Balikpapan Kota", "Jl. Jend. Sudirman RT 22 No. 03", "0542-410201", "sdn.008balikpapankota@gmail.com", 15, 302, "Kepala Disdikbud, Juli 2019. Belum dikonfirmasi dinas"],
  ["SD Negeri 003 Balikpapan Kota", ["SD inklusif", "SD"], "30401549", ["Negeri", "A"], "Balikpapan Kota", "Jl. Wiluyo Puspoyudo", "Belum tercatat", "sdn003_balsel@yahoo.co.id", 29, 639, "Sosialisasi sekolah inklusi, 19 November 2024"],
  ["SMP Negeri 5 Balikpapan", ["SMP inklusif", "SMP"], "30401474", ["Negeri", "A"], "Balikpapan Selatan", "Jl. Prona II No. 104 RT 24", "0542-764142", "balikpapansmpn5@gmail.com", 49, 1287, "Disebut semua sumber berita dan jurnal. Belum dikonfirmasi dinas"],
  ["SMP Negeri 4 Balikpapan", ["SMP inklusif", "SMP"], "30401475", ["Negeri", "A"], "Balikpapan Barat", "Jl. Bukit Pelajar I, Sidodadi", "0542-734042", "smp4_bpp@yahoo.com", 50, 1195, "Disdikbud Juli 2019 dan jurnal 2023. Belum dikonfirmasi dinas"],
  ["SD IT Luqman Al-Hakim", ["SD inklusif", "SD"], "30402848", ["Swasta", "A"], "Balikpapan Tengah", "Jl. Sultan Alauddin RT 01 No. 26, Karang Bugis Dalam", "0542-441930", "sditluqmanalhakimbpp@gmail.com", 13, 162, "Penelitian inklusi 2023. Belum dikonfirmasi dinas"],
];
const J = "Lampiran 9 Juknis SPMB 2025";
const SMASMK = [
  ["SMA Negeri 1 Balikpapan", ["SMA inklusif", "SMA"], "30401491", ["Negeri", "A"], "Balikpapan Kota", "Jl. Kapten Piere Tendean RT 30 No. 63", "0542-422742", "sman1bpp@gmail.com", 57, 1369, J],
  ["SMA Negeri 2 Balikpapan", ["SMA inklusif", "SMA"], "30401490", ["Negeri", "A"], "Balikpapan Utara", "Jl. Soekarno Hatta, Strat IV", "0542-424686", "school@sman2bpp.sch.id", 51, 1052, J],
  ["SMA Negeri 4 Balikpapan", ["SMA inklusif", "SMA"], "30401488", ["Negeri", "A"], "Balikpapan Selatan", "Jl. Sepinggan Baru III No. 36", "0542-780044", "sman4bpn@gmail.com", 47, 1186, J],
  ["SMA Negeri 7 Balikpapan", ["SMA inklusif", "SMA"], "30401504", ["Negeri", "A"], "Balikpapan Timur", "Jl. Mulawarman", "0542-790025 *", "sma7balikpapan@gmail.com", 53, 1089, J],
  ["SMA Negeri 8 Balikpapan", ["SMA inklusif", "SMA"], "30402920", ["Negeri", "A"], "Balikpapan Barat", "Jl. AMD Gn. IV RT 14", "0542-7584515 *", "info@sman8bpp.sch.id", 45, 869, J],
  ["SMK Negeri 4 Balikpapan", ["SMK inklusif", "SMK"], "30401508", ["Negeri", "A"], "Balikpapan Selatan", "Jl. Belibis 3 No. 01", "0542-873890", "school@smkn4balikpapan.sch.id", 67, 1260, J],
  ["SMK Negeri 5 Balikpapan", ["SMK inklusif", "SMK"], "30404599", ["Negeri", "B"], "Balikpapan Timur", "Jl. Mulawarman No. 65", "0542-790767", "smkn5bpn@gmail.com", 70, 1229, J],
  ["SMK Negeri 6 Balikpapan", ["SMK inklusif", "SMK"], "30409918", ["Negeri", "B"], "Balikpapan Utara", "Jl. Soekarno Hatta KM 7 RT 61, Graha Indah", "0542-8830033", "sekolah@smkn6-bpn.sch.id", 75, 1414, J],
];
const GROUPS = [
  { label: "A. SLB: 4 sekolah. Pembina: Cabang Dinas Pendidikan Wilayah I (provinsi)", rows: SLB },
  { label: "B. SD dan SMP inklusif: 5 sekolah tercatat. Pembina: Dinas Pendidikan dan Kebudayaan Kota Balikpapan", rows: SDSMP },
  { label: "C. SMA dan SMK inklusif: 8 sekolah negeri. Pembina: Cabang Dinas Pendidikan Wilayah I (provinsi)", rows: SMASMK },
];

const rp = (n) => n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
const sum = (a, i) => a.reduce((t, r) => t + r[i], 0);
const all = GROUPS.flatMap((g) => g.rows);

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
    verticalAlign: VerticalAlign.TOP, margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: lines.map((t, i) => new Paragraph({
      alignment: o.align || AlignmentType.LEFT, spacing: { after: 20, line: 240 },
      children: runs(t, { size: o.size || 15, bold: o.bold || (o.boldFirst && i === 0), color: o.color || INK }),
    })),
  });
}

const widths = [400, 1850, 1000, 1000, 950, 1000, 2100, 1300, 2300, 560, 740, 2238];
if (widths.reduce((a, b) => a + b, 0) !== W) throw new Error("lebar kolom " + widths.reduce((a, b) => a + b, 0) + " tidak sama dengan " + W);
const header = ["No", "Nama sekolah", "Jenis dan jenjang", "NPSN", "Status, akreditasi", "Kecamatan", "Alamat", "Telepon", "Surel", "Guru", "Siswa", "Dasar status"];
const RIGHT = [9, 10];

const headRow = new TableRow({
  tableHeader: true, cantSplit: true,
  children: header.map((t, i) => cell(t, widths[i], { fill: NAVY, color: "FFFFFF", bold: true, size: 15, align: RIGHT.includes(i) ? AlignmentType.RIGHT : undefined })),
});
const bodyRows = [];
let no = 0;
for (const g of GROUPS) {
  bodyRows.push(new TableRow({ cantSplit: true, children: [cell(g.label, W, { span: widths.length, fill: "DCE6F2", bold: true, color: NAVY, size: 16 })] }));
  g.rows.forEach((r, ri) => {
    no += 1;
    const [nama, jj, npsn, st, kec, alamat, tel, surel, guru, siswa, dasar] = r;
    const fill = ri % 2 ? "F7F9FC" : undefined;
    const vals = [String(no), nama, jj[0] === "SLB" ? jj : jj[0], npsn, st.join(", "), kec, alamat, tel, surel, String(guru), rp(siswa), dasar];
    bodyRows.push(new TableRow({
      cantSplit: true,
      children: vals.map((v, i) => cell(v, widths[i], { fill, boldFirst: i === 1, bold: i === 1, align: RIGHT.includes(i) ? AlignmentType.RIGHT : undefined })),
    }));
  });
}
const totalVals = ["", `**Jumlah ${all.length} sekolah**`, "", "", "", "", "", "", "", `**${sum(all, 8)}**`, `**${rp(sum(all, 9))}**`, ""];
bodyRows.push(new TableRow({
  cantSplit: true,
  children: totalVals.map((v, i) => cell(v, widths[i], { fill: "DCE6F2", align: RIGHT.includes(i) ? AlignmentType.RIGHT : undefined })),
}));
const tabel = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows: [headRow, ...bodyRows] });

const p = (text, o = {}) => new Paragraph({
  children: runs(text, o.run || {}), spacing: { after: o.after ?? 100, before: o.before ?? 0, line: 270 }, keepNext: o.keepNext,
});
const note = (text) => new Paragraph({ children: runs(text, { size: 15, color: MUTED }), spacing: { before: 40, after: 70, line: 250 } });

const children = [
  new Paragraph({ heading: "Heading1", children: [new TextRun({ text: "Daftar Sekolah SLB dan Sekolah Inklusif di Kota Balikpapan" })] }),
  p(`Daftar ini memuat ${all.length} sekolah: 4 SLB, 5 SD dan SMP inklusif yang tercatat, dan 8 SMA dan SMK inklusif. Data dikumpulkan 15 Oktober 2026 untuk proposal pelatihan guru SLB dan inklusif Balikpapan, lalu disusun ulang menjadi satu tabel pada 16 Oktober 2026.`, { after: 140 }),
  tabel,
  note("* Nomor tercatat tanpa kode area di sumber. Kami menambahkan kode area Balikpapan (0542). Kolom telepon dan faks SLB di sumber berisi nomor yang sama, jadi telepon dulu untuk memastikan nomor itu aktif sebelum dipakai sebagai kontak undangan."),
  note("Guru dan siswa: portal Sekolah Kita, Kemendikdasmen, semester 2026/2027 Gasal. NPSN, alamat, telepon, surel, dan akreditasi: Referensi Kemendikdasmen dan Sekolah Kita. Daftar SMA dan SMK inklusif: Lampiran 9 Petunjuk Teknis SPMB SMA, SMK, dan SLB Kota Balikpapan 2025 (cek apakah Juknis 2026 mengubahnya)."),
  note("Kelompok B belum resmi. Keputusan Wali Kota yang menetapkan SD dan SMP inklusif belum terbit di sumber publik. Lima sekolah di tabel berasal dari pernyataan Kepala Dinas (2019), berita 2024, dan penelitian 2023. SDN 010 sempat disebut berita 2020 tanpa kecamatan dan ada di keenam kecamatan, jadi tidak dimasukkan."),
  note("Kecamatan Balikpapan Timur dan Balikpapan Utara belum punya SD atau SMP inklusif yang tercatat. [ISI TIM: nama sekolah dari Keputusan Wali Kota, diminta ke Disdikbud Kota Balikpapan]"),
  note("Instansi pembina. Cabang Dinas Pendidikan Wilayah I: Jl. A.W. Syahrani, Batu Ampar, Balikpapan Utara, cabdindik.kaltim1@gmail.com. Disdikbud Kota Balikpapan: Jl. Ruhui Rahayu I, Sepinggan, Balikpapan Selatan, 0812-5591-5591, disdik@balikpapan.go.id."),
];

const footer = new Footer({
  children: [new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: LINE, space: 6 } },
    tabStops: [{ type: TabStopType.RIGHT, position: W }],
    children: [
      new TextRun({ text: "Sanggabiz  |  Daftar Sekolah SLB dan Inklusif Kota Balikpapan", size: 14, color: MUTED }),
      new TextRun({ text: "\t", size: 14 }),
      new TextRun({ children: ["Halaman ", PageNumber.CURRENT], size: 14, color: MUTED }),
    ],
  })],
});

const doc = new Document({
  creator: "Sanggabiz", title: "Daftar Sekolah SLB dan Sekolah Inklusif di Kota Balikpapan",
  styles: {
    default: { document: { run: { font: FONT, size: 19, color: INK } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: NAVY },
        paragraph: { spacing: { before: 0, after: 200 }, outlineLevel: 0, border: { bottom: { style: BorderStyle.SINGLE, size: 14, color: GOLD, space: 6 } } } },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 900, bottom: 900, left: MARGIN, right: MARGIN } } },
    footers: { default: footer }, children,
  }],
});

const outDir = path.join(__dirname, "..", "daftar-sekolah");
fs.mkdirSync(outDir, { recursive: true });
const out = path.join(outDir, "v1 Daftar Sekolah SLB dan Inklusif Kota Balikpapan.docx");
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log("ditulis:", out, "| sekolah:", all.length, "guru:", sum(all, 8), "siswa:", rp(sum(all, 9)));
});
