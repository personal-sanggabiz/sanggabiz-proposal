// Membuat bagian "Sekolah Sasaran, Kontak, dan Jalur Pendaftaran Peserta" (docx) untuk ditempel manual ke proposal utama.
// Jalankan: NODE_PATH=<folder node_modules berisi docx> node build_bagian_sasaran.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Footer,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageNumber, LevelFormat, VerticalAlign, TabStopType,
} = require("docx");

const NAVY = "0B396D", INK = "1D1D1F", MUTED = "5C5F66", GOLD = "FFC00C", LINE = "C9D2DE";
const FONT = "Montserrat";
const W = 9638;

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
const h1 = (text) => new Paragraph({ heading: "Heading1", children: [new TextRun({ text })], keepNext: true });
const h2 = (text) => new Paragraph({ heading: "Heading2", children: [new TextRun({ text })], keepNext: true });
const bullet = (text) => new Paragraph({ numbering: { reference: "bul", level: 0 }, children: runs(text), spacing: { after: 70, line: 290 } });
const note = (text) => new Paragraph({ children: runs(text, { size: 16, color: MUTED }), spacing: { before: 60, after: 160, line: 260 } });

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
function table(widths, header, rows) {
  const head = new TableRow({ tableHeader: true, cantSplit: true, children: header.map((t, i) => cell(t, widths[i], { fill: NAVY, color: "FFFFFF", bold: true, size: 16 })) });
  const body = rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], { fill: ri % 2 ? "F7F9FC" : undefined, bold: false })) }));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows: [head, ...body] });
}

// ---------- isi ----------
const isi = [
  h1("Sekolah Sasaran, Kontak, dan Jalur Pendaftaran Peserta"),
  p("Workshop ini mengundang 80 guru dari sekolah luar biasa (SLB) dan sekolah inklusif di Kota Balikpapan. Bagian ini menyebut sekolah sasaran beserta kontaknya, instansi yang membina mereka, dan cara kami mengundang peserta. Data sekolah berasal dari Referensi Data Pendidikan Kemendikdasmen dan dokumen resmi dinas pendidikan, diakses pada 15 Oktober 2026 [S1, S2]."),

  h2("A. Tiga kelompok sekolah dan instansi pembinanya"),
  p("Pendidikan khusus dan pendidikan menengah menjadi urusan pemerintah provinsi. SD dan SMP menjadi urusan pemerintah kota [S5]. Karena itu undangan kami berjalan lewat dua instansi: Cabang Dinas Pendidikan Wilayah I yang mewakili Dinas Pendidikan dan Kebudayaan Provinsi Kalimantan Timur di Balikpapan, dan Dinas Pendidikan dan Kebudayaan Kota Balikpapan."),
  table([1900, 3400, 2338, 2000],
    ["Kelompok", "Sekolah sasaran", "Instansi pembina", "Dasar data"],
    [
      ["**1. SLB** (jenjang SDLB sampai SMLB)", "4 sekolah: SLB Negeri Balikpapan, SLB Tunas Bangsa, SLB Dharma Kencana, Sekolah Khusus Santo Fransiskus Assisi (daftar di bagian B)", "Cabang Dinas Pendidikan Wilayah I", "Referensi Kemendikdasmen [S1]"],
      ["**2. SMA dan SMK inklusif**", "8 sekolah negeri: SMA Negeri 1, 2, 4, 7, 8 dan SMK Negeri 4, 5, 6 Balikpapan (daftar di bagian C)", "Cabang Dinas Pendidikan Wilayah I", "Lampiran Juknis SPMB 2025 [S2]"],
      ["**3. SD dan SMP inklusif**", "Sekolah yang ditetapkan dengan Keputusan Wali Kota, paling sedikit satu di tiap kecamatan (Balikpapan Selatan, Timur, Utara, Tengah, Barat, dan Kota)", "Dinas Pendidikan dan Kebudayaan Kota Balikpapan", "Peraturan Wali Kota Nomor 9 Tahun 2018 [S3]"],
    ]),
  p("", { after: 60 }),

  h2("B. Daftar SLB di Kota Balikpapan"),
  p("Referensi Kemendikdasmen mencatat empat SLB di Balikpapan, dua di Kecamatan Balikpapan Selatan dan dua di Kecamatan Balikpapan Kota [S1]. Seluruhnya kami undang."),
  table([380, 2250, 2650, 2658, 1700],
    ["No", "Sekolah", "Alamat", "Kontak", "Status dan jenjang"],
    [
      ["1", ["**SLB Negeri Balikpapan**", "NPSN 30401525"], "Jl. Letkol Pol. HM. Asnawi Arbain RT 14, Kel. Sepinggan Raya, Kec. Balikpapan Selatan", ["Telepon 0542-8879984", "slbn.balikpapan@yahoo.com", "slbnbalikpapan.sch.id"], ["Negeri", "Akreditasi A", "TKLB, SDLB, SMPLB, SMLB"]],
      ["2", ["**SLB Tunas Bangsa**", "NPSN 30402883"], "Jl. Kutilang RSS Damai III RT 28, Kel. Gunung Bahagia, Kec. Balikpapan Selatan", ["Telepon 0542-8870057 *", "tunasbangsa831@yahoo.co.id", "slbtunasbangsabalikpapan.sch.id"], ["Swasta", "Akreditasi B", "SDLB, SMPLB, SMLB"]],
      ["3", ["**SLB Dharma Kencana Balikpapan**", "NPSN 30402882"], "Jl. Pembangunan No. 30 Blok A, Kel. Telaga Sari, Kec. Balikpapan Kota", ["Telepon 0542-410406", "slbdk@yahoo.co.id"], ["Swasta", "Akreditasi B", "SDLB, SMPLB, SMLB"]],
      ["4", ["**Sekolah Khusus Santo Fransiskus Assisi**", "NPSN 69965143"], "Jl. Mayor (Pol) Zaenal Arifin No. 40 RT 18, Kel. Damai, Kec. Balikpapan Kota", ["Telepon 0542-8794310", "skhfransiskusassisi@gmail.com"], ["Swasta", "Akreditasi B", "TKLB, SDLB, SMPLB, SMLB"]],
    ]),
  note("* Data sumber mencantumkan nomor 8870057 tanpa kode area. Kami menambahkan kode area Balikpapan (0542). Sumber: Referensi Kemendikdasmen, halaman NPSN masing-masing sekolah [S1]."),
  p("Jumlah siswa yang tercatat di sumber publik baru untuk SLB Negeri Balikpapan, yaitu 401 siswa (219 SD, 96 SMP, 86 SMA) [S6]. Data siswa dan guru tiga SLB lainnya kami minta langsung dari sekolah melalui formulir pendaftaran, sehingga susunan kelompok kerja sesuai kondisi sekolah."),

  h2("C. Sekolah inklusif"),
  p("**SMA dan SMK.** Delapan sekolah negeri berikut tercantum sebagai penyelenggara pendidikan inklusif pada Lampiran 9 Petunjuk Teknis SPMB SMA, SMK, dan SLB Kota Balikpapan Tahun 2025 [S2]. Penetapannya mengacu pada Surat Keputusan Kepala Dinas Pendidikan dan Kebudayaan Provinsi Kalimantan Timur Nomor 400.3/12060/DISDIKBUD.V/2023."),
  table([2000, 7638],
    ["Jenjang", "Sekolah"],
    [
      ["SMA", "SMA Negeri 1 Balikpapan, SMA Negeri 2 Balikpapan, SMA Negeri 4 Balikpapan, SMA Negeri 7 Balikpapan, SMA Negeri 8 Balikpapan"],
      ["SMK", "SMK Negeri 4 Balikpapan, SMK Negeri 5 Balikpapan, SMK Negeri 6 Balikpapan"],
    ]),
  p("", { after: 60 }),
  p("**SD dan SMP.** Peraturan Wali Kota Balikpapan Nomor 9 Tahun 2018 mewajibkan setiap kecamatan menyelenggarakan paling sedikit satu pendidikan inklusif pada jenjang SD dan SMP, dan menetapkan sekolahnya dengan Keputusan Wali Kota (Pasal 4 ayat 3 dan ayat 5) [S3]. Kami memohon daftar sekolah yang sudah ditetapkan itu kepada Dinas Pendidikan dan Kebudayaan Kota Balikpapan. [ISI TIM: daftar SD dan SMP penyelenggara inklusif beserta nomor Keputusan Wali Kota, setelah diterima dari Disdikbud Kota]"),
  p("Pasal 12 sampai 14 peraturan yang sama membagi pendidik inklusif menjadi guru kelas atau guru mata pelajaran dan Guru Pembimbing Khusus (GPK). Kami mengundang keduanya, karena modul ajar berdiferensiasi disusun guru kelas bersama GPK."),

  h2("D. Kontak instansi pembina"),
  table([2300, 3100, 2438, 1800],
    ["Instansi", "Alamat", "Kontak", "Peran dalam program"],
    [
      [["**Cabang Dinas Pendidikan Wilayah I**", "Dinas Pendidikan dan Kebudayaan Provinsi Kalimantan Timur"], "Jl. A.W. Syahrani, Batu Ampar, Kec. Balikpapan Utara. Wilayah kerja: Kota Balikpapan dan Kabupaten Penajam Paser Utara", ["Surel cabdindik.kaltim1@gmail.com", "cabdin1.disdikbud.kaltimprov.go.id"], "Rekomendasi dan undangan untuk SLB serta SMA dan SMK"],
      [["**Dinas Pendidikan dan Kebudayaan Kota Balikpapan**"], "Jl. Ruhui Rahayu I, Kel. Sepinggan, Kec. Balikpapan Selatan", ["Telepon 0812-5591-5591", "Surel disdik@balikpapan.go.id"], "Rekomendasi, daftar sekolah, dan undangan untuk SD dan SMP inklusif"],
    ]),
  note("Sumber: Petunjuk Teknis SPMB 2025 untuk Cabang Dinas [S2] dan situs Dinas Pendidikan dan Kebudayaan Kota Balikpapan [S4]."),

  h2("E. Cara kami mengundang 80 peserta"),
  p("Daftar guru peserta belum tersedia, sehingga rekrutmen menjadi pekerjaan kami. Alurnya sebagai berikut."),
  p("1. **Surat permohonan.** Setelah tanggal pelaksanaan ditetapkan, kami mengirim surat permohonan rekomendasi dan data sekolah kepada Cabang Dinas Pendidikan Wilayah I dan Dinas Pendidikan dan Kebudayaan Kota Balikpapan.", { indent: { left: 360, hanging: 360 } }),
  p("2. **Undangan ke sekolah.** Undangan resmi dikirim kepada kepala 4 SLB, 8 SMA dan SMK inklusif, serta SD dan SMP inklusif. Tiap sekolah menunjuk guru yang memegang kelas dengan murid berkebutuhan khusus. GPK diprioritaskan.", { indent: { left: 360, hanging: 360 } }),
  p("3. **Pendaftaran lewat formulir.** Guru mengisi nama, sekolah, jenjang, mata pelajaran, lama mengajar, jenis kebutuhan murid di kelasnya, dan kebutuhan akses bagi dirinya sendiri. Jawaban formulir menjadi pemetaan awal untuk pemateri.", { indent: { left: 360, hanging: 360 } }),
  p("4. **Konfirmasi dan daftar tunggu.** Tim kami mengonfirmasi kehadiran lewat WhatsApp sebelum acara. Sekolah yang berhalangan digantikan oleh guru dari daftar tunggu.", { indent: { left: 360, hanging: 360 } }),
  p("5. **Susunan peserta.** [ISI TIM: kuota peserta per kelompok sekolah (SLB, SMA dan SMK, SD dan SMP) dengan jumlah total 80 orang]", { indent: { left: 360, hanging: 360 } }),

  h2("Rujukan bagian ini"),
  p("[S1] Kemendikdasmen. Referensi Data Pendidikan, halaman satuan pendidikan per NPSN 30401525, 30402883, 30402882, dan 69965143, serta daftar program SLB Kota Balikpapan. referensi.data.kemendikdasmen.go.id. Diakses 15 Oktober 2026.", { run: { size: 16 }, after: 50, line: 250 }),
  p("[S2] Cabang Dinas Pendidikan Wilayah I Provinsi Kalimantan Timur (2025). Petunjuk Teknis SPMB SMA, SMK, dan SLB Kota Balikpapan Tahun Pelajaran 2025/2026, Lampiran 9. SK Nomor 400.3.8/370.2/Cabdisdikwil.1/2025, 22 April 2025. berkas.spmb.id/cabdinbalikpapan/juknis.250521095215.pdf.", { run: { size: 16 }, after: 50, line: 250 }),
  p("[S3] Peraturan Wali Kota Balikpapan Nomor 9 Tahun 2018 tentang Penyelenggaraan Pendidikan Inklusif. peraturan.bpk.go.id/Download/88906/PERWALI%20NOMOR%209%202018.pdf.", { run: { size: 16 }, after: 50, line: 250 }),
  p("[S4] Dinas Pendidikan dan Kebudayaan Kota Balikpapan. Halaman kontak. disdikbud.balikpapan.go.id. Diakses 15 Oktober 2026.", { run: { size: 16 }, after: 50, line: 250 }),
  p("[S5] Undang-Undang Nomor 23 Tahun 2014 tentang Pemerintahan Daerah, Lampiran huruf A, pembagian urusan pemerintahan bidang pendidikan.", { run: { size: 16 }, after: 50, line: 250 }),
  p("[S6] ANTARA Kaltim. PT Angkasa Pura I Sepinggan Balikpapan beri bantuan ke SLBN. kaltim.antaranews.com/berita/207852.", { run: { size: 16 }, after: 50, line: 250 }),
];

const footer = new Footer({
  children: [new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: LINE, space: 6 } },
    tabStops: [{ type: TabStopType.RIGHT, position: W }],
    children: [
      new TextRun({ text: "Sanggabiz  |  Bagian Sekolah Sasaran dan Kontak", size: 14, color: MUTED }),
      new TextRun({ text: "\t", size: 14 }),
      new TextRun({ children: ["Halaman ", PageNumber.CURRENT], size: 14, color: MUTED }),
    ],
  })],
});

const doc = new Document({
  creator: "Sanggabiz", title: "Sekolah Sasaran, Kontak, dan Jalur Pendaftaran Peserta",
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
    footers: { default: footer }, children: isi,
  }],
});

const outDir = path.join(__dirname, "..", "bagian-sasaran");
fs.mkdirSync(outDir, { recursive: true });
const out = path.join(outDir, "v1 Bagian Sekolah Sasaran dan Kontak - SLB dan Inklusif Balikpapan.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ditulis:", out); });
