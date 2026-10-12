// Membuat proposal docx. Jalankan: NODE_PATH=$(npm root -g) node build_proposal.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, PageBreak, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, Footer,
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
      alignment: o.align || AlignmentType.LEFT, spacing: { after: o.gap ?? 30, line: 260 },
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
        fill: c && typeof c.bg === "string" ? c.bg : (ri % 2 ? "F7F9FC" : undefined), align: al, size: o.size, gap: o.gap, bold: !!(c && typeof c === "object" && !Array.isArray(c) && c.bold === true),
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
    ["C.2", "Advokasi Sekolah", "1", "Person", "10", "Sekolah", 500000],
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

// ---------- bagian 3 dan 4, versi revisi (waktu longgar, aktivitas disederhanakan) ----------
const FIRST = (lines) => ({ text: lines, bg: "EEF3F9" });
function sesi(judul, waktu, topik) {
  const widths = [1350, 1650, 2300, 2688, 1650];
  return [
    h3(judul),
    p(waktu, { after: 70, keepNext: true, run: { color: MUTED, size: 17 } }),
    table(widths, ["Topik dan durasi", "Tujuan pembelajaran", "Outline materi", "Aktivitas", "Output dan ukuran keberhasilan"],
      topik.map((t) => [FIRST([t[0]]), t[1], t[2], t[3], t[4]]), { size: 16, gap: 60 }),
  ];
}

const bab3 = [
  h1("3. Silabus Materi"),
  h2("3.1 Gambaran Program"),
  table([2400, 7238], ["Hal", "Keterangan"], [
    ["Fokus", "Menyusun modul ajar berdiferensiasi untuk kelas dengan murid berkebutuhan beragam"],
    ["Peserta", "80 guru SLB dan sekolah inklusif"],
    ["Format", "Tatap muka, 5 jam acara inti. Materi sekitar 40 persen dan praktik sekitar 60 persen. Setiap sesi memiliki waktu longgar untuk instruksi dan peralihan kegiatan"],
    ["Pengajar", "1 pemateri lokal"],
    ["Pelaksana", "1 project manager, 1 MC dan moderator"],
    ["Alat bantu", "AI sebagai alat bantu merapikan draf. Isi pokok modul dikerjakan guru"],
    ["Hasil", "Template, modul ringkas, draf inti modul ajar sendiri, rencana tindak lanjut, e-sertifikat"],
  ]),
  h2("3.2 Pendekatan Pelatihan"),
  p("Program disusun dengan metode yang lazim dipakai dalam pelatihan orang dewasa, dipilih agar guru pulang membawa hasil yang nyata."),
  table([2600, 7038], ["Metode", "Penerapan di workshop ini"], [
    [["**ADDIE** [11]", "Analysis, Design, Development, Implementation, Evaluation"], "Analisis lewat formulir pendaftaran dan pre-test. Rancangan lewat silabus dan rundown. Pengembangan berupa template, modul, dan bahan ajar. Pelaksanaan berupa workshop setengah hari. Evaluasi lewat post-test, rubrik, kuesioner, dan tindak lanjut empat minggu."],
    [["**Andragogi** [12]", "Knowles"], "Peserta belajar dari kasus murid di kelasnya sendiri. Pemateri berperan sebagai fasilitator, dan porsi praktik lebih besar dari materi."],
    [["**Taksonomi Bloom revisi** [13]", "Anderson dan Krathwohl"], "Tujuan tiap topik memakai kata kerja yang dapat diamati, misalnya menyebutkan, menjelaskan, dan menyusun."],
    [["**AI sebagai alat bantu**", "Bukan bahan utama"], "Guru menyusun draf kasar sendiri pada template. Setelah itu AI dipakai untuk merapikannya menjadi versi yang lebih enak dibaca dan siap dipakai. Guru memeriksa hasilnya, dan data pribadi murid tidak dimasukkan ke AI."],
    [["**Kirkpatrick** [14]", "Empat level"], "Reaksi diukur lewat kuesioner, pembelajaran lewat pre-test, post-test, dan rubrik produk, perilaku lewat laporan penerapan di grup tindak lanjut."],
  ]),
  h2("3.3 Silabus per Sesi"),
  p("Silabus terdiri dari lima sesi dan dua belas topik. Setiap sesi sudah menyediakan waktu longgar, sehingga bila ada kegiatan yang berjalan lebih lama, sesi tetap selesai tepat waktu."),
  table([900, 4338, 1400, 3000], ["Sesi", "Topik", "Durasi", "Hasil utama"], [
    ["1", "Ragam kebutuhan belajar dan prinsip pembelajaran berdiferensiasi", "80 menit", "Kartu kebutuhan murid, hasil pencocokan, satu ide penerapan"],
    ["2", "Membaca profil belajar murid", "45 menit", "Lembar profil satu murid"],
    ["3", "Workshop menyusun modul ajar berdiferensiasi", "90 menit", "Draf inti modul ajar"],
    ["4", "Presentasi mini dan umpan balik", "20 menit", "Daftar perbaikan prioritas"],
    ["5", "Rencana tindak lanjut dan evaluasi", "20 menit", "Rencana dua minggu, data post-test"],
  ], { align: [AlignmentType.CENTER, AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.LEFT] }),

  ...sesi("Sesi 1. Ragam Kebutuhan Belajar dan Prinsip Pembelajaran Berdiferensiasi", "08.20 - 09.40 WIB, 80 menit, materi, termasuk waktu longgar sekitar 11 menit", [
    ["**Topik 1.1** Ragam kebutuhan belajar murid (25 menit)",
      "Peserta mampu menyebutkan lima kelompok kebutuhan belajar dan memberi satu contoh dari kelasnya.",
      ["1. Arti kebutuhan dan hambatan belajar", "2. Lima kelompok kebutuhan: komunikasi dan bahasa, akademik dan kognitif, sensorik, motorik, sosial dan emosi", "3. Lima contoh profil murid"],
      ["Pemateri memaparkan dengan contoh di layar.", "Tiap peserta menulis satu murid dari kelasnya pada kartu tanpa nama, lalu menempelkannya pada poster meja yang berisi lima zona kebutuhan.", "Peserta membaca poster mejanya bersama-sama."],
      ["**Output:** kartu murid di poster meja dan satu catatan pribadi di buku kerja.", "**Ukuran:** setiap peserta menempelkan satu kartu pada zona yang sesuai."]],
    ["**Topik 1.2** Prinsip dan empat unsur pembelajaran berdiferensiasi (30 menit)",
      "Peserta mampu menjelaskan prinsip pembelajaran berdiferensiasi dan mengelompokkan contoh kegiatan ke empat unsurnya.",
      ["1. Prinsip: guru menyesuaikan pembelajaran dengan kesiapan, minat, dan profil belajar murid", "2. Empat unsur: konten, proses, produk, lingkungan belajar", "3. Kaitan dengan capaian pembelajaran Kurikulum Merdeka dan akomodasi yang layak"],
      ["Pemateri memaparkan dengan satu contoh pelajaran, yaitu pecahan di kelas 4 SD.", "Tiap meja menerima 8 kartu contoh kegiatan dan mencocokkannya pada poster empat unsur.", "Pemateri membahas jawaban bersama."],
      ["**Output:** hasil pencocokan 8 kartu tiap meja.", "**Ukuran:** tiap meja menjawab benar minimal 6 dari 8 kartu."]],
    ["**Topik 1.3** Contoh penerapan di kelas campuran (25 menit)",
      "Peserta mampu mengenali bagian contoh yang dapat dipakai di kelasnya dan menuliskan satu ide penerapan.",
      ["1. Contoh di SLB: satu topik untuk tiga kelompok kemampuan", "2. Contoh di sekolah inklusif: satu kelas reguler dengan murid beragam", "3. Satu topik dengan ragam konten, proses, dan produk"],
      ["Pemateri menayangkan dua contoh, lalu menjawab pertanyaan peserta.", "Tiap peserta menulis satu ide penerapan di buku kerja: bagian mana yang bisa dipakai di kelasnya."],
      ["**Output:** satu ide penerapan di buku kerja.", "**Ukuran:** setiap peserta menuliskan satu ide."]],
  ]),

  ...sesi("Sesi 2. Membaca Profil Belajar Murid", "09.55 - 10.40 WIB, 45 menit (20 menit materi, 25 menit praktik), termasuk waktu longgar sekitar 9 menit", [
    ["**Topik 2.1** Asesmen diagnostik sederhana (20 menit)",
      "Peserta mampu menjelaskan tujuan asesmen diagnostik dan menyebutkan tiga cara sederhana untuk melakukannya.",
      ["1. Tujuan asesmen di awal pembelajaran", "2. Tiga cara: observasi lima menit, percakapan singkat dengan murid, informasi dari orang tua atau guru sebelumnya", "3. Lima pertanyaan kunci: apa yang sudah bisa, apa yang membantu, apa yang menyulitkan, apa minatnya, bagaimana cara berkomunikasinya", "4. Menjaga kerahasiaan data murid"],
      ["Pemateri memaparkan dengan contoh lembar profil yang sudah terisi, lalu peserta bertanya.", "Lembar profil kosong sudah ada di buku kerja."],
      ["**Output:** pemahaman cara mengisi lembar profil.", "**Ukuran:** butir terkait di post-test dijawab benar oleh minimal 80 persen peserta."]],
    ["**Topik 2.2** Mengisi lembar profil satu murid (25 menit)",
      "Peserta mampu menyusun lembar profil belajar untuk satu murid dari kelasnya.",
      ["1. Bagian lembar profil: kekuatan, kebutuhan, minat, strategi yang membantu, target empat minggu", "2. Cara menulis kalimat yang spesifik dan dapat diamati"],
      ["Tiap peserta memilih satu murid dari kelasnya dan mengisi lembar profil tanpa menulis nama murid.", "Bila waktu cukup, peserta menukar lembar dengan teman sebelah dan memberi satu masukan."],
      ["**Output:** satu lembar profil murid.", "**Ukuran:** minimal 80 persen peserta mengisi seluruh bagian lembar."]],
  ]),

  ...sesi("Sesi 3. Workshop Menyusun Modul Ajar Berdiferensiasi", "10.40 - 12.10 WIB, 90 menit, praktik, termasuk waktu longgar sekitar 18 menit", [
    ["**Topik 3.1** Menentukan tujuan dan murid sasaran (20 menit)",
      "Peserta mampu menulis satu tujuan pembelajaran yang dapat diamati dan menentukan murid sasaran.",
      ["1. Memilih satu capaian pembelajaran dari mata pelajaran yang diampu", "2. Menulis tujuan dengan kata kerja yang dapat diamati, misalnya menyebutkan, menjelaskan, membuat", "3. Memilih dua atau tiga murid sasaran dari lembar profil"],
      ["Peserta duduk dalam 8 kelompok berisi 10 guru. Tiap kelompok menunjuk satu ketua yang membacakan langkah dari template dan menjaga waktu. Pemateri dan project manager berkeliling.", "Tiap guru mengisi bagian Identitas dan Tujuan pada template, lalu memeriksa tujuannya dengan daftar periksa yang tercetak di template."],
      ["**Output:** bagian Identitas dan Tujuan pada template.", "**Ukuran:** setiap peserta memiliki satu tujuan yang memenuhi daftar periksa."]],
    ["**Topik 3.2** Merancang ragam konten dan ragam proses (30 menit)",
      "Peserta mampu merancang dua ragam konten dan dua ragam proses belajar untuk satu topik.",
      ["1. Dua ragam konten, misalnya teks bergambar, teks sederhana, benda nyata", "2. Dua ragam proses, misalnya kerja mandiri bertahap, kerja berpasangan, bimbingan langsung guru", "3. Setiap ragam menuju tujuan yang sama"],
      ["Peserta mengisi bagian Diferensiasi Konten dan Diferensiasi Proses pada template. Pertanyaan pemandu sudah tercetak di template.", "Pemateri dan project manager berkeliling, membantu kelompok atau guru yang tertahan, dan memberi contoh tambahan bila dibutuhkan."],
      ["**Output:** bagian konten dan proses pada template.", "**Ukuran:** setiap peserta memiliki dua ragam konten dan dua ragam proses."]],
    ["**Topik 3.3** Merancang produk, lingkungan belajar, dan akomodasi (40 menit)",
      "Peserta mampu menyusun satu pilihan produk, satu penataan lingkungan belajar, dan satu akomodasi untuk melengkapi draf inti modul ajar.",
      ["1. Tiga contoh pilihan produk: lisan, tulisan, karya", "2. Penataan kelas dan dukungan teman sebaya", "3. Akomodasi yang layak untuk murid tertentu"],
      ["Peserta melengkapi bagian Produk, Lingkungan Belajar, dan Akomodasi pada template, lalu memeriksa kelengkapan dengan daftar periksa.", "Tiap peserta memotret drafnya sendiri untuk dikirim ke grup tindak lanjut di Sesi 5, sebagai bahan penilaian tim setelah acara.", "Di akhir topik, tiap kelompok memilih satu draf untuk dipajang di Sesi 4. Bagian asesmen dan refleksi dilengkapi peserta saat tindak lanjut."],
      ["**Output:** draf inti modul ajar pada template.", "**Ukuran:** minimal 80 persen peserta mengisi seluruh bagian inti, dinilai tim dengan rubrik empat butir dari foto draf setelah acara."]],
  ]),

  ...sesi("Sesi 4. Presentasi Mini dan Umpan Balik", "12.10 - 12.30 WIB, 20 menit, termasuk waktu longgar sekitar 4 menit", [
    ["**Topik 4.1** Gallery walk (10 menit)",
      "Peserta mampu memberi umpan balik tertulis pada hasil kerja kelompok lain.",
      ["1. Cara memberi umpan balik: satu hal yang sudah baik dan satu saran"],
      ["Draf pilihan tiap kelompok sudah dipajang di dinding.", "Peserta berkeliling bebas ke dua draf yang ingin dilihat dan menempelkan satu catatan pada tiap draf berisi dua baris: satu hal yang sudah baik dan satu saran."],
      ["**Output:** catatan umpan balik pada draf yang dipajang.", "**Ukuran:** tiap draf menerima minimal 5 catatan."]],
    ["**Topik 4.2** Umpan balik pemateri (6 menit)",
      "Peserta mampu menyebutkan perbaikan prioritas untuk draf modulnya.",
      ["1. Tiga temuan umum dari seluruh draf", "2. Satu contoh perbaikan"],
      ["Pemateri dan project manager mencatat temuan selama gallery walk.", "Pemateri membahas tiga temuan umum dan menunjukkan satu contoh perbaikan di layar, termasuk cara meminta AI merapikan draf kasar menjadi versi yang lebih enak dibaca. Peserta mencatat perbaikan untuk drafnya di buku kerja."],
      ["**Output:** daftar perbaikan prioritas di buku kerja.", "**Ukuran:** setiap peserta mencatat minimal satu perbaikan."]],
  ]),

  ...sesi("Sesi 5. Rencana Tindak Lanjut dan Evaluasi", "12.30 - 12.50 WIB, 20 menit, termasuk waktu longgar sekitar 4 menit", [
    ["**Topik 5.1** Rencana tindak lanjut dua minggu (8 menit)",
      "Peserta mampu menyusun rencana menerapkan modulnya di kelas dalam dua minggu.",
      ["1. Tiga pertanyaan rencana: kapan, di kelas mana, apa yang dicoba"],
      ["Peserta menulis rencana pada lembar satu halaman di buku kerja. Rencana itu memuat jadwal merapikan draf dengan bantuan AI, lalu mencobanya di kelas."],
      ["**Output:** rencana tindak lanjut satu halaman.", "**Ukuran:** setiap peserta menuliskan rencana."]],
    ["**Topik 5.2** Post-test, kuesioner, dan grup tindak lanjut (12 menit)",
      "Peserta mampu menunjukkan peningkatan pemahaman pada post-test.",
      ["1. Post-test 10 butir dengan soal yang sama seperti pre-test", "2. Kuesioner kepuasan 4 pertanyaan", "3. Bergabung ke grup tindak lanjut"],
      ["Peserta memindai satu kode QR penutup di layar. Kode itu membuka grup tindak lanjut, post-test, dan kuesioner. Setelah bergabung, peserta mengirim foto draf ke grup.", "Kertas cadangan disediakan bagi peserta tanpa ponsel atau sinyal. Project manager dan MC membantu peserta yang membutuhkan."],
      ["**Output:** data post-test dan kuesioner dari seluruh peserta, keanggotaan grup tindak lanjut, dan foto draf.", "**Ukuran:** rata-rata skor post-test naik minimal 20 persen dari pre-test, dan minimal 75 persen peserta bergabung di grup sebelum acara berakhir."]],
  ]),

  h2("3.4 Hasil yang Dibawa Pulang"),
  p("Setiap hasil di bawah ini dihasilkan pada topik yang disebutkan, sehingga peserta dapat menelusuri asalnya.", { keepNext: true }),
  table([3400, 2500, 3738], ["Hasil", "Dihasilkan pada", "Kegunaan"], [
    ["Catatan pribadi kebutuhan murid", "Topik 1.1", "Memetakan kebutuhan belajar murid di kelas sendiri"],
    ["Satu ide penerapan", "Topik 1.3", "Bahan awal merancang modul"],
    ["Lembar profil satu murid", "Topik 2.2", "Dasar menyusun modul untuk murid sasaran"],
    ["Draf inti modul ajar berdiferensiasi", "Topik 3.1 sampai 3.3", "Dipakai di kelas dan dilengkapi saat tindak lanjut"],
    ["Daftar perbaikan prioritas", "Topik 4.2", "Menyempurnakan modul sebelum dipakai"],
    ["Panduan memakai AI sebagai alat bantu", "Modul ringkas, dibahas di Topik 4.2", "Merapikan draf kasar menjadi versi siap pakai"],
    ["Rencana tindak lanjut dua minggu", "Topik 5.1", "Panduan menerapkan modul di kelas"],
    ["Buku kerja peserta", "Dibagikan saat registrasi", "Menyimpan semua catatan dan lembar kerja hari itu"],
    ["Template modul ajar berdiferensiasi, cetak dan digital", "Dibagikan saat registrasi", "Dipakai untuk menyusun modul berikutnya"],
    ["Modul ringkas sekitar 20 halaman", "Dibagikan digital lewat grup tindak lanjut", "Rujukan setelah workshop (isi di Lampiran B)"],
    ["E-sertifikat peserta", "Dikirim setelah acara", "Bukti keikutsertaan"],
  ]),
  h2("3.5 Evaluasi dan Tindak Lanjut"),
  table([1500, 2400, 3738, 2000], ["Level", "Yang diukur", "Alat", "Waktu"], [
    ["1. Reaksi", "Kepuasan peserta", "Kuesioner 4 pertanyaan", "Akhir acara"],
    ["2. Pembelajaran", "Pengetahuan dan keterampilan", "Pre-test dan post-test 10 soal yang sama, rubrik draf inti 4 butir", "Awal dan akhir acara"],
    ["3. Perilaku", "Penerapan di kelas", "Laporan peserta di grup tindak lanjut (foto atau cerita singkat)", "Minggu 1 sampai 4"],
  ]),
  p("Dampak pada murid (level 4) berada di luar cakupan program setengah hari ini dan dapat dirancang pada tahap lanjutan bila pemberi dana menghendaki.", { before: 100 }),
  p("Grup WhatsApp tindak lanjut berjalan empat minggu dan dimoderatori project manager untuk tanya jawab dan berbagi hasil penerapan. Pada masa itu peserta melengkapi bagian asesmen dan refleksi pada template, lalu merapikan drafnya dengan bantuan AI dan membagikan versi akhirnya di grup. Laporan akhir kami serahkan dalam 7 hari kerja setelah acara, berisi daftar hadir, hasil tes, penilaian draf, hasil kuesioner, dan dokumentasi."),
];

const bab4 = [
  h1("4. Rundown Acara"),
  p("Jam di bawah mengikuti usulan awal. Bila tanggal dan tempat berubah, jam dapat digeser tanpa mengubah durasi tiap sesi. Setiap sesi sudah menyediakan waktu longgar untuk instruksi dan peralihan kegiatan."),
  table([1700, 1100, 4700, 2138], ["Waktu (WIB)", "Durasi", "Kegiatan", "Pemandu"], [
    ["07.30 - 08.00", "30 menit", "Registrasi dan absensi lewat formulir daring (pre-test di formulir yang sama), pembagian ID card dan perlengkapan", "Project manager dan MC"],
    ["08.00 - 08.20", "20 menit", "Pembukaan dan sambutan, pengantar program", "MC dan Sanggabiz"],
    ["08.20 - 09.40", "80 menit", "**Sesi 1.** Ragam kebutuhan belajar dan prinsip pembelajaran berdiferensiasi", "Pemateri"],
    ["09.40 - 09.55", "15 menit", "Coffee break", "Project manager"],
    ["09.55 - 10.40", "45 menit", "**Sesi 2.** Membaca profil belajar murid", "Pemateri"],
    ["10.40 - 12.10", "90 menit", "**Sesi 3.** Workshop menyusun modul ajar berdiferensiasi", "Pemateri dan project manager"],
    ["12.10 - 12.30", "20 menit", "**Sesi 4.** Presentasi mini dan umpan balik", "Pemateri"],
    ["12.30 - 12.50", "20 menit", "**Sesi 5.** Rencana tindak lanjut, post-test, kuesioner", "Sanggabiz"],
    ["12.50 - 13.00", "10 menit", "Penutupan dan informasi e-sertifikat", "MC"],
    ["13.00 - 13.45", "45 menit", "Makan siang dan ramah tamah", "Project manager"],
  ], { align: [AlignmentType.CENTER, AlignmentType.CENTER] }),
  p("Pembagian waktu belajar: sesi materi 100 menit (sekitar 40 persen) dan sesi praktik 155 menit (sekitar 60 persen) dari 255 menit. Di dalamnya tersedia waktu longgar sekitar 45 menit.", { before: 140 }),
];

const surat = [
  h1("Surat Pengantar"),
  p("Nomor: [ISI TIM: nomor surat]", { after: 20 }),
  p("Lampiran: 1 berkas proposal", { after: 20 }),
  p("Perihal: **Penawaran Workshop Pembelajaran Berdiferensiasi untuk Guru SLB dan Sekolah Inklusif di Balikpapan**", { after: 240 }),
  p("Kepada Yth.", { after: 20 }),
  p("[ISI TIM: nama dan jabatan penerima]", { after: 20 }),
  p("[ISI TIM: instansi penerima]", { after: 20 }),
  p("di Balikpapan", { after: 240 }),
  p("Dengan hormat,"),
  p("Sanggabiz menyampaikan penawaran workshop setengah hari untuk 80 guru SLB dan sekolah inklusif di Balikpapan. Workshop ini berfokus pada satu kemampuan, yaitu menyusun modul ajar berdiferensiasi yang dapat langsung dipakai di kelas. Setiap peserta pulang membawa template dan modul ringkas, serta satu draf modul ajar yang sudah mulai disusun dalam sesi. Peserta juga diperkenalkan pada AI sebagai alat bantu untuk merapikan draf menjadi versi yang siap dipakai. Isi pokok modul tetap dikerjakan guru."),
  p("Proposal ini memuat latar belakang, tujuan, tempat dan peserta, silabus, rundown, anggaran, dan profil pemateri. Kami siap menyesuaikan jadwal dan isi sesuai arahan Bapak/Ibu."),
  p("Demikian penawaran ini kami sampaikan. Terima kasih atas perhatiannya.", { after: 300 }),
  p("Hormat kami,", { after: 20 }),
  p("**Sanggabiz**", { after: 700 }),
  p("[ISI TIM: tanda tangan dan nama penandatangan]", { after: 20 }),
  p("[ISI TIM: jabatan penandatangan]", { after: 20 }),
  p("Kontak: [ISI TIM: nomor telepon dan surel]"),
];

const bab1 = [
  h1("1. Pendahuluan", true),
  h2("1.1 Latar Belakang"),
  p("Pendidikan inklusif menempatkan anak berkebutuhan khusus di kelas yang sama dengan teman sebayanya. Agar berjalan baik, guru perlu menyesuaikan cara mengajar dengan kebutuhan tiap murid. Menurut data Kemendikbudristek per Desember 2023, ada 40.164 satuan pendidikan formal di Indonesia yang memiliki peserta didik berkebutuhan khusus. Dari jumlah itu, 5.956 satuan pendidikan (14,83 persen) sudah memiliki guru pembimbing khusus [1]."),
  centerImg("grafik-guru-pembimbing-khusus.png", 560, 192),
  caption("Gambar 1. Satuan pendidikan dengan peserta didik berkebutuhan khusus menurut ketersediaan guru pembimbing khusus"),
  p("Guru yang mendampingi anak berkebutuhan khusus di sekolah inklusif banyak yang berlatar PGSD, PGPAUD, atau bidang nonkependidikan, dan hanya sebagian yang lulusan Pendidikan Luar Biasa. Mereka mendampingi banyak murid sekaligus, dan pelatihan yang tersedia masih perlu diperluas [2]."),
  p("Kurikulum Merdeka meminta guru merancang pembelajaran sesuai kesiapan, minat, dan profil belajar murid. Kajian di beberapa sekolah dasar inklusif mencatat kendala yang sama: pemahaman guru tentang kebutuhan belajar anak masih perlu diperkuat, pelatihan perlu ditambah, waktu persiapan ringkas, dan jumlah murid per kelas besar. Pembelajaran berdiferensiasi sudah mulai dijalankan, tetapi belum mencakup semua aspek, terutama konten dan lingkungan belajar [3][4][5]."),
  p("Kebutuhan serupa tercatat di Kalimantan Timur. Komisi IV DPRD Kaltim menyampaikan bahwa guru berlatar pendidikan luar biasa perlu ditambah, dan mengusulkan pembukaan program studi PLB bersama Universitas Mulawarman [6]. SLB negeri berada di bawah Dinas Pendidikan dan Kebudayaan Provinsi, yang mengelola 11 SLB di Kaltim [7]. Salah satunya SLB Negeri Balikpapan, berdiri sejak 1984 dan melayani 401 siswa [8]. Sekolah inklusif di Balikpapan berada di bawah Pemerintah Kota, yang telah menerbitkan Peraturan Wali Kota Nomor 3 Tahun 2013 tentang penyelenggaraan pendidikan inklusif [9]."),
  centerImg("grafik-siswa-slbn-balikpapan.png", 520, 213),
  caption("Gambar 2. Jumlah siswa SLB Negeri Balikpapan per jenjang"),
  p("Peraturan Pemerintah Nomor 13 Tahun 2020 mewajibkan pemerintah pusat dan daerah memfasilitasi akomodasi yang layak bagi peserta didik penyandang disabilitas, termasuk menyiapkan pendidik dan menyediakan kurikulum [10]. Penguatan kemampuan guru adalah bagian dari kewajiban itu, dan workshop ini dirancang untuk mendukung pemenuhannya."),
  h2("1.2 Mengapa Pembelajaran Berdiferensiasi"),
  p("Kami memilih satu kemampuan agar peserta pulang dengan hasil yang nyata. Ada empat alasan."),
  bullet("Setiap guru di kelas inklusif mengajar murid dengan kebutuhan berbeda dalam satu ruangan. Menyesuaikan konten, proses, dan produk belajar adalah pekerjaan harian mereka."),
  bullet("Hasilnya dapat dibuat dalam setengah hari. Peserta menyusun satu modul ajar memakai template, lalu membawanya pulang."),
  bullet("Materinya berlaku untuk semua jenis kekhususan, sehingga sesuai untuk peserta campuran dari SLB dan sekolah inklusif."),
  bullet("Waktu persiapan guru ringkas. Karena itu peserta diperkenalkan pada AI sebagai alat bantu: setelah draf kasar selesai, AI dapat merapikannya menjadi versi yang lebih enak dibaca dan siap dipakai. AI bukan bahan utama pembelajaran, dan isi pokok modul tetap dikerjakan guru."),
  p("Data khusus Balikpapan tentang tingkat pemahaman guru terhadap pembelajaran berdiferensiasi belum kami temukan. Karena itu workshop dibuka dengan pre-test. Hasilnya mencatat posisi awal peserta dan menjadi pembanding untuk hasil akhir.", { before: 80 }),
  h2("1.3 Tujuan"),
  p("**Tujuan umum.** Meningkatkan kemampuan 80 guru SLB dan sekolah inklusif di Balikpapan dalam merancang pembelajaran yang sesuai dengan kebutuhan beragam muridnya."),
  p("**Tujuan khusus.** Setelah workshop, peserta mampu:", { after: 60 }),
  bullet("menjelaskan empat unsur pembelajaran berdiferensiasi, yaitu konten, proses, produk, dan lingkungan belajar,"),
  bullet("membaca profil belajar seorang murid dengan asesmen diagnostik sederhana,"),
  bullet("menyusun modul ajar berdiferensiasi memakai template yang disediakan,"),
  bullet("menyusun rencana penerapan untuk dua minggu pertama di kelas masing-masing,"),
  bullet("memakai AI sebagai alat bantu untuk merapikan draf modul, dengan isi pokok tetap buatan guru."),
  h2("1.4 Indikator Keberhasilan"),
  table([2200, 3700, 3738], ["Indikator", "Target", "Cara mengukur"], [
    ["Kehadiran", "80 guru hadir", "Daftar hadir"],
    ["Kenaikan pemahaman", "Rata-rata skor post-test naik minimal 20 persen dari pre-test", "Pre-test dan post-test, 10 soal yang sama"],
    ["Produk", "Minimal 80 persen peserta menghasilkan draf inti modul ajar", "Rubrik 4 butir: kesesuaian profil murid, ragam konten, ragam proses, pilihan produk"],
    ["Kepuasan", "Rata-rata minimal 4 dari 5", "Kuesioner akhir"],
    ["Tindak lanjut", "Minimal 75 persen peserta bergabung di grup tindak lanjut", "Data anggota grup"],
  ]),
  h2("1.5 Tentang Sanggabiz"),
  p("Sanggabiz adalah mitra solusi bisnis yang menyediakan layanan department-as-a-service: keuangan, sumber daya manusia, layanan pelanggan, pemasaran, manajemen proyek, administrasi, dan operasional. Di bidang pelatihan, kami mengelola program dari perencanaan sampai pelaporan."),
  table([2400, 7238], ["Hal", "Keterangan"], [
    ["Skala", "400+ proyek dan layanan, 500+ pakar dan trainer, 6.000+ mitra engagement"],
    ["Pengalaman multikota", "Pelatihan MSME Rumah BUMN di lebih dari sembilan kota, termasuk Samarinda"],
    ["Klien", "Lembaga pendidikan, BUMN, yayasan, dan pemerintah daerah, antara lain Universitas Muhammadiyah Yogyakarta, Yakes Pertamina, dan Pemerintah Kota Yogyakarta"],
    ["Pendekatan pelatihan", "ADDIE, andragogi, taksonomi Bloom revisi, evaluasi Kirkpatrick (lihat bagian 3.2)"],
  ]),
  p("Sanggabiz belum memiliki program khusus pendidikan luar biasa. Karena itu isi materi dipegang pemateri lokal yang berpengalaman di bidang pendidikan inklusif. Kami mengerjakan bagian yang kami kuasai: perekrutan peserta, logistik, penyusunan modul, evaluasi, dan pelaporan. Dalam menyusun template dan modul ringkas, kami juga memakai AI sebagai alat bantu penulisan, dan pemateri memeriksa isinya sebelum dibagikan.", { before: 120 }),
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [4819, 4819],
    rows: [new TableRow({ children: [["foto-pelatihan-1.jpg", 227, 170], ["foto-pelatihan-2.jpg", 302, 170]].map(([f, w, h]) => new TableCell({
      width: { size: 4819, type: WidthType.DXA },
      borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [img(f, w, h, "jpg")] })],
    })) })],
  }),
  caption("Gambar 3. Dokumentasi pelatihan yang diselenggarakan Sanggabiz"),
];

const bab2 = [
  h1("2. Tempat, Waktu, dan Peserta", true),
  h2("2.1 Tempat Pelaksanaan"),
  p("Workshop diadakan di Balikpapan, di tempat yang mudah dijangkau dari SLB dan sekolah di berbagai kecamatan. Anggaran tidak memuat biaya sewa tempat, sehingga tempat disediakan tanpa biaya oleh pihak penyedia."),
  table([2600, 7038], ["Hal", "Keterangan"], [
    ["Tempat", "[ISI TIM: nama tempat dan alamat di Balikpapan]"],
    ["Ruang", "[ISI TIM: nama ruang dan luas]"],
    ["Penyedia tempat", "[ISI TIM: nama pihak penyedia tempat]"],
    ["Kapasitas yang dibutuhkan", "Minimal 80 orang, dengan susunan 8 kelompok"],
    ["Fasilitas yang dibutuhkan", "Proyektor, tata suara, akses internet, lahan parkir"],
    ["Akses disabilitas", "[ISI TIM: lift atau jalur kursi roda dan toilet]"],
  ]),
  spacer(120),
  centerImg("skema-susunan-ruang.png", 520, 292),
  caption("Gambar 4. Skema susunan ruang untuk 80 peserta"),
  h2("2.2 Waktu Pelaksanaan"),
  table([3000, 6638], ["Hal", "Keterangan"], [
    ["Tanggal", "[ISI TIM: tanggal pelaksanaan]"],
    ["Jam", "07.30 sampai 13.45 WIB, termasuk registrasi dan makan siang"],
    ["Acara inti", "08.00 sampai 13.00 WIB (5 jam)"],
  ]),
  h2("2.3 Peserta dan Cara Rekrutmen"),
  p("Peserta adalah 80 guru SLB dan sekolah inklusif di Balikpapan, dari semua jenjang dan semua jenis kekhususan. Kuota dibagi seimbang antara guru SLB dan guru sekolah inklusif, lalu disesuaikan dengan jumlah pendaftar. Peserta tidak membayar, karena biaya ditanggung pemberi dana."),
  p("Sanggabiz yang mencari dan mengonfirmasi peserta. Jalur yang kami pakai:", { after: 60 }),
  bullet("kunjungan langsung ke 10 sekolah sasaran oleh petugas advokasi, yaitu pemateri lokal atau tim Sanggabiz yang mengenal sekolah-sekolah tersebut,"),
  bullet("surat undangan dari dinas pendidikan melalui pemberi dana,"),
  bullet("kelompok kerja guru dan musyawarah kerja kepala sekolah (KKG dan MKKS),"),
  bullet("kepala SLB negeri dan swasta di Balikpapan."),
  table([2200, 7438], ["Waktu", "Kegiatan rekrutmen"], [
    ["H-21", "Petugas advokasi mengunjungi 10 sekolah sasaran dan menyerahkan undangan. Undangan juga dikirim ke dinas, KKG, MKKS, dan kepala sekolah. Formulir pendaftaran daring dibuka"],
    ["H-14", "Pendaftaran ditutup bila kuota penuh. Formulir mencatat jenjang, mata pelajaran, dan kendala utama peserta"],
    ["H-7", "Project manager mengonfirmasi kehadiran lewat WhatsApp. Peserta cadangan dihubungi bila ada yang berhalangan"],
    ["H-3", "Project manager mengirim pengingat berisi lokasi, jam, dan petunjuk akses"],
    ["Hari-H", "Registrasi dan absensi lewat formulir daring yang sama dengan pre-test. ID card sudah disiapkan dan disusun per sekolah"],
  ], { align: [AlignmentType.CENTER, AlignmentType.LEFT] }),
];

const bab5 = [
  h1("5. Anggaran", true),
  h2("5.1 Dasar Penyusunan"),
  bullet("Peserta 80 orang. Snack dan makan siang dihitung untuk 85 orang, yaitu peserta, pemateri, MC, project manager, dan petugas lainnya. Perlengkapan peserta dan e-sertifikat dihitung untuk 80 orang."),
  bullet("Modul/Materi mencakup penyusunan modul ringkas dan template, bahan ajar interaktif (kartu pencocokan, poster kelompok, lembar kerja), serta cetak buku kerja peserta."),
  bullet("Biaya pemateri, MC, dan project manager dihitung per hari kegiatan. Pemateri membawakan 4 jam tatap muka."),
  bullet("Advokasi sekolah mencakup kunjungan ke 10 sekolah sasaran untuk menyerahkan undangan dan menjelaskan program. Sewa mobil mendukung kunjungan itu dan kebutuhan hari-H."),
  bullet("Monitoring dan reporting mencakup pengolahan pre-test dan post-test, penilaian draf peserta, pemantauan grup tindak lanjut empat minggu, dan laporan akhir."),
  bullet("Biaya sewa tempat tidak termasuk dalam anggaran ini."),
  bullet("Seluruh harga sudah termasuk pajak."),
  new Paragraph({ heading: "Heading2", pageBreakBefore: true, keepNext: true, children: [new TextRun({ text: "5.2 Rincian Anggaran" })] }),
  tabel,
  new Paragraph({ spacing: { before: 160, after: 40 }, children: [new TextRun({ text: "Catatan *)", bold: true, size: 17 })] }),
  p("Anggaran di atas sudah termasuk pajak.", { after: 20, run: { size: 17 } }),
  p("Anggaran bersifat estimatif dan masih dapat dinegosiasikan sesuai kesepakatan bersama.", { run: { size: 17 } }),
];

const bab6 = [
  h1("6. Pemateri", true),
  p("Pemateri berasal dari Balikpapan atau Kalimantan Timur. Dengan begitu tidak ada biaya mobilisasi, dan jaringan lokalnya dapat dipakai untuk tindak lanjut. Kriteria yang kami pakai:", { after: 60 }),
  bullet("berdomisili di Balikpapan atau Kalimantan Timur,"),
  bullet("berlatar pendidikan luar biasa atau pendidikan inklusif, atau praktisi dengan pengalaman nyata mendampingi murid berkebutuhan khusus,"),
  bullet("pernah melatih guru atau memandu workshop,"),
  bullet("pernah menyusun modul ajar atau program pembelajaran individual."),
  spacer(80),
  (() => {
    const lab = (t) => ({ text: t, bold: true, bg: LIGHT });
    return table([2600, 7038], ["Pemateri: Pakar Pendidikan Inklusif", ""], [
      [lab("Nama dan gelar"), "[ISI TIM: nama lengkap dan gelar]"],
      [lab("Jabatan dan instansi"), "[ISI TIM: jabatan dan instansi]"],
      [lab("Bidang keahlian"), "[ISI TIM: bidang keahlian]"],
      [lab("Riwayat pendidikan"), "[ISI TIM: pendidikan terakhir dan bidang]"],
      [lab("Pengalaman melatih guru"), "[ISI TIM: pengalaman melatih guru, jumlah pelatihan, tahun]"],
      [lab("Pengalaman menyusun modul atau program pembelajaran individual"), "[ISI TIM: pengalaman terkait]"],
      [lab("Peran di workshop"), "Membawakan Sesi 1 sampai 4: materi, pemandu workshop, dan umpan balik"],
      [lab("Foto"), "[ISI TIM: tempel foto pemateri di sini]"],
    ]);
  })(),
];

const bab7 = [
  h1("7. Penutup", true),
  p("Workshop ini dirancang terfokus: satu kemampuan, satu hari kerja, satu produk yang dapat dipakai. Kami berharap proposal ini menjadi dasar diskusi lebih lanjut. Jadwal, jumlah peserta, dan isi sesi dapat kami sesuaikan dengan arahan Bapak/Ibu."),
  p("Kontak: [ISI TIM: nama, nomor telepon, dan surel penghubung]"),
  h2("Lampiran A. Contoh Template Modul Ajar Berdiferensiasi"),
  p("Template ini dibagikan kepada peserta dan diisi pada Sesi 3.", { after: 100 }),
  table([3000, 6638], ["Bagian", "Diisi guru"], [
    ["Identitas (inti)", "Mata pelajaran, fase atau kelas, alokasi waktu, nama guru"],
    ["Capaian dan tujuan pembelajaran (inti)", ""],
    ["Profil belajar murid (inti)", "Kesiapan, minat, kebutuhan khusus, hal yang membantu murid belajar"],
    ["Diferensiasi konten (inti)", "Ragam A: .......   Ragam B: ......."],
    ["Diferensiasi proses (inti)", "Ragam A: .......   Ragam B: ......."],
    ["Diferensiasi produk (inti)", "Pilihan produk yang boleh dipilih murid"],
    ["Lingkungan belajar (inti)", "Penataan kelas, alat bantu, dukungan teman sebaya"],
    ["Asesmen", "Asesmen diagnostik dan asesmen formatif"],
    ["Akomodasi yang layak (inti)", "Penyesuaian yang disiapkan untuk murid tertentu"],
    ["Refleksi", "Apa yang berhasil dan apa yang diubah pada pertemuan berikutnya"],
  ]),
  p("Bagian bertanda (inti) diisi saat workshop. Asesmen dan Refleksi dilengkapi saat tindak lanjut. Setelah semua bagian terisi, peserta dapat meminta AI merapikan draf menjadi versi yang lebih enak dibaca. Isi pokoknya tetap dari guru.", { before: 100 }),
  h2("Lampiran B. Isi Modul Ringkas"),
  bullet("Prinsip pembelajaran berdiferensiasi dan empat unsurnya."),
  bullet("Ragam kebutuhan belajar murid dan cara mengenalinya."),
  bullet("Asesmen diagnostik sederhana beserta lembar profil murid."),
  bullet("Langkah menyusun modul ajar berdiferensiasi."),
  bullet("Dua contoh modul ajar yang sudah terisi."),
  bullet("Panduan memakai AI sebagai alat bantu: contoh perintah untuk merapikan draf, cara memeriksa hasilnya, dan hal yang tidak dimasukkan ke AI, seperti nama dan data pribadi murid."),
  bullet("Daftar periksa sebelum dan sesudah mengajar."),
  bullet("Daftar bacaan dan sumber lanjutan."),
  h2("Lampiran C. Rujukan"),
  ...[
    "Kemendikbudristek, data satuan pendidikan dengan peserta didik berkebutuhan khusus per Desember 2023. Medcom, \"34 Ribu Sekolah dengan Siswa Disabilitas Tak Punya Guru Pembimbing Khusus\". https://www.medcom.id/pendidikan/cerita-guru/RkjGDw9K-34-ribu-sekolah-dengan-siswa-disabilitas-tak-punya-guru-pembimbing-khusus. Juga ANTARA, https://www.antaranews.com/berita/4038030",
    "Artikel jurnal Universitas Negeri Malang tentang guru pendamping khusus di sekolah inklusif. https://journal2.um.ac.id/index.php/jo/article/download/59676/15100",
    "Penerapan Pembelajaran Berdiferensiasi dan Inklusif di Kelas. Jurnal Universitas Tanjungpura. https://jurnal.untan.ac.id/index.php/jpdpb/article/download/87400/75676603823",
    "Implementasi Pendekatan Berdiferensiasi untuk Anak Berkebutuhan Khusus di SD Mardi Waluya. SIPENDAS 2025, 1(2): 11-18. https://jurnal.stkipkusumanegara.ac.id/index.php/sipendas/article/view/2439",
    "Implementasi Pembelajaran Berdiferensiasi untuk Anak Berkebutuhan Khusus di Sekolah Dasar. Jurnal Cendikia. https://jurnal.kolibi.org/index.php/cendikia/article/view/4655",
    "ANTARA Kaltim, \"DPRD Kaltim minta Pemprov tingkatkan kualitas SLB\". https://kaltim.antaranews.com/berita/197400/dprd-kaltim-minta-pemprov-tingkatkan-kualitas-slb",
    "Sonora.id, \"Disdikbud Kaltim Sasar SLB di Kaltim sebagai Upaya Pemerataan Pendidikan\". https://www.sonora.id/read/424272404/disdikbud-kaltim-sasar-slb-di-kaltim-sebagai-upaya-pemerataan-pendidikan",
    "ANTARA Kaltim, \"PT Angkasa Pura I Sepinggan Balikpapan beri bantuan ke SLBN\". https://kaltim.antaranews.com/berita/207852/pt-angkasa-pura-i-sepinggan-balikpapan-beri-bantuan-ke-slbn",
    "Implementation of Inclusive Education in Elementary Schools in Balikpapan City: Challenges, Solutions, and Impacts. https://journal-fip.um.ac.id/index.php/icsar/article/view/3272",
    "Peraturan Pemerintah Nomor 13 Tahun 2020 tentang Akomodasi yang Layak untuk Peserta Didik Penyandang Disabilitas. https://peraturan.go.id/files/pp13-2020pjl.pdf",
    "Model ADDIE (Analysis, Design, Development, Implementation and Evaluation). Jurnal Ilmiah Pendidikan Agama Islam, UIN Sunan Gunung Djati. https://journal.uinsgd.ac.id/index.php/jipai/article/download/11042/pdf",
    "Konsep Andragogi dalam Pembelajaran Kursus dan Pelatihan Jarak Jauh: Studi Literatur. Jurnal Pendidikan Luar Sekolah, Universitas Negeri Surabaya. https://ejournal.unesa.ac.id/index.php/jurnal-pendidikan-luar-sekolah/article/view/42812",
    "Taksonomi Bloom Revisi. Universitas Atma Jaya Yogyakarta. https://dpm.uajy.ac.id/wp-content/uploads/2011/08/taksonomi_Bloom_revisi.pdf",
    "Penerapan Model Kirkpatrick untuk Evaluasi Pelatihan. Jurnal Diklat Teknis. https://media.neliti.com/media/publications/275133-penerapan-model-kirkpatrick-untuk-evalua-ced0cdda.pdf",
  ].map((t, i) => new Paragraph({ spacing: { after: 70, line: 260 }, indent: { left: 420, hanging: 420 }, children: [new TextRun({ text: `[${i + 1}]\t${t}`, size: 16 })], tabStops: [{ type: TabStopType.LEFT, position: 420 }] })),
];

const cover = [
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({
      height: { value: 12300, rule: HeightRule.EXACT },
      children: [new TableCell({
        width: { size: W, type: WidthType.DXA },
        shading: { fill: NAVY, type: ShadingType.CLEAR, color: "auto" },
        borders: { top: { style: BorderStyle.NONE, size: 0, color: NAVY }, bottom: { style: BorderStyle.NONE, size: 0, color: NAVY }, left: { style: BorderStyle.NONE, size: 0, color: NAVY }, right: { style: BorderStyle.NONE, size: 0, color: NAVY } },
        margins: { top: 500, bottom: 400, left: 560, right: 560 },
        children: [
          new Paragraph({ spacing: { after: 2600 }, children: [img("logo_white.png", 150, 40)] }),
          new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "PROPOSAL PELATIHAN", font: FONT, bold: true, size: 20, color: GOLD, characterSpacing: 60 })] }),
          new Paragraph({ spacing: { after: 240, line: 400, lineRule: "atLeast" }, children: [new TextRun({ text: "Workshop Pembelajaran Berdiferensiasi untuk Guru SLB dan Sekolah Inklusif", font: FONT, bold: true, size: 52, color: "FFFFFF" })] }),
          new Paragraph({ spacing: { after: 2400 }, border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: GOLD, space: 14 } }, children: [new TextRun({ text: "Capacity building 80 guru di Balikpapan, Kalimantan Timur", font: FONT, size: 24, color: "DCE6F2" })] }),
          new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Diajukan oleh Sanggabiz", font: FONT, bold: true, size: 22, color: "FFFFFF" })] }),
          new Paragraph({ children: [new TextRun({ text: "2026", font: FONT, size: 22, color: "DCE6F2" })] }),
        ],
      })],
    })],
  }),
  spacer(300),
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [3212, 3213, 3213],
    rows: [new TableRow({ children: [
      ["80 guru", "SLB dan sekolah inklusif di Balikpapan"],
      ["Setengah hari", "Tatap muka, 40 persen materi dan 60 persen praktik"],
      ["Dibawa pulang", "Template dan modul ringkas, draf inti modul ajar sendiri"],
    ].map(([a, b], i) => new TableCell({
      width: { size: 3212 + (i ? 1 : 0), type: WidthType.DXA },
      borders: { top: { style: BorderStyle.SINGLE, size: 18, color: GOLD }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
      margins: { top: 120, bottom: 60, left: 120, right: 160 },
      children: [
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: a, bold: true, size: 26, color: NAVY })] }),
        new Paragraph({ children: [new TextRun({ text: b, size: 17, color: MUTED })] }),
      ],
    })) })],
  }),
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

const page = { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } };
const doc = new Document({
  creator: "Sanggabiz", title: "Workshop Pembelajaran Berdiferensiasi untuk Guru SLB dan Sekolah Inklusif",
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
  sections: [
    { properties: { page: { ...page, margin: { top: 700, bottom: 700, left: 1134, right: 1134 } } }, children: cover },
    { properties: { page }, footers: { default: footer }, children: [...surat, ...bab1, ...bab2, ...bab3, ...bab4, ...bab5, ...bab6, ...bab7] },
  ],
});

const outDir = path.join(__dirname, "..", "proposal");
fs.mkdirSync(outDir, { recursive: true });
const out = path.join(outDir, "v3 Proposal Workshop Pembelajaran Berdiferensiasi - SLB dan Inklusif Balikpapan.docx");
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log("ditulis:", out);
  console.log('v3 selesai');
});
