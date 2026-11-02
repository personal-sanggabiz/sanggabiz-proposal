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
    ["Hasil", "Template, modul ringkas, draf inti modul ajar sendiri, rencana tindak lanjut, sertifikat"],
  ]),
  h2("3.2 Pendekatan Pelatihan"),
  p("Program disusun dengan metode yang lazim dipakai dalam pelatihan orang dewasa, dipilih agar guru pulang membawa hasil yang nyata."),
  table([2600, 7038], ["Metode", "Penerapan di workshop ini"], [
    [["**ADDIE** [11]", "Analysis, Design, Development, Implementation, Evaluation"], "Analisis lewat formulir pendaftaran dan pre-test. Rancangan lewat silabus dan rundown. Pengembangan berupa template, modul, dan bahan ajar. Pelaksanaan berupa workshop setengah hari. Evaluasi lewat post-test, rubrik, kuesioner, dan tindak lanjut empat minggu."],
    [["**Andragogi** [12]", "Knowles"], "Peserta belajar dari kasus murid di kelasnya sendiri. Pemateri berperan sebagai fasilitator, dan porsi praktik lebih besar dari materi."],
    [["**Taksonomi Bloom revisi** [13]", "Anderson dan Krathwohl"], "Tujuan tiap topik memakai kata kerja yang dapat diamati, misalnya menyebutkan, menjelaskan, dan menyusun."],
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
      ["Peserta duduk dalam 8 kelompok berisi 10 guru. Tiap pendamping memandu dua kelompok.", "Tiap guru mengisi bagian Identitas dan Tujuan pada template, lalu memeriksa tujuannya dengan daftar periksa yang tercetak di template."],
      ["**Output:** bagian Identitas dan Tujuan pada template.", "**Ukuran:** setiap peserta memiliki satu tujuan yang memenuhi daftar periksa."]],
    ["**Topik 3.2** Merancang ragam konten dan ragam proses (30 menit)",
      "Peserta mampu merancang dua ragam konten dan dua ragam proses belajar untuk satu topik.",
      ["1. Dua ragam konten, misalnya teks bergambar, teks sederhana, benda nyata", "2. Dua ragam proses, misalnya kerja mandiri bertahap, kerja berpasangan, bimbingan langsung guru", "3. Setiap ragam menuju tujuan yang sama"],
      ["Peserta mengisi bagian Diferensiasi Konten dan Diferensiasi Proses pada template. Pertanyaan pemandu sudah tercetak di template.", "Pendamping berkeliling di dua kelompoknya dan membantu guru yang tertahan. Pemateri berkeliling dan memberi contoh tambahan bila dibutuhkan."],
      ["**Output:** bagian konten dan proses pada template.", "**Ukuran:** setiap peserta memiliki dua ragam konten dan dua ragam proses."]],
    ["**Topik 3.3** Merancang produk, lingkungan belajar, dan akomodasi (40 menit)",
      "Peserta mampu menyusun satu pilihan produk, satu penataan lingkungan belajar, dan satu akomodasi untuk melengkapi draf inti modul ajar.",
      ["1. Tiga contoh pilihan produk: lisan, tulisan, karya", "2. Penataan kelas dan dukungan teman sebaya", "3. Akomodasi yang layak untuk murid tertentu"],
      ["Peserta melengkapi bagian Produk, Lingkungan Belajar, dan Akomodasi pada template, lalu memeriksa kelengkapan dengan daftar periksa.", "Dengan izin peserta, pendamping memotret draf yang sudah terisi untuk penilaian tim setelah acara.", "Di akhir topik, tiap kelompok memilih satu draf untuk dipajang di Sesi 4. Bagian asesmen dan refleksi dilengkapi peserta saat tindak lanjut."],
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
      ["Pemateri dan pendamping mencatat temuan selama gallery walk.", "Pemateri membahas tiga temuan umum dan menunjukkan satu contoh perbaikan di layar. Peserta mencatat perbaikan untuk drafnya di buku kerja."],
      ["**Output:** daftar perbaikan prioritas di buku kerja.", "**Ukuran:** setiap peserta mencatat minimal satu perbaikan."]],
  ]),

  ...sesi("Sesi 5. Rencana Tindak Lanjut dan Evaluasi", "12.30 - 12.50 WIB, 20 menit, termasuk waktu longgar sekitar 4 menit", [
    ["**Topik 5.1** Rencana tindak lanjut dua minggu (8 menit)",
      "Peserta mampu menyusun rencana menerapkan modulnya di kelas dalam dua minggu.",
      ["1. Tiga pertanyaan rencana: kapan, di kelas mana, apa yang dicoba"],
      ["Peserta menulis rencana pada lembar satu halaman di buku kerja."],
      ["**Output:** rencana tindak lanjut satu halaman.", "**Ukuran:** setiap peserta menuliskan rencana."]],
    ["**Topik 5.2** Post-test, kuesioner, dan grup tindak lanjut (12 menit)",
      "Peserta mampu menunjukkan peningkatan pemahaman pada post-test.",
      ["1. Post-test 10 butir dengan soal yang sama seperti pre-test", "2. Kuesioner kepuasan 4 pertanyaan", "3. Bergabung ke grup tindak lanjut"],
      ["Peserta memindai satu kode QR penutup di layar. Kode itu membuka grup tindak lanjut, post-test, dan kuesioner.", "Kertas cadangan disediakan bagi peserta tanpa ponsel atau sinyal. Pendamping membantu peserta yang membutuhkan."],
      ["**Output:** data post-test dan kuesioner dari seluruh peserta, serta keanggotaan grup tindak lanjut.", "**Ukuran:** rata-rata skor post-test naik minimal 20 persen dari pre-test, dan minimal 75 persen peserta bergabung di grup sebelum acara berakhir."]],
  ]),

  h2("3.4 Hasil yang Dibawa Pulang"),
  p("Setiap hasil di bawah ini dihasilkan pada topik yang disebutkan, sehingga peserta dapat menelusuri asalnya.", { keepNext: true }),
  table([3400, 2500, 3738], ["Hasil", "Dihasilkan pada", "Kegunaan"], [
    ["Catatan pribadi kebutuhan murid", "Topik 1.1", "Memetakan kebutuhan belajar murid di kelas sendiri"],
    ["Satu ide penerapan", "Topik 1.3", "Bahan awal merancang modul"],
    ["Lembar profil satu murid", "Topik 2.2", "Dasar menyusun modul untuk murid sasaran"],
    ["Draf inti modul ajar berdiferensiasi", "Topik 3.1 sampai 3.3", "Dipakai di kelas dan dilengkapi saat tindak lanjut"],
    ["Daftar perbaikan prioritas", "Topik 4.2", "Menyempurnakan modul sebelum dipakai"],
    ["Rencana tindak lanjut dua minggu", "Topik 5.1", "Panduan menerapkan modul di kelas"],
    ["Buku kerja peserta", "Dibagikan saat registrasi", "Menyimpan semua catatan dan lembar kerja hari itu"],
    ["Template modul ajar berdiferensiasi, cetak dan digital", "Dibagikan saat registrasi", "Dipakai untuk menyusun modul berikutnya"],
    ["Modul ringkas sekitar 20 halaman", "Dibagikan saat registrasi", "Rujukan setelah workshop (isi di Lampiran B)"],
    ["Sertifikat peserta", "Penutupan", "Bukti keikutsertaan"],
  ]),
  h2("3.5 Evaluasi dan Tindak Lanjut"),
  table([1500, 2400, 3738, 2000], ["Level", "Yang diukur", "Alat", "Waktu"], [
    ["1. Reaksi", "Kepuasan peserta", "Kuesioner 4 pertanyaan", "Akhir acara"],
    ["2. Pembelajaran", "Pengetahuan dan keterampilan", "Pre-test dan post-test 10 soal yang sama, rubrik draf inti 4 butir", "Awal dan akhir acara"],
    ["3. Perilaku", "Penerapan di kelas", "Laporan peserta di grup tindak lanjut (foto atau cerita singkat)", "Minggu 1 sampai 4"],
  ]),
  p("Dampak pada murid (level 4) berada di luar cakupan program setengah hari ini dan dapat dirancang pada tahap lanjutan bila pemberi dana menghendaki.", { before: 100 }),
  p("Grup WhatsApp tindak lanjut berjalan empat minggu dan dimoderatori tim Sanggabiz untuk tanya jawab dan berbagi hasil penerapan. Pada masa itu peserta melengkapi bagian asesmen dan refleksi pada template. Laporan akhir kami serahkan dalam 7 hari kerja setelah acara, berisi daftar hadir, hasil tes, penilaian draf, hasil kuesioner, dan dokumentasi."),
];

const bab4 = [
  h1("4. Rundown Acara"),
  p("Jam di bawah mengikuti usulan awal. Bila tanggal dan tempat berubah, jam dapat digeser tanpa mengubah durasi tiap sesi. Setiap sesi sudah menyediakan waktu longgar untuk instruksi dan peralihan kegiatan."),
  table([1700, 1100, 4700, 2138], ["Waktu (WIB)", "Durasi", "Kegiatan", "Pemandu"], [
    ["07.30 - 08.00", "30 menit", "Registrasi, daftar hadir, pre-test lewat ponsel (kode QR di meja registrasi), pembagian perlengkapan", "Panitia dan IO"],
    ["08.00 - 08.20", "20 menit", "Pembukaan dan sambutan, pengantar program", "MC dan Sanggabiz"],
    ["08.20 - 09.40", "80 menit", "**Sesi 1.** Ragam kebutuhan belajar dan prinsip pembelajaran berdiferensiasi", "Pemateri"],
    ["09.40 - 09.55", "15 menit", "Coffee break", "Panitia"],
    ["09.55 - 10.40", "45 menit", "**Sesi 2.** Membaca profil belajar murid", "Pemateri"],
    ["10.40 - 12.10", "90 menit", "**Sesi 3.** Workshop menyusun modul ajar berdiferensiasi", "Pemateri dan pendamping"],
    ["12.10 - 12.30", "20 menit", "**Sesi 4.** Presentasi mini dan umpan balik", "Pemateri"],
    ["12.30 - 12.50", "20 menit", "**Sesi 5.** Rencana tindak lanjut, post-test, kuesioner", "Sanggabiz"],
    ["12.50 - 13.00", "10 menit", "Penutupan dan pembagian sertifikat", "MC"],
    ["13.00 - 13.45", "45 menit", "Makan siang dan ramah tamah", "Panitia"],
  ], { align: [AlignmentType.CENTER, AlignmentType.CENTER] }),
  p("Pembagian waktu belajar: sesi materi 100 menit (sekitar 40 persen) dan sesi praktik 155 menit (sekitar 60 persen) dari 255 menit. Di dalamnya tersedia waktu longgar sekitar 45 menit.", { before: 140 }),
];

const page = { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } };
const doc = new Document({
  creator: "Sanggabiz", title: "Bagian 3 dan 4",
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
  sections: [{ properties: { page }, children: [...bab3, ...[new Paragraph({ children: [new PageBreak()] })], ...bab4] }],
});
const out = path.join(__dirname, "..", "proposal", "Bagian 3 dan 4 revisi (untuk disalin).docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ditulis:", out); });
