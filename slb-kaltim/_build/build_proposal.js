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

// ---------- anggaran (satu opsi, harga akhir sudah termasuk pajak) ----------
const ANGGARAN = [
  { group: "A. Pemateri dan pemandu acara" },
  ["1", "Pemateri lokal, pakar pendidikan inklusif (Sesi 1 sampai 4), 4 jam tatap muka", 4, "jam", 1000000],
  ["2", "MC dan moderator", 1, "orang", 500000],
  { group: "B. Tempat dan konsumsi" },
  ["3", "Paket meeting setengah hari di Quest Hotel Balikpapan, Orchid Room (ruang, proyektor dan tata suara, coffee break, makan siang)", 90, "orang", 148500],
  { group: "C. Materi dan bahan ajar" },
  ["4", "Penyusunan modul ringkas dan template modul ajar berdiferensiasi", 1, "paket", 3500000],
  ["5", "Bahan ajar interaktif: kuis dan lembar kerja digital", 1, "paket", 3000000],
  ["6", "Cetak modul ringkas dan template (hitam putih, sampul warna)", 90, "eksemplar", 15000],
  ["7", "Perlengkapan peserta: ID card, pulpen, kartu kerja", 90, "set", 15000],
  ["8", "Sertifikat peserta", 80, "lembar", 5000],
  { group: "D. Rekrutmen, pendampingan, dan koordinasi" },
  ["9", "Rekrutmen dan koordinasi peserta: undangan, formulir pendaftaran, konfirmasi, pengingat", 1, "paket", 4000000],
  ["10", "Pendamping kelompok dan panitia lokal", 4, "orang", 300000],
  { group: "E. Evaluasi, dokumentasi, dan tindak lanjut" },
  ["11", "Pre-test, post-test, penilaian produk, dan laporan akhir", 1, "paket", 2500000],
  ["12", "Dokumentasi foto", 1, "paket", 800000],
  ["13", "Tindak lanjut empat minggu lewat grup WhatsApp", 1, "paket", 1000000],
  { group: "F. Lain-lain" },
  ["14", "Transportasi lokal dan kebutuhan hari-H", 1, "paket", 500000],
];
let TOTAL = 0;
const anggaranRows = ANGGARAN.map((r) => {
  if (r.group) return [r];
  const jumlah = r[2] * r[4]; TOTAL += jumlah;
  return [r[0], r[1], String(r[2]), r[3], fmt(r[4]), fmt(jumlah)];
});
const PER_PESERTA = Math.round(TOTAL / 80 / 1000) * 1000;
function tabelAnggaran() {
  const widths = [500, 3900, 700, 1100, 1700, 1738];
  const sum = (label, val) => [{ text: "", bg: "F2F5FA" }, { text: label, bold: true, bg: "F2F5FA" }, { text: "", bg: "F2F5FA" }, { text: "", bg: "F2F5FA" }, { text: "", bg: "F2F5FA" }, { text: val, bold: true, bg: "F2F5FA" }];
  return table(widths, ["No", "Komponen", "Vol", "Satuan", "Harga satuan", "Jumlah"], [
    ...anggaranRows,
    sum("TOTAL (sudah termasuk pajak)", fmt(TOTAL)),
    sum("Per peserta (80 orang, dibulatkan)", fmt(PER_PESERTA)),
  ], { align: [AlignmentType.CENTER, AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.RIGHT, AlignmentType.RIGHT] });
}

// ---------- silabus ----------
const SESI = [
  { judul: "Sesi 1. Ragam Kebutuhan Belajar dan Prinsip Pembelajaran Berdiferensiasi", waktu: "08.20 - 09.40 WIB, 80 menit, materi", topik: [
    ["**Topik 1.1** Ragam kebutuhan belajar murid di kelas inklusif (25 menit)",
      "Peserta mampu menyebutkan lima kelompok kebutuhan belajar dan memberi satu contoh dari kelasnya sendiri.",
      ["1. Arti kebutuhan belajar dan hambatan belajar", "2. Lima kelompok kebutuhan: komunikasi dan bahasa, akademik dan kognitif, sensorik, motorik, sosial dan emosi", "3. Satu kelas campuran dengan lima contoh profil murid"],
      ["Menit 0 sampai 10: pemateri memaparkan outline dan menayangkan lima contoh profil murid di layar.", "Menit 10 sampai 25: tiap peserta menulis satu murid dari kelasnya pada kartu tanpa nama, lalu menempelkannya di papan kelompok sesuai kelompok kebutuhan. Kartu dan spidol tersedia di meja."],
      ["Output: 80 kartu kebutuhan murid dan satu catatan pribadi per peserta.", "Ukuran: 100 persen peserta menempelkan kartu pada kelompok yang sesuai."]],
    ["**Topik 1.2** Prinsip dan empat unsur pembelajaran berdiferensiasi (30 menit)",
      "Peserta mampu menjelaskan prinsip pembelajaran berdiferensiasi dan mengelompokkan contoh kegiatan ke empat unsurnya.",
      ["1. Pengertian dan prinsip: guru menyesuaikan pembelajaran dengan kesiapan, minat, dan profil belajar murid", "2. Empat unsur: konten (apa yang dipelajari), proses (cara belajar), produk (bukti hasil belajar), lingkungan belajar (penataan kelas dan suasana)", "3. Beda diferensiasi dengan perlakuan sama untuk semua murid dan dengan pengajaran perorangan penuh", "4. Kaitan dengan capaian pembelajaran Kurikulum Merdeka dan akomodasi yang layak"],
      ["Menit 0 sampai 15: pemateri memaparkan dengan satu contoh pelajaran (pecahan di kelas 4 SD) yang punya tiga ragam konten.", "Menit 15 sampai 30: kuis interaktif di ponsel peserta lewat kode QR di layar. Ada 8 contoh kegiatan yang harus dicocokkan dengan empat unsur, lalu jawaban dibahas bersama."],
      ["Output: hasil kuis 8 butir tiap peserta.", "Ukuran: minimal 80 persen peserta menjawab benar sedikitnya 6 dari 8 butir."]],
    ["**Topik 1.3** Contoh penerapan di kelas campuran (25 menit)",
      "Peserta mampu mengenali bagian contoh yang dapat diterapkan di kelasnya dan menuliskan tiga ide penerapan.",
      ["1. Contoh di SLB: satu topik untuk tiga kelompok kemampuan", "2. Contoh di sekolah inklusif: satu kelas reguler dengan murid berkebutuhan beragam", "3. Satu topik yang sama dengan ragam konten, ragam proses, dan pilihan produk", "4. Kekeliruan yang sering terjadi dan cara menghindarinya"],
      ["Menit 0 sampai 12: pemateri menayangkan dua contoh yang sudah jadi.", "Menit 12 sampai 25: peserta berdiskusi berpasangan dengan teman sebelah dan menjawab dua pertanyaan di lembar catatan: bagian mana yang bisa dipakai di kelas saya, dan apa yang perlu diubah."],
      ["Output: tiga ide penerapan pada lembar catatan tiap peserta.", "Ukuran: setiap peserta menuliskan tiga ide."]],
  ] },
  { judul: "Sesi 2. Membaca Profil Belajar Murid", waktu: "09.55 - 10.40 WIB, 45 menit (20 menit materi, 25 menit praktik)", topik: [
    ["**Topik 2.1** Asesmen diagnostik sederhana (20 menit, materi)",
      "Peserta mampu menjelaskan tujuan asesmen diagnostik dan menyebutkan tiga cara sederhana untuk melakukannya.",
      ["1. Tujuan asesmen diagnostik di awal pembelajaran", "2. Tiga cara: observasi terstruktur lima menit, percakapan singkat dengan murid, informasi dari orang tua atau guru sebelumnya", "3. Lima pertanyaan kunci: apa yang sudah dapat dilakukan murid, apa yang membantunya, apa yang menyulitkannya, apa minatnya, cara berkomunikasi yang paling nyaman baginya", "4. Etika: menjaga kerahasiaan data murid"],
      ["Pemateri memaparkan 15 menit dengan contoh lembar profil yang sudah terisi, lalu 5 menit tanya jawab. Lembar profil kosong dibagikan di akhir paparan."],
      ["Output: lembar profil kosong dan pemahaman cara mengisinya.", "Ukuran: butir terkait pada post-test dijawab benar oleh minimal 80 persen peserta."]],
    ["**Topik 2.2** Mengisi lembar profil satu murid (25 menit, praktik)",
      "Peserta mampu menyusun lembar profil belajar untuk satu murid dari kelasnya.",
      ["1. Bagian lembar profil: kekuatan, kebutuhan, minat, strategi yang membantu, target empat minggu", "2. Cara menulis kalimat yang spesifik dan dapat diamati"],
      ["Menit 0 sampai 15: tiap peserta memilih satu murid dari kelasnya dan mengisi lembar profil tanpa menulis nama murid.", "Menit 15 sampai 25: peserta bertukar lembar dengan teman sebelah dan menuliskan satu masukan pada lembar tersebut."],
      ["Output: satu lembar profil murid yang terisi lengkap.", "Ukuran: minimal 80 persen peserta mengisi seluruh bagian lembar."]],
  ] },
  { judul: "Sesi 3. Workshop Menyusun Modul Ajar Berdiferensiasi", waktu: "10.40 - 12.10 WIB, 90 menit, praktik", topik: [
    ["**Topik 3.1** Menentukan tujuan dan murid sasaran (20 menit)",
      "Peserta mampu menulis satu tujuan pembelajaran yang dapat diamati dan menentukan murid sasaran diferensiasi.",
      ["1. Memilih satu capaian pembelajaran dari mata pelajaran yang diampu", "2. Menulis tujuan dengan kata kerja yang dapat diamati, misalnya menyebutkan, menjelaskan, menghitung, membuat", "3. Memilih dua atau tiga murid sasaran dari lembar profil", "4. Daftar periksa tujuan: spesifik, dapat diamati, sesuai fase murid"],
      ["Peserta duduk dalam 8 kelompok berisi 10 guru, tiap kelompok didampingi satu pendamping dari tim Sanggabiz. Tiap guru mengisi bagian Identitas dan Tujuan pada template. Satu teman sekelompok memeriksa tujuan dengan daftar periksa dan menandai butir yang sudah terpenuhi."],
      ["Output: bagian Identitas dan Tujuan pada template terisi.", "Ukuran: seluruh peserta memiliki satu tujuan yang memenuhi daftar periksa."]],
    ["**Topik 3.2** Merancang ragam konten dan ragam proses (35 menit)",
      "Peserta mampu merancang dua ragam konten dan dua ragam proses belajar untuk satu topik.",
      ["1. Dua ragam konten, misalnya teks bergambar, teks sederhana, benda nyata", "2. Dua ragam proses, misalnya kerja mandiri bertahap, kerja berpasangan, bimbingan langsung guru", "3. Alat bantu yang sudah tersedia di sekolah", "4. Pemeriksaan: setiap ragam menuju tujuan yang sama"],
      ["Peserta mengisi bagian Diferensiasi Konten dan Diferensiasi Proses pada template. Pendamping berkeliling, membaca isian, dan mengajukan satu pertanyaan pemandu kepada tiap guru. Pemateri berkeliling dan memberi contoh tambahan bagi kelompok yang membutuhkan."],
      ["Output: bagian konten dan proses pada template terisi.", "Ukuran: setiap peserta memiliki dua ragam konten dan dua ragam proses."]],
    ["**Topik 3.3** Merancang pilihan produk, lingkungan belajar, dan asesmen (35 menit)",
      "Peserta mampu menyusun pilihan produk belajar, penataan lingkungan belajar, dan asesmen untuk melengkapi modul ajar.",
      ["1. Tiga pilihan produk: lisan, tulisan, karya", "2. Penataan kelas dan dukungan teman sebaya", "3. Akomodasi yang layak untuk murid tertentu", "4. Rubrik tiga tingkat untuk menilai produk"],
      ["Peserta melengkapi bagian Produk, Lingkungan Belajar, Asesmen, dan Akomodasi pada template, lalu memeriksa kelengkapan dengan daftar periksa sepuluh butir."],
      ["Output: draf modul ajar berdiferensiasi yang lengkap pada template.", "Ukuran: minimal 80 persen peserta menghasilkan draf lengkap, dinilai dengan rubrik empat butir (kesesuaian profil murid, ragam konten, ragam proses, pilihan produk)."]],
  ] },
  { judul: "Sesi 4. Presentasi Mini dan Umpan Balik", waktu: "12.10 - 12.30 WIB, 20 menit", topik: [
    ["**Topik 4.1** Gallery walk (12 menit)",
      "Peserta mampu memberi umpan balik tertulis pada hasil kerja kelompok lain.",
      ["1. Cara memberi umpan balik: satu hal yang sudah baik dan satu saran"],
      ["Tiap kelompok memajang satu draf modul di dinding. Peserta berkeliling selama 12 menit dan menempelkan dua catatan pada tiap draf yang dikunjungi: satu hal yang sudah baik dan satu saran."],
      ["Output: catatan umpan balik pada draf modul tiap kelompok.", "Ukuran: setiap draf yang dipajang menerima minimal 5 catatan."]],
    ["**Topik 4.2** Umpan balik pemateri (8 menit)",
      "Peserta mampu menyebutkan perbaikan prioritas untuk draf modulnya.",
      ["1. Tiga temuan umum dari seluruh draf berdasarkan rubrik", "2. Satu contoh perbaikan"],
      ["Pemateri membahas tiga temuan umum dan menunjukkan satu contoh perbaikan di layar. Peserta mencatat perbaikan yang berlaku untuk drafnya."],
      ["Output: daftar perbaikan prioritas pada lembar catatan peserta.", "Ukuran: setiap peserta mencatat minimal satu perbaikan."]],
  ] },
  { judul: "Sesi 5. Rencana Tindak Lanjut dan Evaluasi", waktu: "12.30 - 12.50 WIB, 20 menit", topik: [
    ["**Topik 5.1** Rencana tindak lanjut dua minggu (10 menit)",
      "Peserta mampu menyusun rencana menerapkan modulnya di kelas dalam dua minggu.",
      ["1. Lima pertanyaan rencana: kapan, di kelas mana, untuk murid siapa, apa yang diuji, siapa yang membantu"],
      ["Peserta menulis rencana pada lembar satu halaman, lalu bergabung ke grup tindak lanjut lewat kode QR di layar."],
      ["Output: rencana tindak lanjut satu halaman dan keanggotaan grup tindak lanjut.", "Ukuran: minimal 75 persen peserta bergabung di grup sebelum acara berakhir."]],
    ["**Topik 5.2** Post-test dan kuesioner (10 menit)",
      "Peserta mampu menunjukkan peningkatan pemahaman pada post-test.",
      ["1. Post-test 10 butir dengan soal yang sama seperti pre-test", "2. Kuesioner kepuasan 6 pertanyaan"],
      ["Peserta mengisi post-test dan kuesioner lewat ponsel melalui kode QR di layar. Pendamping membantu peserta yang membutuhkan."],
      ["Output: data post-test dan kuesioner dari seluruh peserta.", "Ukuran: rata-rata skor post-test naik minimal 20 persen dari pre-test."]],
  ] },
];
function tabelSesi(s) {
  const widths = [1350, 1650, 2300, 2688, 1650];
  return [
    h3(s.judul),
    p(s.waktu, { after: 70, keepNext: true, run: { color: MUTED, size: 17 } }),
    table(widths, ["Topik dan durasi", "Tujuan pembelajaran", "Outline materi", "Aktivitas", "Output dan ukuran keberhasilan"], s.topik, { size: 15, split: true }),
  ];
}

// ---------- isi ----------
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
  p("Sanggabiz menyampaikan penawaran workshop setengah hari untuk 80 guru SLB dan sekolah inklusif di Balikpapan. Workshop ini berfokus pada satu kemampuan, yaitu menyusun modul ajar berdiferensiasi yang dapat langsung dipakai di kelas. Setiap peserta pulang membawa template dan modul ringkas, serta satu draf modul ajar yang sudah mulai disusun dalam sesi."),
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
  p("Kami memilih satu kemampuan agar peserta pulang dengan hasil yang nyata. Ada tiga alasan."),
  bullet("Setiap guru di kelas inklusif mengajar murid dengan kebutuhan berbeda dalam satu ruangan. Menyesuaikan konten, proses, dan produk belajar adalah pekerjaan harian mereka."),
  bullet("Hasilnya dapat dibuat dalam setengah hari. Peserta menyusun satu modul ajar memakai template, lalu membawanya pulang."),
  bullet("Materinya berlaku untuk semua jenis kekhususan, sehingga sesuai untuk peserta campuran dari SLB dan sekolah inklusif."),
  p("Data khusus Balikpapan tentang tingkat pemahaman guru terhadap pembelajaran berdiferensiasi belum kami temukan. Karena itu workshop dibuka dengan pre-test. Hasilnya mencatat posisi awal peserta dan menjadi pembanding untuk hasil akhir.", { before: 80 }),
  h2("1.3 Tujuan"),
  p("**Tujuan umum.** Meningkatkan kemampuan 80 guru SLB dan sekolah inklusif di Balikpapan dalam merancang pembelajaran yang sesuai dengan kebutuhan beragam muridnya."),
  p("**Tujuan khusus.** Setelah workshop, peserta mampu:", { after: 60 }),
  bullet("menjelaskan empat unsur pembelajaran berdiferensiasi, yaitu konten, proses, produk, dan lingkungan belajar,"),
  bullet("membaca profil belajar seorang murid dengan asesmen diagnostik sederhana,"),
  bullet("menyusun modul ajar berdiferensiasi memakai template yang disediakan,"),
  bullet("menyusun rencana penerapan untuk dua minggu pertama di kelas masing-masing."),
  h2("1.4 Indikator Keberhasilan"),
  table([2200, 3700, 3738], ["Indikator", "Target", "Cara mengukur"], [
    ["Kehadiran", "80 guru hadir", "Daftar hadir"],
    ["Kenaikan pemahaman", "Rata-rata skor post-test naik minimal 20 persen dari pre-test", "Pre-test dan post-test, 10 soal yang sama"],
    ["Produk", "Minimal 80 persen peserta menghasilkan draf modul ajar lengkap", "Rubrik 4 butir: kesesuaian profil murid, ragam konten, ragam proses, pilihan produk"],
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
  p("Sanggabiz belum memiliki program khusus pendidikan luar biasa. Karena itu isi materi dipegang pemateri lokal yang berpengalaman di bidang pendidikan inklusif. Kami mengerjakan bagian yang kami kuasai: perekrutan peserta, logistik, penyusunan modul, evaluasi, dan pelaporan.", { before: 120 }),
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
  p("Workshop diadakan di Quest Hotel Balikpapan, di Jalan Jenderal Sudirman yang berada di kawasan tengah kota, sehingga mudah dijangkau dari SLB dan sekolah di berbagai kecamatan."),
  table([2600, 7038], ["Hal", "Keterangan"], [
    ["Hotel", "Quest Hotel Balikpapan, Jalan Jenderal Sudirman No. 11, Balikpapan"],
    ["Ruang", "Orchid Room, luas sekitar 120 m2"],
    ["Kapasitas", "28 sampai 100 orang, menyesuaikan susunan. Susunan kami: 80 peserta dalam 8 kelompok"],
    ["Fasilitas", "Proyektor, tata suara, akses internet ruang rapat, lahan parkir hotel"],
    ["Paket", "Meeting setengah hari: ruang, coffee break, makan siang. Harga acuan Rp 148.500 per orang"],
    ["Akses disabilitas", "[ISI TIM: konfirmasi lift atau jalur kursi roda dan toilet ke pihak hotel]"],
  ]),
  p("[ISI TIM: tempel foto Orchid Room dari pihak hotel di sini, cantumkan sumber foto]", { before: 140, align: AlignmentType.CENTER }),
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
  bullet("surat undangan dari dinas pendidikan melalui pemberi dana,"),
  bullet("kelompok kerja guru dan musyawarah kerja kepala sekolah (KKG dan MKKS),"),
  bullet("kepala SLB negeri dan swasta di Balikpapan,"),
  bullet("jaringan mitra lokal dan kampus di Kalimantan Timur."),
  table([2200, 7438], ["Waktu", "Kegiatan rekrutmen"], [
    ["H-21", "Undangan dikirim ke dinas, KKG, MKKS, dan kepala sekolah. Formulir pendaftaran daring dibuka"],
    ["H-14", "Pendaftaran ditutup bila kuota penuh. Formulir mencatat jenjang, mata pelajaran, dan kendala utama peserta"],
    ["H-7", "Konfirmasi kehadiran lewat WhatsApp. Peserta cadangan dihubungi bila ada yang berhalangan"],
    ["H-3", "Pengingat berisi lokasi, jam, dan petunjuk akses"],
    ["Hari-H", "Registrasi dan daftar hadir"],
  ], { align: [AlignmentType.CENTER, AlignmentType.LEFT] }),
];

const bab3 = [
  h1("3. Silabus Materi", true),
  h2("3.1 Gambaran Program"),
  table([2400, 7238], ["Hal", "Keterangan"], [
    ["Fokus", "Menyusun modul ajar berdiferensiasi untuk kelas dengan murid berkebutuhan beragam"],
    ["Peserta", "80 guru SLB dan sekolah inklusif"],
    ["Format", "Tatap muka, 5 jam acara inti. Materi sekitar 40 persen dan praktik sekitar 60 persen"],
    ["Pengajar", "1 pemateri lokal"],
    ["Hasil", "Template, modul ringkas, draf modul ajar sendiri, rencana tindak lanjut, sertifikat"],
  ]),
  h2("3.2 Pendekatan Pelatihan"),
  p("Program disusun dengan metode yang lazim dipakai dalam pelatihan dan dipilih sesuai kebutuhan guru sebagai peserta dewasa yang akan menghasilkan produk."),
  table([2300, 4938, 2400], ["Metode", "Penerapan di workshop ini", "Terlihat pada"], [
    [["**ADDIE**", "(Analysis, Design, Development, Implementation, Evaluation) [11]"], ["**Analysis:** formulir pendaftaran mencatat jenjang, mata pelajaran, dan kendala utama. Pre-test mencatat posisi awal.", "**Design:** tujuan, silabus, dan rundown.", "**Development:** template, modul ringkas, bahan ajar interaktif, instrumen tes.", "**Implementation:** workshop setengah hari.", "**Evaluation:** post-test, rubrik produk, kuesioner, tindak lanjut empat minggu."], "Bagian 2.3, 3.3, 3.4, 3.5, dan 4"],
    [["**Andragogi**", "(Knowles) [12]"], "Peserta belajar dari pengalaman sendiri. Setiap sesi memakai kasus murid dari kelas peserta. Pemateri berperan sebagai fasilitator, dan porsi praktik lebih besar dari materi (sekitar 60 persen).", "Aktivitas di silabus"],
    [["**Taksonomi Bloom revisi**", "(Anderson dan Krathwohl) [13]"], "Tujuan setiap topik memakai kata kerja yang dapat diamati: menyebutkan dan menjelaskan (memahami), menyusun dan merancang (menciptakan).", "Tujuan pembelajaran di silabus"],
    [["**Kirkpatrick**", "(empat level) [14]"], "Level 1 reaksi lewat kuesioner. Level 2 pembelajaran lewat pre-test, post-test, dan rubrik produk. Level 3 perilaku lewat laporan penerapan di grup tindak lanjut.", "Bagian 3.5"],
  ]),
  h2("3.3 Silabus per Sesi"),
  p("Silabus terdiri dari lima sesi dan sebelas topik. Setiap topik memuat tujuan pembelajaran, outline materi, aktivitas, output, dan ukuran keberhasilan."),
  table([900, 4338, 1400, 3000], ["Sesi", "Topik", "Durasi", "Hasil utama"], [
    ["1", "Ragam kebutuhan belajar dan prinsip pembelajaran berdiferensiasi", "80 menit", "Kartu kebutuhan murid, hasil kuis, tiga ide penerapan"],
    ["2", "Membaca profil belajar murid", "45 menit", "Lembar profil satu murid"],
    ["3", "Workshop menyusun modul ajar berdiferensiasi", "90 menit", "Draf modul ajar lengkap"],
    ["4", "Presentasi mini dan umpan balik", "20 menit", "Daftar perbaikan prioritas"],
    ["5", "Rencana tindak lanjut dan evaluasi", "20 menit", "Rencana dua minggu, data post-test"],
  ], { align: [AlignmentType.CENTER, AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.LEFT] }),
  ...SESI.flatMap(tabelSesi),
  h2("3.4 Hasil yang Dibawa Pulang"),
  p("Setiap hasil di bawah ini dihasilkan pada topik yang disebutkan, sehingga peserta dapat menelusuri asalnya."),
  table([3400, 2500, 3738], ["Hasil", "Dihasilkan pada", "Kegunaan"], [
    ["Kartu kebutuhan murid dan catatan pribadi", "Topik 1.1", "Memetakan kebutuhan belajar murid di kelas sendiri"],
    ["Tiga ide penerapan", "Topik 1.3", "Bahan awal merancang modul"],
    ["Lembar profil satu murid", "Topik 2.2", "Dasar menyusun modul untuk murid sasaran"],
    ["Draf modul ajar berdiferensiasi yang lengkap", "Topik 3.1 sampai 3.3", "Dipakai langsung di kelas"],
    ["Daftar perbaikan prioritas", "Topik 4.2", "Menyempurnakan modul sebelum dipakai"],
    ["Rencana tindak lanjut dua minggu", "Topik 5.1", "Panduan menerapkan modul di kelas"],
    ["Template modul ajar berdiferensiasi, cetak dan digital", "Dibagikan saat registrasi", "Dipakai untuk menyusun modul berikutnya"],
    ["Modul ringkas sekitar 20 halaman", "Dibagikan saat registrasi", "Rujukan setelah workshop (isi di Lampiran B)"],
    ["Sertifikat peserta", "Penutupan", "Bukti keikutsertaan"],
  ]),
  h2("3.5 Evaluasi dan Tindak Lanjut"),
  table([1500, 2400, 3738, 2000], ["Level", "Yang diukur", "Alat", "Waktu"], [
    ["1. Reaksi", "Kepuasan peserta", "Kuesioner 6 pertanyaan", "Akhir acara"],
    ["2. Pembelajaran", "Pengetahuan dan keterampilan", "Pre-test dan post-test 10 soal yang sama, rubrik produk 4 butir", "Awal dan akhir acara"],
    ["3. Perilaku", "Penerapan di kelas", "Laporan peserta di grup tindak lanjut (foto atau cerita singkat)", "Minggu 1 sampai 4"],
  ]),
  p("Dampak pada murid (level 4) berada di luar cakupan program setengah hari ini dan dapat dirancang pada tahap lanjutan bila pemberi dana menghendaki.", { before: 100 }),
  p("Grup WhatsApp tindak lanjut berjalan empat minggu dan dimoderatori tim Sanggabiz untuk tanya jawab dan berbagi hasil penerapan. Laporan akhir kami serahkan dalam 7 hari kerja setelah acara, berisi daftar hadir, hasil tes, penilaian produk, hasil kuesioner, dan dokumentasi."),
];

const bab4 = [
  h1("4. Rundown Acara", true),
  p("Jam di bawah mengikuti usulan awal. Bila tanggal dan tempat berubah, jam dapat digeser tanpa mengubah durasi tiap sesi."),
  table([1700, 1100, 4700, 2138], ["Waktu (WIB)", "Durasi", "Kegiatan", "Pemandu"], [
    ["07.30 - 08.00", "30 menit", "Registrasi, daftar hadir, pembagian perlengkapan", "Panitia"],
    ["08.00 - 08.20", "20 menit", "Pembukaan dan sambutan (10 menit), pre-test lewat ponsel (10 menit)", "MC dan Sanggabiz"],
    ["08.20 - 09.40", "80 menit", "**Sesi 1.** Ragam kebutuhan belajar dan prinsip pembelajaran berdiferensiasi", "Pemateri"],
    ["09.40 - 09.55", "15 menit", "Coffee break", "Panitia"],
    ["09.55 - 10.40", "45 menit", "**Sesi 2.** Membaca profil belajar murid", "Pemateri"],
    ["10.40 - 12.10", "90 menit", "**Sesi 3.** Workshop menyusun modul ajar berdiferensiasi", "Pemateri dan pendamping"],
    ["12.10 - 12.30", "20 menit", "**Sesi 4.** Presentasi mini dan umpan balik", "Pemateri"],
    ["12.30 - 12.50", "20 menit", "**Sesi 5.** Rencana tindak lanjut, post-test, kuesioner", "Sanggabiz"],
    ["12.50 - 13.00", "10 menit", "Penutupan dan pembagian sertifikat", "MC"],
    ["13.00 - 13.45", "45 menit", "Makan siang dan ramah tamah", "Panitia"],
  ], { align: [AlignmentType.CENTER, AlignmentType.CENTER] }),
  p("Perbandingan waktu belajar: materi 100 menit (sekitar 40 persen) dan praktik 155 menit (sekitar 60 persen) dari 255 menit.", { before: 140 }),
];

const bab5 = [
  h1("5. Anggaran", true),
  h2("5.1 Dasar Penyusunan"),
  bullet("Peserta 80 orang. Konsumsi dan perlengkapan dihitung untuk 90 orang, yaitu peserta, pemateri, pendamping, dan panitia."),
  bullet("Harga tempat mengacu pada paket meeting setengah hari di Quest Hotel Balikpapan sebesar Rp 148.500 per orang [15]. Harga final mengikuti penawaran tertulis hotel."),
  bullet("Honor pemateri berada di bawah batas Rp 1.700.000 per jam yang lazim dipakai untuk narasumber pakar atau praktisi dalam standar biaya pemerintah daerah. Pemateri berasal dari Balikpapan atau Kaltim, sehingga tidak ada biaya tiket dan penginapan."),
  bullet("Seluruh harga sudah termasuk pajak."),
  h2("5.2 Rincian Anggaran"),
  tabelAnggaran(),
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
    ["Identitas", "Mata pelajaran, fase atau kelas, alokasi waktu, nama guru"],
    ["Capaian dan tujuan pembelajaran", ""],
    ["Profil belajar murid", "Kesiapan, minat, kebutuhan khusus, hal yang membantu murid belajar"],
    ["Diferensiasi konten", "Ragam A: .......   Ragam B: ......."],
    ["Diferensiasi proses", "Ragam A: .......   Ragam B: ......."],
    ["Diferensiasi produk", "Pilihan produk yang boleh dipilih murid"],
    ["Lingkungan belajar", "Penataan kelas, alat bantu, dukungan teman sebaya"],
    ["Asesmen", "Asesmen diagnostik dan asesmen formatif"],
    ["Akomodasi yang layak", "Penyesuaian yang disiapkan untuk murid tertentu"],
    ["Refleksi", "Apa yang berhasil dan apa yang diubah pada pertemuan berikutnya"],
  ]),
  h2("Lampiran B. Isi Modul Ringkas"),
  bullet("Prinsip pembelajaran berdiferensiasi dan empat unsurnya."),
  bullet("Ragam kebutuhan belajar murid dan cara mengenalinya."),
  bullet("Asesmen diagnostik sederhana beserta lembar profil murid."),
  bullet("Langkah menyusun modul ajar berdiferensiasi."),
  bullet("Dua contoh modul ajar yang sudah terisi."),
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
    "XWORK, Paket Meeting di Kota Balikpapan (Quest Hotel Balikpapan, Orchid Room). https://xwork.co/id/paket-meeting/kota-balikpapan-273. Data ruang rapat: AMI Magazine, https://www.amimagazine.global/Meeting-Event-Venues/Balikpapan-Indonesia/Convention-Hotel/Quest-Hotel-Balikpapan-p57912634",
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
      ["Dibawa pulang", "Template dan modul ringkas, draf modul ajar sendiri"],
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
const out = path.join(outDir, "v2 Proposal Workshop Pembelajaran Berdiferensiasi - SLB dan Inklusif Balikpapan.docx");
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log("ditulis:", out);
  console.log("total", fmt(TOTAL), "per peserta", fmt(Math.round(TOTAL / 80 / 1000) * 1000));
});
