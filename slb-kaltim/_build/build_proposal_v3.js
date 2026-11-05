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

// ---------- v3: sudut pandang event organizer (EO). Materi, modul, dan narasumber di luar lingkup ----------
const WAKTU = "WITA";

// ---------- sekolah sasaran (sumber: One Stop SLB Kaltim, daftar Pertamina Foundation) ----------
// [nama, kecamatan, murid penerima, kuota guru]
const SLB = [
  ["SLB Dharma Kencana Balikpapan", "Balikpapan Kota", 6, 9],
  ["SLB Tunas Bangsa Balikpapan", "Balikpapan Selatan", 5, 7],
  ["SLB Negeri Balikpapan", "Balikpapan Selatan", 4, 5],
];
const SD = [
  ["SDN 02 Balikpapan Utara", "Balikpapan Utara", 4, 5],
  ["SDN 002 Balikpapan Barat", "Balikpapan Barat", 3, 4],
  ["SDN 013 Balikpapan Kota", "Balikpapan Kota", 3, 4],
  ["SDN 006 Balikpapan Selatan", "Balikpapan Selatan", 2, 3],
  ["SDN 008 Balikpapan Selatan", "Balikpapan Selatan", 2, 3],
  ["SDN 015 Balikpapan Kota", "Balikpapan Kota", 2, 3],
  ["SDN 024 Balikpapan Utara", "Balikpapan Utara", 2, 3],
  ["SD Negeri 026 Balikpapan Tengah", "Balikpapan Tengah", 1, 2],
  ["SDN 005 Balikpapan Timur", "Balikpapan Timur", 1, 2],
  ["SDN 001 Balikpapan Barat", "Balikpapan Barat", 1, 2],
  ["SDN 011 Balikpapan Kota", "Balikpapan Kota", 1, 2],
  ["SDN 005 Balikpapan Barat", "Balikpapan Barat", 1, 2],
  ["SDN 009 Balikpapan Timur", "Balikpapan Timur", 1, 2],
  ["SDN 010 Balikpapan Kota", "Balikpapan Kota", 1, 2],
  ["SD-IT As'adiyah Manuntung Balikpapan", "Balikpapan Barat", 1, 2],
  ["SDN 019 Balikpapan Tengah", "Balikpapan Tengah", 1, 2],
  ["SDN 009 Balikpapan Utara [ISI TIM: di daftar tertulis SDB 009, cek nama ke sekolah]", "Balikpapan Utara", 1, 2],
  ["SDN 034 Balikpapan Utara", "Balikpapan Utara", 1, 2],
  ["SDN 016 Balikpapan Utara", "Balikpapan Utara", 1, 2],
  ["SDN 005 Balikpapan Tengah", "Balikpapan Tengah", 1, 2],
  ["SDN 014 Balikpapan Utara", "Balikpapan Utara", 1, 2],
  ["SDN 030 Balikpapan Tengah", "Balikpapan Tengah", 1, 2],
  ["SDN 015 Balikpapan Utara", "Balikpapan Utara", 1, 2],
  ["SD GPIB Kristen Balikpapan", "Balikpapan Barat", 1, 2],
];
const sum = (a, i) => a.reduce((t, r) => t + r[i], 0);
if (SLB.length + SD.length !== 27 || sum(SLB, 3) + sum(SD, 3) !== 80 || sum(SLB, 2) + sum(SD, 2) !== 50) throw new Error("data sekolah tidak cocok");
function tabelSekolah() {
  const widths = [600, 4038, 2200, 1400, 1400];
  const rows = [];
  let no = 0;
  const grup = (judul, arr) => {
    rows.push([{ group: `${judul}. ${arr.length} sekolah, kuota ${sum(arr, 3)} guru` }]);
    arr.forEach((r) => rows.push([String(++no), r[0], r[1], String(r[2]), String(r[3])]));
  };
  grup("A. SLB", SLB);
  grup("B. SD", SD);
  rows.push([{ text: "", bg: "F2F5FA" }, { text: "Jumlah 27 sekolah", bold: true, bg: "F2F5FA" }, { text: "", bg: "F2F5FA" }, { text: String(sum(SLB, 2) + sum(SD, 2)), bold: true, bg: "F2F5FA" }, { text: String(sum(SLB, 3) + sum(SD, 3)), bold: true, bg: "F2F5FA" }]);
  return table(widths, ["No", "Asal sekolah", "Kecamatan", "Murid penerima", "Kuota guru"], rows,
    { size: 16, align: [AlignmentType.CENTER, AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER] });
}

// ---------- anggaran (satu opsi, harga akhir sudah termasuk pajak) ----------
// [kode, uraian, volume, satuan volume, frekuensi, satuan frekuensi, harga satuan]
const ANGGARAN = [
  { group: "A", judul: "Program Implementation Personnel" },
  ["A.1", "Project Manager", 1, "orang", 1, "bulan", 5000000],
  ["A.2", "Project Local", 1, "orang", 1, "bulan", 3000000],
  ["A.3", "MC dan moderator", 2, "orang", 1, "hari", 500000],
  ["A.4", "Dokumentasi: update real time dan video, 5 post reels hasil editing untuk TikTok dan Instagram", 1, "paket", 1, "hari", 3000000],
  { group: "B", judul: "Materi dan Equipment" },
  ["B.1", "Perlengkapan peserta: ID card, pulpen, notebook", 80, "orang", 1, "hari", 80000],
  ["B.2", "Snack", 90, "orang", 1, "hari", 20000],
  ["B.3", "Makan siang", 90, "orang", 1, "hari", 35000],
  ["B.4", "Meja roundtable", 9, "meja", 1, "hari", 200000],
  ["B.5", "Kursi", 90, "kursi", 1, "hari", 20000],
  ["B.6", "Infocus", 2, "unit", 1, "hari", 500000],
  ["B.7", "Media cetak: banner, poster, signage", 1, "paket", 1, "hari", 500000],
  { group: "C", judul: "Transport dan Advokasi" },
  ["C.1", "Transport peserta", 80, "orang", 1, "hari", 100000],
  ["C.2", "Advokasi sekolah", 27, "sekolah", 1, "kali", 300000],
];
let TOTAL = 0;
const anggaranRows = [];
{
  let sub = 0, judul = "", kode = "";
  const tutup = () => { if (kode) anggaranRows.push([{ text: "", bg: "EEF3F9" }, { text: `Subtotal ${kode}`, bold: true, bg: "EEF3F9" }, { text: "", bg: "EEF3F9" }, { text: "", bg: "EEF3F9" }, { text: "", bg: "EEF3F9" }, { text: fmt(sub), bold: true, bg: "EEF3F9" }]); TOTAL += sub; sub = 0; };
  ANGGARAN.forEach((r) => {
    if (r.group) { tutup(); kode = r.group; judul = r.judul; anggaranRows.push([{ group: `${r.group}. ${r.judul}` }]); return; }
    const jumlah = r[2] * r[4] * r[6]; sub += jumlah;
    anggaranRows.push([r[0], r[1], `${r[2]} ${r[3]}`, `${r[4]} ${r[5]}`, fmt(r[6]), fmt(jumlah)]);
  });
  tutup();
}
const PER_PESERTA = Math.round(TOTAL / 80 / 1000) * 1000;
function tabelAnggaran() {
  const widths = [800, 3338, 1250, 1100, 1550, 1600];
  const sumRow = (label, val) => [{ text: "", bg: "F2F5FA" }, { text: label, bold: true, bg: "F2F5FA" }, { text: "", bg: "F2F5FA" }, { text: "", bg: "F2F5FA" }, { text: "", bg: "F2F5FA" }, { text: val, bold: true, bg: "F2F5FA" }];
  return table(widths, ["Code", "Description", "Quantity", "Frequency", "Price (IDR)", "Total (IDR)"], [
    ...anggaranRows,
    sumRow("TOTAL (sudah termasuk pajak)", fmt(TOTAL)),
    sumRow("Per peserta (80 orang, dibulatkan)", fmt(PER_PESERTA)),
  ], { align: [AlignmentType.CENTER, AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.RIGHT, AlignmentType.RIGHT] });
}

// ---------- isi ----------
const surat = [
  h1("Surat Pengantar"),
  p("Nomor: [ISI TIM: nomor surat]", { after: 20 }),
  p("Lampiran: 1 berkas proposal", { after: 20 }),
  p("Perihal: **Penawaran Penyelenggaraan Workshop Pembelajaran Berdiferensiasi untuk Guru SLB dan SD di Balikpapan**", { after: 240 }),
  p("Kepada Yth.", { after: 20 }),
  p("[ISI TIM: nama dan jabatan penerima]", { after: 20 }),
  p("[ISI TIM: instansi penerima]", { after: 20 }),
  p("di Balikpapan", { after: 240 }),
  p("Dengan hormat,"),
  p("Sanggabiz menyampaikan penawaran sebagai event organizer untuk workshop setengah hari bagi 80 guru SLB dan SD di Balikpapan. Kami mengerjakan pencarian dan konfirmasi peserta dari 27 sekolah sasaran, advokasi ke sekolah, penyediaan tempat dan konsumsi, pelaksanaan di hari acara, serta laporan. Materi, modul, dan narasumber disiapkan pihak lain dan tidak termasuk dalam penawaran ini."),
  p(`Acara direncanakan pada Rabu, 11 November 2026, siang hari. Registrasi mulai pukul 12.00 ${WAKTU} dan acara inti pukul 13.00 ${WAKTU}. Proposal ini memuat latar belakang, lingkup pekerjaan, tempat dan peserta, alur kerja, rundown, anggaran, dan susunan tim.`),
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
  p("Kurikulum Merdeka meminta guru merancang pembelajaran sesuai kesiapan, minat, dan profil belajar murid. Kajian di beberapa sekolah dasar inklusif mencatat kendala yang sama: pemahaman guru tentang kebutuhan belajar anak masih perlu diperkuat, pelatihan perlu ditambah, waktu persiapan ringkas, dan jumlah murid per kelas besar. Pembelajaran berdiferensiasi sudah mulai dijalankan, tetapi belum mencakup semua aspek, terutama konten dan lingkungan belajar."),
  p("Kebutuhan serupa tercatat di Kalimantan Timur. Komisi IV DPRD Kaltim menyampaikan bahwa guru berlatar pendidikan luar biasa perlu ditambah, dan mengusulkan pembukaan program studi PLB bersama Universitas Mulawarman [3]. SLB negeri berada di bawah Dinas Pendidikan dan Kebudayaan Provinsi, yang mengelola 11 SLB di Kaltim [4]. Salah satunya SLB Negeri Balikpapan, berdiri sejak 1984 dan melayani 401 siswa [5]. Sekolah inklusif di Balikpapan berada di bawah dua kewenangan. SMA dan SMK dikelola Pemerintah Provinsi melalui Cabang Dinas Pendidikan Wilayah I [6], sedangkan SD dan SMP dikelola Pemerintah Kota. Pemerintah Kota telah menerbitkan Peraturan Wali Kota Nomor 9 Tahun 2018 tentang penyelenggaraan pendidikan inklusif [7]."),
  centerImg("grafik-siswa-slbn-balikpapan.png", 520, 213),
  caption("Gambar 2. Jumlah siswa SLB Negeri Balikpapan per jenjang"),
  p("Peraturan Pemerintah Nomor 13 Tahun 2020 mewajibkan pemerintah pusat dan daerah memfasilitasi akomodasi yang layak bagi peserta didik penyandang disabilitas, termasuk menyiapkan pendidik dan menyediakan kurikulum [8]. Penguatan kemampuan guru adalah bagian dari kewajiban itu, dan workshop ini dirancang untuk mendukung pemenuhannya."),

  h2("1.2 Lingkup Pekerjaan Sanggabiz"),
  p("Workshop ini bertema pembelajaran berdiferensiasi bagi guru yang mendampingi murid berkebutuhan khusus. Sanggabiz bertindak sebagai event organizer. Pembagian pekerjaannya sebagai berikut."),
  table([4819, 4819], ["Dikerjakan Sanggabiz", "Disiapkan pihak lain, di luar penawaran ini"], [
    [["1. Pendataan 27 sekolah sasaran dan penetapan kuota 80 guru", "2. Komunikasi, pengiriman surat, dan advokasi ke sekolah", "3. Formulir pendaftaran, konfirmasi, dan pengingat peserta", "4. Tempat, perabot, perlengkapan, dan konsumsi", "5. Transport peserta", "6. Project manager dan MC di hari acara", "7. Registrasi, daftar hadir, kuesioner kepuasan, dan dokumentasi", "8. Laporan akhir"],
     ["1. Materi dan modul workshop", "2. Narasumber dan honornya", "3. Isi sesi dan pembagian waktu di dalam blok materi", "Sanggabiz menyediakan ruang, waktu, dan perlengkapan agar sesi narasumber berjalan, dan membagikan bahan yang diserahkan narasumber bila diperlukan."]],
  ]),
  h2("1.3 Tujuan Penyelenggaraan"),
  p("**Tujuan umum.** Menyelenggarakan workshop setengah hari yang menghadirkan 80 guru SLB dan SD dari 27 sekolah di Balikpapan, dengan tempat, konsumsi, dan layanan peserta yang tertata."),
  p("**Tujuan khusus.** Sanggabiz memastikan bahwa:", { after: 60 }),
  bullet("27 sekolah sasaran menerima undangan dan mengetahui kuota gurunya,"),
  bullet("80 guru peserta dan peserta cadangan terkonfirmasi sebelum hari acara,"),
  bullet(`tempat dan perlengkapan siap, sehingga registrasi dibuka pukul 12.00 ${WAKTU} dan acara inti dimulai pukul 13.00 ${WAKTU},`),
  bullet("data kehadiran, kepuasan, dan dokumentasi tersaji dalam laporan akhir."),
  h2("1.4 Indikator Keberhasilan"),
  table([2200, 3700, 3738], ["Indikator", "Target", "Cara mengukur"], [
    ["Kehadiran", "80 guru hadir", "Daftar hadir"],
    ["Sebaran sekolah", "27 sekolah terwakili, setiap sekolah mengirim minimal satu guru", "Daftar hadir dibandingkan tabel kuota"],
    ["Ketepatan waktu", `Registrasi dibuka 12.00 ${WAKTU} dan acara inti dimulai 13.00 ${WAKTU}`, "Catatan project manager"],
    ["Kepuasan", "Rata-rata minimal 4 dari 5 untuk layanan penyelenggaraan", "Kuesioner akhir"],
    ["Laporan", "Diserahkan dalam 7 hari kerja setelah acara", "Tanggal serah"],
  ]),
  h2("1.5 Tentang Sanggabiz"),
  p("Sanggabiz adalah mitra solusi bisnis yang menyediakan layanan department-as-a-service: keuangan, sumber daya manusia, layanan pelanggan, pemasaran, manajemen proyek, administrasi, dan operasional. Di bidang pelatihan, kami mengelola program dari perencanaan sampai pelaporan."),
  table([2400, 7238], ["Hal", "Keterangan"], [
    ["Skala", "400+ proyek dan layanan, 500+ pakar dan trainer, 6.000+ mitra engagement"],
    ["Pengalaman multikota", "Pelatihan MSME Rumah BUMN di lebih dari sembilan kota, termasuk Samarinda"],
    ["Klien", "Lembaga pendidikan, BUMN, yayasan, dan pemerintah daerah, antara lain Universitas Muhammadiyah Yogyakarta, Yakes Pertamina, dan Pemerintah Kota Yogyakarta"],
    ["Pendekatan kerja", "Alur kerja dari proses bisnis acara, dengan kerangka ADDIE (lihat bagian 3.1)"],
  ]),
  p("Sanggabiz belum memiliki program khusus pendidikan luar biasa. Karena itu kami membatasi diri pada bagian yang kami kuasai: pencarian peserta, advokasi, logistik, pelaksanaan acara, dan pelaporan. Isi materi ditangani pihak yang berpengalaman di bidang pendidikan inklusif.", { before: 120 }),
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
  p("Workshop diadakan di Kantor Pertamina di Balikpapan. Acara siang hari, dan tempatnya berada di kota yang dapat dijangkau dari SLB dan SD di berbagai kecamatan."),
  table([2600, 7038], ["Hal", "Keterangan"], [
    ["Tempat", "Kantor Pertamina, Balikpapan. [ISI TIM: nama gedung, alamat, dan nama ruang]"],
    ["Kapasitas", "90 kursi: 80 peserta, dan sisanya untuk narasumber, panitia, dan tamu"],
    ["Susunan", "Meja bundar, 8 meja untuk kelompok peserta dan 1 meja untuk registrasi"],
    ["Perlengkapan", "Meja, kursi, dan 2 infocus disewa dari vendor. Tata suara, listrik, dan parkir mengikuti fasilitas gedung. [ISI TIM: konfirmasi tata suara, listrik, parkir, dan Wi-Fi]"],
    ["Biaya tempat", "Disediakan tanpa biaya oleh Kantor Pertamina, sehingga anggaran tidak memuat sewa tempat"],
    ["Akses disabilitas", "[ISI TIM: konfirmasi jalur kursi roda dan toilet ke pengelola gedung]"],
  ]),
  centerImg("skema-susunan-ruang.png", 520, 292),
  caption("Gambar 4. Skema susunan ruang untuk 80 peserta (usulan, menyesuaikan ruang yang tersedia)"),
  h2("2.2 Waktu Pelaksanaan"),
  table([3000, 6638], ["Hal", "Keterangan"], [
    ["Tanggal", "Rabu, 11 November 2026 [ISI TIM: konfirmasi tanggal]"],
    ["Registrasi dan makan siang", `12.00 sampai 13.00 ${WAKTU}`],
    ["Acara inti", `13.00 sampai 17.30 ${WAKTU} (4,5 jam)`],
  ]),
  h2("2.3 Peserta dan Strategi Rekrutmen"),
  p("Peserta adalah 80 guru dari 27 sekolah di Balikpapan, terdiri dari 3 SLB dan 24 SD. Daftar sekolah berasal dari dua daftar penerima beasiswa Sobat Istimewa Pertamina Foundation untuk Kota Balikpapan, tahun ajaran 2024/2025 dan 2025/2026 [9]. Kedua daftar memuat 50 murid penyandang disabilitas dari 27 sekolah. Tuna daksa paling banyak dengan 29 murid, disusul tuna rungu 8, tuna grahita 7, tuna netra 3, tuna ganda 2, dan disabilitas fisik 1. Kecamatan sekolah mengacu pada data Sekolah Kita [10]."),
  p("Kuota guru tiap sekolah sebanding dengan jumlah murid penerima di sekolah itu. Kami menghitung porsi murid penerima terhadap 80, lalu membulatkannya agar jumlah seluruh kuota tepat 80. Peserta tidak membayar, karena biaya ditanggung pemberi dana."),
  tabelSekolah(),
  p("Rekrutmen berjalan di dua jalur sekaligus karena waktu ke hari acara ringkas. Jalur pertama langsung ke sekolah lewat telepon, WhatsApp, dan kunjungan petugas advokasi. Jalur kedua lewat surat dari Dinas Pendidikan dan Kebudayaan Provinsi Kalimantan Timur.", { before: 160 }),
  table([2400, 7238], ["Waktu", "Kegiatan rekrutmen"], [
    ["Jumat, 6 November", "Tim Sanggabiz mulai menghubungi 27 sekolah lewat telepon dan WhatsApp. Kami memperkenalkan acara, menyampaikan kuota guru sekolah itu, dan meminta nama serta kontak guru. Formulir pendaftaran daring dibuka"],
    ["Sabtu dan Minggu, 7 dan 8 November", "Komunikasi dilanjutkan sampai setiap sekolah mengetahui kuotanya. Nama guru yang bersedia hadir masuk ke daftar peserta"],
    ["Senin, 9 November", "Surat dari Dinas Pendidikan dan Kebudayaan Provinsi Kalimantan Timur diperkirakan terbit. Begitu surat terbit, personel kami mengirimkannya ke sekolah sambil melakukan advokasi langsung"],
    ["Selasa, 10 November", "Advokasi dan pengiriman surat berlanjut ke sekolah yang belum terjangkau. Konfirmasi akhir lewat WhatsApp, penetapan 80 peserta dan peserta cadangan, lalu pengingat berisi lokasi, jam, dan petunjuk akses. ID card dicetak dan disusun per sekolah"],
    ["Rabu, 11 November", `Hari acara. Registrasi dan daftar hadir lewat formulir daring mulai pukul 12.00 ${WAKTU}`],
  ]),
  h3("Risiko dan penanganannya"),
  table([3000, 6638], ["Risiko", "Penanganan"], [
    ["Surat dinas terbit setelah Senin, 9 November", "Komunikasi lewat telepon dan WhatsApp sudah berjalan sejak Jumat. Kami menyampaikan undangan awal lebih dulu, lalu mengirim surat begitu terbit"],
    ["Kuota sebuah sekolah tidak terpenuhi", "Kuota yang tidak terisi dialihkan ke guru lain dari 27 sekolah yang sama atau ke peserta cadangan"],
    ["SD berada di bawah kewenangan Pemerintah Kota, bukan provinsi", "Bila sekolah meminta surat dari dinas kota, kami meneruskan permintaan itu ke pemberi dana"],
  ]),
];

const bab3 = [
  h1("3. Alur Kerja dan Pelaksanaan", true),
  h2("3.1 Pendekatan Kerja"),
  p("Kami menyusun pekerjaan dari proses bisnis acara, dan anggaran mengikuti proses itu. Alurnya memakai kerangka ADDIE, yang lazim dipakai untuk menyelenggarakan program pelatihan dari analisis sampai evaluasi [11]."),
  table([2800, 6838], ["Tahap ADDIE", "Penerapan di acara ini"], [
    ["Analysis", "Pendataan 27 sekolah, penetapan kuota, dan formulir pendaftaran yang mencatat sekolah, jenjang, mata pelajaran, dan kebutuhan aksesibilitas peserta"],
    ["Design", "Rundown, susunan ruang, dan daftar kebutuhan tempat, perabot, dan konsumsi"],
    ["Development", "Pemesanan vendor, pencetakan ID card dan media cetak (banner, poster, signage), penyiapan perlengkapan peserta, naskah MC, dan formulir daring"],
    ["Implementation", "Komunikasi dan advokasi ke sekolah, konfirmasi peserta, lalu pelaksanaan di hari acara"],
    ["Evaluation", "Daftar hadir, kuesioner kepuasan, dokumentasi, dan laporan akhir"],
  ]),
  h2("3.2 Tahapan Kerja dan Baris Anggarannya"),
  p("Setiap tahap punya pekerjaan yang dapat ditelusuri ke baris anggaran pada bagian 5.", { after: 100 }),
  table([600, 2600, 4038, 2400], ["No", "Tahap", "Pekerjaan", "Baris anggaran"], [
    ["1", "Pendataan sekolah dan guru", "Daftar 27 sekolah, kuota, kontak guru, dan formulir pendaftaran", "Diserap di A.1 dan A.2"],
    ["2", "Komunikasi dan undangan", "Telepon dan WhatsApp, pengiriman surat, advokasi ke sekolah, konfirmasi, dan pengingat", "C.2, diserap di A.1 dan A.2"],
    ["3", "Teknis acara", "Perabot, infocus, konsumsi, perlengkapan peserta, media cetak, transport peserta, project manager, project local, dan MC", "A.1 sampai A.3, B.1 sampai B.7, C.1"],
    ["4", "Monitoring dan evaluasi", "Daftar hadir, kuesioner kepuasan, dan dokumentasi", "A.4, diserap di A.1 dan A.2"],
    ["5", "Laporan", "Laporan akhir dalam 7 hari kerja", "Diserap di A.1"],
  ], { align: [AlignmentType.CENTER, AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.LEFT] }),
  h2("3.3 Monitoring, Evaluasi, dan Laporan"),
  table([2600, 4838, 2200], ["Yang dipantau", "Cara", "Waktu"], [
    ["Kehadiran dan sebaran sekolah", "Daftar hadir lewat formulir daring, direkap per sekolah dan dibandingkan dengan tabel kuota", "Hari acara"],
    ["Kepuasan penyelenggaraan", "Kuesioner 4 pertanyaan lewat kode QR, kertas cadangan bagi peserta tanpa ponsel atau sinyal", "Akhir acara"],
    ["Hasil tes narasumber", "Bila narasumber menyiapkan pre-test dan post-test, kami menyediakan tautan formulirnya dan merekap hasilnya", "Awal dan akhir acara"],
    ["Dokumentasi", "Update real time, video, dan 5 post reels hasil editing untuk TikTok dan Instagram", "Hari acara"],
  ]),
  p("Laporan akhir kami serahkan dalam 7 hari kerja setelah acara. Isinya daftar hadir, rekap sebaran sekolah, hasil kuesioner, rekap tes bila ada, dan dokumentasi.", { before: 120 }),
];

const bab4 = [
  spacer(240),
  h1("4. Rundown Acara"),
  p("Jam di bawah adalah usulan awal. Bila tanggal atau tempat berubah, jam dapat digeser tanpa mengubah durasi tiap blok. Isi dan pembagian waktu di dalam blok materi ditentukan oleh narasumber."),
  table([1900, 1100, 4638, 2000], [`Waktu (${WAKTU})`, "Durasi", "Kegiatan", "Pemandu"], [
    ["12.00 - 13.00", "60 menit", "Registrasi, daftar hadir lewat formulir daring, pembagian ID card dan perlengkapan, makan siang, dan salat Dzuhur", "Project manager dan tim Sanggabiz"],
    ["13.00 - 13.20", "20 menit", "Pembukaan dan sambutan", "MC"],
    ["13.20 - 15.00", "100 menit", "Blok materi dan praktik 1", "Narasumber"],
    ["15.00 - 15.30", "30 menit", "Coffee break dan salat Ashar", "Project manager"],
    ["15.30 - 17.00", "90 menit", "Blok materi dan praktik 2", "Narasumber"],
    ["17.00 - 17.30", "30 menit", "Kuesioner kepuasan, penutupan, foto bersama, pembagian e-sertifikat, dan pembagian uang transport", "MC dan tim Sanggabiz"],
  ]),
  p("Blok materi berjumlah 190 menit. Bila narasumber memerlukan waktu lebih panjang, acara berakhir setelah pukul 17.30 dan kami menyesuaikan pengaturan pulang peserta.", { before: 120 }),
];

const bab5 = [
  h1("5. Anggaran", true),
  h2("5.1 Dasar Penyusunan"),
  bullet("Peserta 80 orang dari 27 sekolah. Perlengkapan peserta dan transport dihitung per peserta."),
  bullet("Konsumsi dan kursi dihitung 90: 80 peserta, ditambah narasumber, panitia, dan tamu."),
  bullet("Advokasi sekolah dihitung per sekolah untuk 27 sekolah."),
  bullet("Materi, modul, dan narasumber tidak termasuk, karena disiapkan pihak lain."),
  bullet("Tempat disediakan tanpa biaya oleh Kantor Pertamina. E-sertifikat dibuat tim administrasi dan diserap di A.1."),
  bullet("Seluruh harga sudah termasuk pajak."),
  h2("5.2 Rincian Anggaran"),
  tabelAnggaran(),
  h3("Catatan"),
  p("Anggaran di atas sudah termasuk pajak.", { after: 40 }),
  p("Anggaran bersifat estimatif dan masih dapat dinegosiasikan sesuai kesepakatan bersama."),
];

const bab6 = [
  h1("6. Susunan Tim Sanggabiz", true),
  p("Tim bekerja dari Balikpapan dan terhubung dengan kantor Sanggabiz untuk administrasi dan pelaporan."),
  table([2600, 5038, 2000], ["Peran", "Tugas", "Baris anggaran"], [
    ["Project manager", "Memimpin seluruh pekerjaan, mengatur vendor, memantau kuota peserta, dan memandu hari acara", "A.1"],
    ["Project local", "Mengoordinasikan pekerjaan di Balikpapan: tempat, vendor, dan petugas lapangan", "A.2"],
    ["Petugas advokasi sekolah", "Menghubungi dan mengunjungi 27 sekolah, menyerahkan surat, dan mencatat nama guru peserta", "C.2"],
    ["MC dan moderator", "Dua orang memandu jalannya acara sesuai rundown dan naskah", "A.3"],
    ["Dokumentasi", "Update real time, video, dan konten reels untuk TikTok dan Instagram", "A.4"],
    ["Tim administrasi dan logistik Sanggabiz", "Formulir pendaftaran, ID card, pemesanan vendor, rekap kehadiran, dan laporan", "Diserap di A.1"],
  ]),
  p("[ISI TIM: nama project manager, project local, petugas advokasi, MC, dan dokumentasi]", { before: 140 }),
];

const bab7 = [
  h1("7. Penutup"),
  p("Waktu menuju hari acara ringkas, dan karena itu pekerjaan sudah kami petakan per hari sejak komunikasi pertama ke sekolah. Kami berharap proposal ini menjadi dasar kerja sama. Jadwal, jumlah peserta, dan daftar sekolah dapat kami sesuaikan dengan arahan Bapak/Ibu."),
  p("Kontak: [ISI TIM: nama, nomor telepon, dan surel penghubung]"),
  h2("Lampiran A. Rujukan"),
  ...[
    "Kemendikbudristek, data satuan pendidikan dengan peserta didik berkebutuhan khusus per Desember 2023. Medcom, \"34 Ribu Sekolah dengan Siswa Disabilitas Tak Punya Guru Pembimbing Khusus\". https://www.medcom.id/pendidikan/cerita-guru/RkjGDw9K-34-ribu-sekolah-dengan-siswa-disabilitas-tak-punya-guru-pembimbing-khusus. Juga ANTARA, https://www.antaranews.com/berita/4038030",
    "Artikel jurnal Universitas Negeri Malang tentang guru pendamping khusus di sekolah inklusif. https://journal2.um.ac.id/index.php/jo/article/download/59676/15100",
    "ANTARA Kaltim, \"DPRD Kaltim minta Pemprov tingkatkan kualitas SLB\". https://kaltim.antaranews.com/berita/197400/dprd-kaltim-minta-pemprov-tingkatkan-kualitas-slb",
    "Sonora.id, \"Disdikbud Kaltim Sasar SLB di Kaltim sebagai Upaya Pemerataan Pendidikan\". https://www.sonora.id/read/424272404/disdikbud-kaltim-sasar-slb-di-kaltim-sebagai-upaya-pemerataan-pendidikan",
    "ANTARA Kaltim, \"PT Angkasa Pura I Sepinggan Balikpapan beri bantuan ke SLBN\". https://kaltim.antaranews.com/berita/207852/pt-angkasa-pura-i-sepinggan-balikpapan-beri-bantuan-ke-slbn",
    "Cabang Dinas Pendidikan Wilayah I (Balikpapan dan Penajam Paser Utara), Disdikbud Provinsi Kalimantan Timur, Petunjuk Teknis SPMB 2025. https://cabdin1.disdikbud.kaltimprov.go.id",
    "Peraturan Wali Kota Balikpapan Nomor 9 Tahun 2018 tentang Penyelenggaraan Pendidikan Inklusif. https://peraturan.bpk.go.id/Download/88906/PERWALI%20NOMOR%209%202018.pdf",
    "Peraturan Pemerintah Nomor 13 Tahun 2020 tentang Akomodasi yang Layak untuk Peserta Didik Penyandang Disabilitas. https://peraturan.go.id/files/pp13-2020pjl.pdf",
    "Pertamina Foundation, Penerima Beasiswa Sobat Istimewa Kota Balikpapan tahun 2024/2025 dan 2025/2026. Dokumen yang diterima tim Sanggabiz, tanpa tautan publik.",
    "Kemendikdasmen, portal Sekolah Kita, data sekolah semester 2026/2027 Gasal. https://sekolah.data.kemendikdasmen.go.id",
    "Model ADDIE (Analysis, Design, Development, Implementation and Evaluation). Jurnal Ilmiah Pendidikan Agama Islam, UIN Sunan Gunung Djati. https://journal.uinsgd.ac.id/index.php/jipai/article/download/11042/pdf",
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
          new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "PROPOSAL EVENT ORGANIZER", font: FONT, bold: true, size: 20, color: GOLD, characterSpacing: 60 })] }),
          new Paragraph({ spacing: { after: 240, line: 400, lineRule: "atLeast" }, children: [new TextRun({ text: "Penyelenggaraan Workshop Pembelajaran Berdiferensiasi untuk Guru SLB dan SD", font: FONT, bold: true, size: 52, color: "FFFFFF" })] }),
          new Paragraph({ spacing: { after: 2400 }, border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: GOLD, space: 14 } }, children: [new TextRun({ text: "Rekrutmen peserta, advokasi sekolah, dan pelaksanaan acara untuk 80 guru di Balikpapan, Kalimantan Timur", font: FONT, size: 24, color: "DCE6F2" })] }),
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
      ["80 guru", "dari 27 SLB dan SD di Balikpapan"],
      ["11 November 2026", `Rabu, registrasi 12.00 dan acara inti 13.00 ${WAKTU}`],
      ["Dikerjakan Sanggabiz", "Rekrutmen, advokasi, logistik, pelaksanaan, dan laporan"],
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
      new TextRun({ text: "Sanggabiz  |  Penyelenggaraan Workshop Pembelajaran Berdiferensiasi untuk Guru SLB dan SD", size: 14, color: MUTED }),
      new TextRun({ text: "\t", size: 14 }),
      new TextRun({ children: ["Halaman ", PageNumber.CURRENT], size: 14, color: MUTED }),
    ],
  })],
});

const page = { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } };
const doc = new Document({
  creator: "Sanggabiz", title: "Penyelenggaraan Workshop Pembelajaran Berdiferensiasi untuk Guru SLB dan SD",
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
const out = path.join(outDir, "v3 Proposal EO Workshop Pembelajaran Berdiferensiasi - SLB dan SD Balikpapan.docx");
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log("ditulis:", out);
  console.log("total", fmt(TOTAL), "per peserta", fmt(PER_PESERTA));
});
