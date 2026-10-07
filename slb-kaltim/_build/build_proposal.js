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
  children: [new TextRun({ text, size: 16, color: MUTED, italics: true })], spacing: { after: 160 },
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
    cantSplit: true,
    children: r.map((c, i) => {
      if (c && c.group) return cell(c.group, widths.reduce((a, b) => a + b, 0), { fill: "DCE6F2", bold: true, span: widths.length });
      const al = o.align && o.align[i];
      return cell(c && c.text !== undefined ? c.text : c, widths[i], {
        fill: c && typeof c.bg === "string" ? c.bg : (ri % 2 ? "F7F9FC" : undefined), align: al, bold: !!(c && typeof c === "object" && !Array.isArray(c) && c.bold === true),
      });
    }).slice(0, r[0] && r[0].group ? 1 : r.length),
  }));
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows: [head, ...body] });
}
const spacer = (n = 120) => new Paragraph({ spacing: { after: n }, children: [] });
const img = (file, w, h, type = "png") => new ImageRun({ type, data: fs.readFileSync(A(file)), transformation: { width: w, height: h } });
const centerImg = (file, w, h, type) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [img(file, w, h, type)] });

// ---------- anggaran ----------
const paket = {
  Hemat: {
    p1: 1000000, p2: 1000000, fas: 250000, mc: 500000, meet: 100000, modul: 1500000, cetak: 15000,
    kit: 15000, sert: 5000, rek: 2000000, pan: 300000, dok: 800000, eval: 1000000, tl: 250000, trans: 500000,
    kitNama: "Perlengkapan peserta (ID card dan pulpen)", sertNama: "Sertifikat peserta, cetak",
    cetakNama: "Cetak modul ringkas dan template (hitam putih, sampul warna)",
    dokNama: "Dokumentasi foto (1 fotografer)",
    tlNama: "Grup tindak lanjut 4 minggu, dimoderatori tim",
    meetNama: "Paket meeting setengah hari di hotel tengah kota (ruang kelas, proyektor dan suara, 1 kali coffee break, makan siang)",
  },
  Standar: {
    p1: 1500000, p2: 1500000, fas: 400000, mc: 750000, meet: 150000, modul: 2500000, cetak: 25000,
    kit: 45000, sert: 10000, rek: 3000000, pan: 500000, dok: 1800000, eval: 1500000, tl: 1500000, trans: 1000000,
    kitNama: "Perlengkapan peserta (tas jinjing, buku catatan, pulpen, ID card)", sertNama: "Sertifikat peserta, cetak dan map",
    cetakNama: "Cetak modul ringkas dan template (berwarna)",
    dokNama: "Dokumentasi foto dan video ringkas",
    tlNama: "Grup tindak lanjut 4 minggu dan 1 sesi daring 90 menit bersama pemateri",
    meetNama: "Paket meeting setengah hari di hotel tengah kota (ruang kelas lega, proyektor dan suara, 1 kali coffee break, makan siang)",
  },
};
function rab(nama) {
  const k = paket[nama];
  const items = [
    { group: "A. Pemateri dan fasilitator" },
    ["1", "Pemateri 1, pakar pendidikan inklusif (Sesi 1 dan 2)", 2, "jam", k.p1],
    ["2", "Pemateri 2, praktisi dan pemandu workshop (Sesi 2 sampai 4)", 2.5, "jam", k.p2],
    ["3", "Fasilitator kelompok, 1 orang per 20 peserta", 4, "orang", k.fas],
    ["4", "MC dan moderator", 1, "orang", k.mc],
    { group: "B. Tempat dan konsumsi" },
    ["5", k.meetNama, 90, "orang", k.meet],
    { group: "C. Materi dan perlengkapan" },
    ["6", "Penyusunan modul ringkas dan template", 1, "paket", k.modul],
    ["7", k.cetakNama, 90, "eksemplar", k.cetak],
    ["8", k.kitNama, 90, "set", k.kit],
    ["9", k.sertNama, 80, "lembar", k.sert],
    { group: "D. Rekrutmen dan koordinasi" },
    ["10", "Rekrutmen peserta: undangan, formulir daftar, konfirmasi, pengingat", 1, "paket", k.rek],
    ["11", "Koordinator lapangan dan panitia lokal, hari-H dan persiapan", 3, "orang", k.pan],
    { group: "E. Evaluasi, dokumentasi, dan tindak lanjut" },
    ["12", k.dokNama, 1, "paket", k.dok],
    ["13", "Pre-test, post-test, penilaian produk, dan laporan akhir", 1, "paket", k.eval],
    ["14", k.tlNama, 1, "paket", k.tl],
    { group: "F. Lain-lain" },
    ["15", "Transportasi lokal dan kebutuhan hari-H", 1, "paket", k.trans],
  ];
  let sub = 0;
  const rows = items.map((r) => {
    if (r.group) return [r];
    const jumlah = r[2] * r[4]; sub += jumlah;
    return [r[0], r[1], String(r[2]).replace(".", ","), r[3], fmt(r[4]), fmt(jumlah)];
  });
  const r1000 = (x) => Math.round(x / 1000) * 1000;
  const cad = r1000(sub * 0.05), fee = r1000((sub + cad) * 0.10), ppn = r1000((sub + cad + fee) * 0.11), total = sub + cad + fee + ppn;
  return { rows, sub, cad, fee, ppn, total, perPeserta: r1000(total / 80) };
}
const R = { Hemat: rab("Hemat"), Standar: rab("Standar") };
function rabTable(nama) {
  const r = R[nama];
  const widths = [500, 3900, 700, 1100, 1700, 1738];
  const sum = (label, val, bold) => [{ text: "", bg: "F2F5FA" }, { text: label, bold, bg: "F2F5FA" }, { text: "", bg: "F2F5FA" }, { text: "", bg: "F2F5FA" }, { text: "", bg: "F2F5FA" }, { text: fmt(val), bold, bg: "F2F5FA" }];
  return table(widths, ["No", "Komponen", "Vol", "Satuan", "Harga satuan", "Jumlah"], [
    ...r.rows,
    sum("Subtotal", r.sub, true),
    sum("Cadangan 5 persen", r.cad),
    sum("Biaya manajemen Sanggabiz 10 persen", r.fee),
    sum("PPN 11 persen", r.ppn),
    sum("TOTAL PAKET " + nama.toUpperCase(), r.total, true),
    sum("Per peserta (80 orang)", r.perPeserta, true),
  ], { align: [AlignmentType.CENTER, AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.RIGHT, AlignmentType.RIGHT] });
}

// ---------- isi ----------
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
  p("Sanggabiz menyampaikan penawaran workshop setengah hari untuk 80 guru SLB dan sekolah inklusif di Balikpapan. Workshop ini berfokus pada satu kemampuan, yaitu menyusun modul ajar berdiferensiasi yang bisa langsung dipakai di kelas. Setiap peserta pulang membawa template dan modul ringkas, serta satu draf modul ajar yang sudah mulai mereka susun dalam sesi."),
  p("Proposal ini memuat latar belakang, tujuan, tempat dan peserta, silabus, rundown, anggaran dalam dua paket (Hemat dan Standar), dan profil pemateri. Kami siap menyesuaikan jadwal dan isi sesuai arahan Bapak/Ibu."),
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
  p("Pendidikan inklusif menempatkan anak berkebutuhan khusus di kelas yang sama dengan teman sebayanya. Agar itu berhasil, guru harus bisa menyesuaikan cara mengajar dengan kebutuhan tiap murid. Menurut data Kemendikbudristek per Desember 2023, ada 40.164 satuan pendidikan formal di Indonesia yang memiliki peserta didik berkebutuhan khusus. Hanya 5.956 di antaranya, atau 14,83 persen, yang memiliki guru pembimbing khusus [1]."),
  centerImg("grafik-guru-pembimbing-khusus.png", 560, 192),
  caption("Gambar 1. Satuan pendidikan dengan peserta didik berkebutuhan khusus menurut ketersediaan guru pembimbing khusus."),
  p("Guru yang mendampingi anak berkebutuhan khusus di sekolah inklusif sering bukan lulusan Pendidikan Luar Biasa. Banyak yang berlatar PGSD, PGPAUD, atau bidang nonkependidikan. Mereka menghadapi rasio pendampingan yang tinggi dengan pelatihan yang terbatas [2]."),
  p("Kurikulum Merdeka meminta guru merancang pembelajaran sesuai kesiapan, minat, dan profil belajar murid. Kajian di beberapa sekolah dasar inklusif menemukan kendala yang sama: pemahaman guru tentang kebutuhan belajar anak masih terbatas, pelatihan kurang, waktu persiapan sempit, dan jumlah murid per kelas banyak. Pembelajaran berdiferensiasi sudah mulai dijalankan, tetapi belum menyentuh semua aspek, terutama konten dan lingkungan belajar [3][4][5]."),
  p("Kebutuhan serupa tercatat di Kalimantan Timur. Komisi IV DPRD Kaltim menyatakan guru berlatar pendidikan luar biasa masih kurang dan mengusulkan pembukaan program studi PLB bersama Universitas Mulawarman [6]. SLB negeri berada di bawah Dinas Pendidikan dan Kebudayaan Provinsi, yang mengelola 11 SLB di Kaltim [7]. Salah satunya SLB Negeri Balikpapan, berdiri sejak 1984 dan melayani 401 siswa [8]. Sekolah inklusif di Balikpapan berada di bawah Pemerintah Kota, yang telah menerbitkan Peraturan Wali Kota Nomor 3 Tahun 2013 tentang penyelenggaraan pendidikan inklusif [9]."),
  centerImg("grafik-siswa-slbn-balikpapan.png", 520, 213),
  caption("Gambar 2. Jumlah siswa SLB Negeri Balikpapan per jenjang."),
  p("Peraturan Pemerintah Nomor 13 Tahun 2020 mewajibkan pemerintah pusat dan daerah memfasilitasi akomodasi yang layak bagi peserta didik penyandang disabilitas, termasuk menyiapkan pendidik dan menyediakan kurikulum [10]. Penguatan kemampuan guru adalah bagian dari kewajiban itu, dan workshop ini dirancang untuk membantu memenuhinya."),
  h2("1.2 Mengapa Pembelajaran Berdiferensiasi"),
  p("Kami memilih satu kemampuan agar peserta pulang dengan hasil yang nyata. Ada tiga alasan."),
  bullet("Setiap guru di kelas inklusif mengajar murid dengan kebutuhan berbeda dalam satu ruangan. Menyesuaikan konten, proses, dan produk belajar adalah pekerjaan harian mereka."),
  bullet("Hasilnya bisa dibuat dalam setengah hari. Peserta menyusun satu modul ajar memakai template, lalu membawanya pulang."),
  bullet("Materinya berlaku untuk semua jenis kekhususan, sehingga cocok untuk peserta campuran dari SLB dan sekolah inklusif."),
  p("Kami belum menemukan data khusus Balikpapan tentang tingkat pemahaman guru terhadap pembelajaran berdiferensiasi. Karena itu workshop dibuka dengan pre-test. Hasilnya mencatat posisi awal peserta dan menjadi pembanding untuk hasil akhir.", { before: 80 }),
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
    ["Produk", "Minimal 80 persen peserta menghasilkan draf modul ajar", "Rubrik 4 butir: kesesuaian profil murid, ragam konten, ragam proses, pilihan produk"],
    ["Kepuasan", "Rata-rata minimal 4 dari 5", "Kuesioner akhir"],
    ["Tindak lanjut", "Minimal 75 persen peserta bergabung di grup tindak lanjut", "Data anggota grup"],
  ]),
  h2("1.5 Tentang Sanggabiz"),
  p("Sanggabiz adalah mitra solusi bisnis yang menyediakan layanan department-as-a-service: keuangan, sumber daya manusia, layanan pelanggan, pemasaran, manajemen proyek, administrasi, dan operasional. Di bidang pelatihan, kami mengelola program dari perencanaan sampai pelaporan."),
  table([2400, 7238], ["Hal", "Keterangan"], [
    ["Skala", "400+ proyek dan layanan, 500+ pakar dan trainer, 6.000+ mitra engagement"],
    ["Pengalaman multikota", "Pelatihan MSME Rumah BUMN di lebih dari sembilan kota, termasuk Samarinda"],
    ["Klien", "Lembaga pendidikan, BUMN, yayasan, dan pemerintah daerah, antara lain Universitas Muhammadiyah Yogyakarta, Yakes Pertamina, dan Pemerintah Kota Yogyakarta"],
    ["Kerangka kerja pelatihan", "Define, Design, Deliver, Discuss"],
  ]),
  p("Sanggabiz belum pernah menyelenggarakan program khusus pendidikan luar biasa. Karena itu isi materi dipegang pemateri lokal yang berpengalaman di bidang pendidikan inklusif. Kami mengerjakan bagian yang kami kuasai: perekrutan peserta, logistik, penyusunan modul, evaluasi, dan pelaporan.", { before: 120 }),
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [4819, 4819],
    rows: [new TableRow({ children: [["foto-pelatihan-1.jpg", 227, 170], ["foto-pelatihan-2.jpg", 302, 170]].map(([f, w, h]) => new TableCell({
      width: { size: 4819, type: WidthType.DXA },
      borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [img(f, w, h, "jpg")] })],
    })) })],
  }),
  caption("Gambar 3. Dokumentasi pelatihan yang diselenggarakan Sanggabiz."),
];

const bab2 = [
  h1("2. Tempat, Waktu, dan Peserta", true),
  h2("2.1 Tempat Pelaksanaan"),
  p("Workshop diadakan di Balikpapan, di hotel kawasan tengah kota (Jalan Jenderal Sudirman dan sekitarnya) agar mudah dijangkau dari SLB dan sekolah di berbagai kecamatan."),
  table([3000, 6638], ["Hal", "Keterangan"], [
    ["Hotel terpilih", "[ISI TIM: nama hotel setelah penawaran dikonfirmasi]"],
    ["Ruang", "Kapasitas minimal 80 orang dengan susunan kelompok, proyektor dan tata suara"],
    ["Akses", "Lift atau jalur kursi roda, toilet yang bisa dipakai pengguna kursi roda, dan lahan parkir. Beberapa guru adalah penyandang disabilitas"],
    ["Alternatif", "Aula milik pemerintah daerah bila tersedia, yang dapat menurunkan biaya"],
  ]),
  h2("2.2 Waktu Pelaksanaan"),
  table([3000, 6638], ["Hal", "Keterangan"], [
    ["Tanggal", "[ISI TIM: tanggal pelaksanaan]"],
    ["Jam", "07.30 sampai 13.45 WIB, termasuk registrasi dan makan siang"],
    ["Durasi acara inti", "5 jam, dari 08.00 sampai 13.00 WIB"],
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
    ["H-7", "Konfirmasi kehadiran lewat WhatsApp. Peserta cadangan dihubungi bila ada yang mundur"],
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
    ["Pengajar", "2 pemateri lokal dan 4 fasilitator kelompok"],
    ["Hasil", "Template, modul ringkas, draf modul ajar sendiri, rencana tindak lanjut, sertifikat"],
  ]),
  h2("3.2 Kerangka Kerja"),
  table([2000, 7638], ["Tahap", "Isi di workshop ini"], [
    ["Define", "Pre-test dan profil kelas peserta, dikumpulkan lewat formulir pendaftaran"],
    ["Design", "Template modul ajar dan modul ringkas disiapkan tim bersama pemateri"],
    ["Deliver", "Workshop tatap muka dengan praktik terbimbing"],
    ["Discuss", "Presentasi hasil, umpan balik, post-test, evaluasi, dan grup tindak lanjut"],
  ]),
  h2("3.3 Silabus per Sesi"),
  table([760, 2900, 2200, 1700, 2078], ["Sesi", "Topik dan isi", "Aktivitas", "Hasil", "Ukuran keberhasilan"], [
    ["1", ["**Ragam kebutuhan belajar dan prinsip pembelajaran berdiferensiasi** (80 menit, materi)", "Ragam kebutuhan belajar murid secara umum, empat unsur diferensiasi (konten, proses, produk, lingkungan belajar), contoh praktik di kelas campuran"], "Paparan singkat dan diskusi tiga kasus murid", "Catatan kunci peserta", "Minimal 80 persen peserta dapat menyebut empat unsur diferensiasi di post-test"],
    ["2", ["**Membaca profil belajar murid** (45 menit: 20 materi, 25 praktik)", "Asesmen diagnostik sederhana, cara menggali informasi dari murid, orang tua, dan guru sebelumnya, mengisi lembar profil"], "Contoh asesmen, lalu mengisi lembar profil untuk satu murid dari kelas sendiri", "Lembar profil satu murid", "Minimal 80 persen peserta menyelesaikan lembar profil"],
    ["3", ["**Workshop menyusun modul ajar berdiferensiasi** (90 menit, praktik)", "Memilih capaian pembelajaran, merancang dua ragam konten, dua ragam proses, satu pilihan produk, dan penyesuaian lingkungan belajar"], "Kerja kelompok: 8 kelompok, tiap kelompok sekitar 10 guru dengan satu fasilitator, memakai template", "Draf modul ajar pada template", "Minimal 80 persen peserta menghasilkan draf modul ajar"],
    ["4", ["**Presentasi mini dan umpan balik** (20 menit)"], "Tiap kelompok menunjukkan satu contoh hasil, pemateri memberi umpan balik", "Catatan perbaikan", "Setiap kelompok mendapat umpan balik"],
    ["5", ["**Rencana tindak lanjut dan evaluasi** (20 menit)"], "Menulis rencana penerapan dua minggu, mengisi post-test dan kuesioner", "Rencana tindak lanjut satu halaman", "Semua peserta mengisi post-test dan kuesioner"],
  ], { align: [AlignmentType.CENTER] }),
  h2("3.4 Hasil yang Dibawa Pulang"),
  bullet("Template modul ajar berdiferensiasi, cetak dan digital (contoh di Lampiran A)."),
  bullet("Modul ringkas sekitar 20 halaman (isi di Lampiran B)."),
  bullet("Draf modul ajar buatan sendiri."),
  bullet("Rencana tindak lanjut dua minggu."),
  bullet("Sertifikat peserta."),
  h2("3.5 Evaluasi dan Tindak Lanjut"),
  p("Pre-test dan post-test memakai 10 soal yang sama, sehingga kenaikan skor bisa dibandingkan langsung. Produk peserta dinilai dengan rubrik empat butir. Kuesioner kepuasan diisi di akhir acara."),
  p("Tindak lanjut berjalan empat minggu lewat grup WhatsApp yang dimoderatori tim Sanggabiz untuk tanya jawab dan berbagi hasil penerapan. Pada Paket Standar, ditambah satu sesi daring 90 menit bersama pemateri di minggu ketiga untuk membahas hasil di kelas."),
  p("Laporan akhir kami serahkan dalam 7 hari kerja setelah acara. Isinya daftar hadir, hasil tes, penilaian produk, hasil kuesioner, dan dokumentasi."),
];

const rundownRows = [
  ["07.30 - 08.00", "30 menit", "Registrasi, daftar hadir, pembagian perlengkapan", "Panitia"],
  ["08.00 - 08.20", "20 menit", "Pembukaan, sambutan, pre-test", "MC dan Sanggabiz"],
  ["08.20 - 09.40", "80 menit", "**Sesi 1.** Ragam kebutuhan belajar dan prinsip pembelajaran berdiferensiasi", "Pemateri 1"],
  ["09.40 - 09.55", "15 menit", "Coffee break", "Panitia"],
  ["09.55 - 10.40", "45 menit", "**Sesi 2.** Membaca profil belajar murid", "Pemateri 1 dan 2"],
  ["10.40 - 12.10", "90 menit", "**Sesi 3.** Workshop menyusun modul ajar berdiferensiasi", "Pemateri 2 dan fasilitator"],
  ["12.10 - 12.30", "20 menit", "**Sesi 4.** Presentasi mini dan umpan balik", "Pemateri 1 dan 2"],
  ["12.30 - 12.50", "20 menit", "**Sesi 5.** Rencana tindak lanjut, post-test, kuesioner", "Sanggabiz"],
  ["12.50 - 13.00", "10 menit", "Penutupan dan pembagian sertifikat", "MC"],
  ["13.00 - 13.45", "45 menit", "Makan siang dan ramah tamah", "Panitia"],
];
const bab4 = [
  h1("4. Rundown Acara"),
  p("Jam di bawah mengikuti usulan awal. Bila tanggal dan tempat berubah, jam dapat digeser tanpa mengubah durasi tiap sesi."),
  table([1700, 1100, 4700, 2138], ["Waktu (WIB)", "Durasi", "Kegiatan", "Pemandu"], rundownRows, { align: [AlignmentType.CENTER, AlignmentType.CENTER] }),
  p("Perbandingan waktu: materi 100 menit (sekitar 40 persen) dan praktik 155 menit (sekitar 60 persen) dari 255 menit waktu belajar.", { before: 140 }),
];

const bab5 = [
  h1("5. Anggaran", true),
  h2("5.1 Dasar Penyusunan"),
  bullet("Peserta 80 orang. Konsumsi dan perlengkapan dihitung 90 orang untuk peserta, pemateri, fasilitator, dan panitia."),
  bullet("Harga tempat mengacu pada paket meeting setengah hari di hotel tengah kota Balikpapan, sekitar Rp 148.500 sampai Rp 150.000 per orang pada Oktober 2026. Paket Hemat memakai Rp 100.000 per orang untuk ruang yang lebih sederhana. Harga final mengikuti penawaran tertulis hotel."),
  bullet("Honor pemateri berada di bawah batas Rp 1.700.000 per jam yang lazim dipakai untuk narasumber pakar atau praktisi dalam standar biaya pemerintah daerah. Pemateri berasal dari Balikpapan atau Kaltim, sehingga tidak ada biaya tiket dan penginapan."),
  bullet("Biaya manajemen Sanggabiz 10 persen dan PPN 11 persen dihitung atas subtotal setelah cadangan."),
  h2("5.2 Rincian Paket Hemat"),
  rabTable("Hemat"),
  h2("5.3 Rincian Paket Standar"),
  rabTable("Standar"),
  h2("5.4 Perbandingan Paket"),
  table([3400, 3119, 3119], ["Hal", "Hemat", "Standar"], [
    ["Ruang", "Ruang meeting sederhana, kapasitas pas", "Ruang lebih lega untuk kerja kelompok"],
    ["Honor pemateri per jam", fmt(paket.Hemat.p1), fmt(paket.Standar.p1)],
    ["Perlengkapan peserta", "ID card dan pulpen", "Tas jinjing, buku catatan, pulpen, ID card"],
    ["Modul dicetak", "Hitam putih, sampul warna", "Berwarna"],
    ["Dokumentasi", "Foto", "Foto dan video ringkas"],
    ["Tindak lanjut", "Grup WhatsApp 4 minggu", "Grup WhatsApp 4 minggu dan 1 sesi daring 90 menit"],
    [{ text: "Total", bold: true }, { text: fmt(R.Hemat.total), bold: true }, { text: fmt(R.Standar.total), bold: true }],
    [{ text: "Per peserta", bold: true }, { text: fmt(R.Hemat.perPeserta), bold: true }, { text: fmt(R.Standar.perPeserta), bold: true }],
  ]),
  p("Kami merekomendasikan Paket Standar. Ruangnya lebih lega untuk delapan kelompok kerja, dan sesi daring di minggu ketiga membantu guru memperbaiki modul setelah mencobanya di kelas. Paket Hemat tetap memenuhi seluruh tujuan inti dan seluruh hasil yang dibawa pulang.", { before: 140 }),
];

function pemateriCard(no, peran) {
  const lab = (t) => ({ text: t, bold: true, bg: LIGHT });
  return table([2600, 7038], [`Pemateri ${no}: ${peran}`, ""], [
    [lab("Nama dan gelar"), "[ISI TIM: nama lengkap dan gelar]"],
    [lab("Jabatan dan instansi"), "[ISI TIM: jabatan dan instansi]"],
    [lab("Bidang keahlian"), "[ISI TIM: bidang keahlian]"],
    [lab("Riwayat pendidikan"), "[ISI TIM: pendidikan terakhir dan bidang]"],
    [lab("Pengalaman melatih guru"), "[ISI TIM: pengalaman melatih guru, jumlah pelatihan, tahun]"],
    [lab("Pengalaman menyusun modul atau program pembelajaran individual"), "[ISI TIM: pengalaman terkait]"],
    [lab("Peran di workshop"), no === 1 ? "Sesi 1 dan Sesi 2 (materi), umpan balik di Sesi 4" : "Sesi 2 (praktik), Sesi 3 (workshop), umpan balik di Sesi 4"],
    [lab("Foto"), "[ISI TIM: tempel foto pemateri di sini]"],
  ]);
}
const bab6 = [
  h1("6. Pemateri", true),
  p("Pemateri berasal dari Balikpapan atau Kalimantan Timur. Dengan begitu tidak ada biaya mobilisasi, dan jaringan lokalnya bisa dipakai untuk tindak lanjut. Kriteria yang kami pakai:", { after: 60 }),
  bullet("berdomisili di Balikpapan atau Kalimantan Timur,"),
  bullet("berlatar pendidikan luar biasa atau pendidikan inklusif, atau praktisi dengan pengalaman nyata mendampingi murid berkebutuhan khusus,"),
  bullet("pernah melatih guru atau memandu workshop,"),
  bullet("pernah menyusun modul ajar atau program pembelajaran individual."),
  spacer(80),
  pemateriCard(1, "Pakar Pendidikan Inklusif"),
  spacer(200),
  pemateriCard(2, "Praktisi dan Pemandu Workshop"),
];

const bab7 = [
  h1("7. Penutup", true),
  p("Workshop ini sengaja dibuat sempit: satu kemampuan, satu hari kerja, satu produk yang bisa dipakai. Kami berharap proposal ini dapat menjadi dasar diskusi lebih lanjut. Jadwal, jumlah peserta, dan isi sesi dapat kami sesuaikan dengan arahan Bapak/Ibu."),
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
  ].map((t, i) => new Paragraph({ spacing: { after: 70, line: 260 }, indent: { left: 420, hanging: 420 }, children: [new TextRun({ text: `[${i + 1}]\t${t}`, size: 16 })], tabStops: [{ type: TabStopType.LEFT, position: 420 }] })),
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
const out = path.join(outDir, "Proposal Workshop Pembelajaran Berdiferensiasi - SLB dan Inklusif Balikpapan.docx");
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log("ditulis:", out);
  for (const n of ["Hemat", "Standar"]) console.log(n, "subtotal", fmt(R[n].sub), "total", fmt(R[n].total), "per peserta", fmt(R[n].perPeserta));
});
