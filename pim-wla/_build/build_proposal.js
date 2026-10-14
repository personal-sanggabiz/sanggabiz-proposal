// Membuat proposal docx WLA PIM. Jalankan: NODE_PATH=$(npm root -g) node build_proposal.js
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
const h1 = (text, pb = true) => new Paragraph({ heading: "Heading1", children: [new TextRun({ text })], pageBreakBefore: pb, keepNext: true });
const h2 = (text) => new Paragraph({ heading: "Heading2", children: [new TextRun({ text })], keepNext: true });
const h3 = (text) => new Paragraph({ children: [new TextRun({ text, bold: true, size: 20, color: NAVY })], spacing: { before: 200, after: 90 }, keepNext: true });
const bullet = (text) => new Paragraph({ numbering: { reference: "bul", level: 0 }, children: runs(text), spacing: { after: 70, line: 290 } });
const num = (text, ref = "num") => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(text), spacing: { after: 70, line: 290 } });
const spacer = (n = 120) => new Paragraph({ spacing: { after: n }, children: [] });
const fmt = (n) => "Rp " + Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");

// kotak catatan (satu sel berlatar muda dengan garis emas di kiri)
function callout(title, text) {
  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: W, type: WidthType.DXA }, shading: { fill: LIGHT, type: ShadingType.CLEAR, color: "auto" },
      borders: { top: none, bottom: none, right: none, left: { style: BorderStyle.SINGLE, size: 24, color: GOLD } },
      margins: { top: 110, bottom: 110, left: 180, right: 160 },
      children: [
        new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: title, bold: true, size: 17, color: NAVY, characterSpacing: 20 })] }),
        ...(Array.isArray(text) ? text : [text]).map((t) => new Paragraph({ spacing: { after: 40, line: 280 }, children: runs(t, { size: 18 }) })),
      ],
    })] })],
  });
}

// ---------- penolong tabel ----------
const thin = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: thin, bottom: thin, left: thin, right: thin };
function cell(content, width, o = {}) {
  const lines = Array.isArray(content) ? content : [content];
  return new TableCell({
    width: { size: width, type: WidthType.DXA }, borders, columnSpan: o.span,
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
  const total = widths.reduce((a, b) => a + b, 0);
  const head = new TableRow({
    tableHeader: true, cantSplit: true,
    children: header.map((t, i) => cell(t, widths[i], { fill: NAVY, color: "FFFFFF", bold: true, size: 16, align: o.headAlign && o.headAlign[i] })),
  });
  const body = rows.map((r, ri) => {
    if (r && r.group) return new TableRow({ cantSplit: true, children: [cell(r.group, total, { fill: "DCE6F2", bold: true, span: widths.length })] });
    return new TableRow({
      cantSplit: !o.split,
      children: r.map((c, i) => {
        const obj = c && typeof c === "object" && !Array.isArray(c);
        return cell(obj && c.text !== undefined ? c.text : c, widths[i], {
          fill: obj && typeof c.bg === "string" ? c.bg : (o.firstBold && i === 0 ? LIGHT : (ri % 2 ? "F7F9FC" : undefined)),
          align: o.align && o.align[i], size: o.size, bold: (obj && c.bold === true) || (o.firstBold && i === 0),
        });
      }),
    });
  });
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: [head, ...body] });
}
const img = (file, w, h) => new ImageRun({ type: "png", data: fs.readFileSync(A(file)), transformation: { width: w, height: h } });

// ---------- timeline (format standar tipe 2) ----------
function gantt() {
  const cols = ["M1-2", "M3-4", "M5-6", "M7-8", "M9-10", "M11-12", "M13-14", "M15-16", "M17-18"];
  const rows = [
    ["Kick-off, penguncian ruang lingkup, data room", [0]],
    ["Pendataan posisi dan SOP yang ada", [0, 1]],
    ["Survei dan workload diary", [1, 2]],
    ["Wawancara BEI", [1, 2, 3]],
    ["Perhitungan FTE dan analisis WLA", [2, 3, 4]],
    ["Penguraian proses dan draf RACI", [2, 3, 4, 5]],
    ["Workshop validasi per unit", [4, 5, 6]],
    ["Kalibrasi lintas fungsi", [5, 6]],
    ["Rekomendasi terintegrasi", [6, 7]],
    ["Pengesahan eksekutif dan serah terima", [7, 8]],
  ];
  const wNo = 450, wAct = 3100, wC = Math.floor((W - wNo - wAct) / 9);
  const widths = [wNo, wAct, ...cols.map(() => wC)];
  const total = widths.reduce((a, b) => a + b, 0);
  const head = new TableRow({ tableHeader: true, children: ["No", "Tahapan kegiatan", ...cols].map((t, i) => cell(t, widths[i], { fill: LIGHT, bold: true, size: 14, align: i === 1 ? AlignmentType.LEFT : AlignmentType.CENTER })) });
  const body = rows.map(([name, on], ri) => new TableRow({ cantSplit: true, children: [
    cell(String(ri + 1), wNo, { align: AlignmentType.CENTER, size: 15 }),
    cell(name, wAct, { size: 15 }),
    ...cols.map((_, ci) => on.includes(ci)
      ? cell("V", wC, { fill: NAVY, color: "FFFFFF", bold: true, align: AlignmentType.CENTER, size: 15 })
      : cell("", wC)),
  ] }));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: [head, ...body] });
}

// ---------- anggaran (satu opsi, harga akhir sudah termasuk pajak) ----------
const KONSULTANSI = 125000000, PELATIHAN = 40000000, TOTAL = KONSULTANSI + PELATIHAN;

// ---------- kurikulum pelatihan (format standar tipe 1) ----------
const HARI = [
  { judul: "Hari 1. Dasar ABK dan Membaca Temuan", modul: [
    { no: 1, judul: "Fondasi ABK dan Konteks PIM", tujuan: "Peserta mampu menjelaskan tujuan ABK dan membedakan beban kerja, kinerja, dan jumlah pegawai (headcount).",
      materi: ["Tujuan ABK dan kaitannya dengan struktur dan proses", "FTE dan BEI", "Batas antara beban kerja, kinerja, dan headcount", "Temuan studi PIM sebelumnya"],
      akt: ["Menit 0-15: fasilitator membuka dan memaparkan temuan studi PIM sebelumnya.", "Menit 15-40: fasilitator menjelaskan konsep dengan contoh satu posisi.", "Menit 40-75: lima kelompok (masing-masing 6 orang) menulis lima isu beban kerja di unitnya pada kartu, menempelkan, dan mengelompokkannya di flipchart.", "Menit 75-90: tanya jawab dan penegasan istilah."],
      alat: "Slide, flipchart, kartu isu, spidol", ukur: "Setiap kelompok menghasilkan peta isu. Minimal 80 persen peserta menjawab benar 4 dari 5 butir kuis istilah.", out: "Peta isu awal per kelompok dan hasil pre-test" },
    { no: 2, judul: "Logika Perhitungan", tujuan: "Peserta mampu menghitung FTE satu posisi dari volume, frekuensi, dan waktu penyelesaian.",
      materi: ["Jam kerja efektif tahunan", "Volume, frekuensi, dan waktu standar", "Annualisasi dan productive allowance", "Confidence level dan asumsi perhitungan"],
      akt: ["Menit 0-25: fasilitator menjelaskan rumus dengan satu contoh hitung.", "Menit 25-70: setiap peserta menerima lembar kasus satu posisi berisi 8 aktivitas, lalu menghitung jam kerja tahunan dan FTE dengan kalkulator atau lembar hitung.", "Menit 70-90: pembahasan kunci jawaban dan sebab selisih."],
      alat: "Lembar kasus, lembar hitung FTE, kalkulator", ukur: "Minimal 80 persen peserta menghasilkan FTE dengan selisih paling banyak 0,10 dari kunci jawaban.", out: "Lembar latihan hitung FTE" },
    { no: 3, judul: "Membaca Temuan", tujuan: "Peserta mampu menganalisis heatmap dan menelusuri akar penyebab beban kerja tinggi.",
      materi: ["Heatmap underload, normal, stretched, overload", "Root-cause tree", "Capacity gap", "Risiko salah tafsir"],
      akt: ["Menit 0-20: fasilitator menunjukkan contoh heatmap dengan data yang sudah dianonimkan.", "Menit 20-65: tiap kelompok menganalisis satu studi kasus berisi tiga posisi overload dan menggambar pohon akar masalah.", "Menit 65-90: satu kelompok presentasi, fasilitator dan kelompok lain memberi umpan balik."],
      alat: "Cetak heatmap, studi kasus, flipchart", ukur: "Setiap kelompok menemukan sedikitnya dua akar masalah per posisi dan satu risiko salah tafsir.", out: "Analisis studi kasus dan pohon akar masalah" },
    { no: 4, judul: "Management Interpretation Clinic", tujuan: "Peserta mampu merumuskan hipotesis dan pertanyaan validasi bagi Manager, VP, HR, dan incumbent.",
      materi: ["Gejala dan akar masalah", "Pertanyaan validasi per peran", "Cara menyampaikan temuan kepada manajemen"],
      akt: ["Menit 0-20: fasilitator membahas pola gejala dan akar masalah.", "Menit 20-60: peserta berkelompok menurut peran, lalu menyusun hipotesis dan sedikitnya tiga pertanyaan validasi untuk tiap peran.", "Menit 60-90: keliling antarkelompok (gallery walk), setiap kelompok memberi satu masukan."],
      alat: "Template hipotesis, flipchart", ukur: "Tiap hipotesis terhubung ke satu gejala dan memiliki pertanyaan validasi.", out: "Daftar hipotesis dan pertanyaan validasi" },
  ] },
  { judul: "Hari 2. Pengumpulan Data dan Wawancara", modul: [
    { no: 5, judul: "Desain Pengumpulan Data", tujuan: "Peserta mampu menyusun activity inventory yang lengkap untuk satu posisi.",
      materi: ["Activity inventory dan definisi output", "Pemicu dan unit volume", "Frekuensi dan seasonality", "Sumber bukti dan workload diary"],
      akt: ["Menit 0-20: fasilitator menjelaskan setiap kolom template.", "Menit 20-70: tiap peserta mengisi template untuk satu posisi di unitnya, sedikitnya 8 aktivitas.", "Menit 70-90: peserta bertukar template berpasangan dan memeriksa dengan daftar cek."],
      alat: "Template activity inventory, daftar cek", ukur: "Minimal 80 persen peserta mengisi 8 aktivitas dengan semua kolom wajib terisi.", out: "Draf activity inventory" },
    { no: 6, judul: "Teknik Wawancara ABK", tujuan: "Peserta mampu menyusun panduan wawancara terstruktur yang menggali hidden work.",
      materi: ["Structured interview dan probing", "Kejadian kritis, rework, waiting, pekerjaan insidental", "Peak period", "Bias self-report"],
      akt: ["Menit 0-25: fasilitator menjelaskan teknik probing dan bias.", "Menit 25-60: dua fasilitator memperagakan satu wawancara, peserta mengamati dengan rubrik lalu membahasnya.", "Menit 60-90: peserta berpasangan menyusun interview guide."],
      alat: "Rubrik pengamatan, template interview guide", ukur: "Panduan memuat sedikitnya 10 pertanyaan terbuka, termasuk 3 pertanyaan untuk hidden work.", out: "Interview guide terisi" },
    { no: 7, judul: "Role Play dan Observasi", tujuan: "Peserta mampu mewawancarai dengan pertanyaan netral dan mencatat bukti secara tertib.",
      materi: ["Peran interviewer, interviewee, observer", "Pencatatan bukti", "Neutral probing", "Kalibrasi estimasi waktu"],
      akt: ["Menit 0-15: pembagian kelompok tiga orang dan penjelasan skenario.", "Menit 15-75: tiga putaran, masing-masing 20 menit (wawancara 12 menit, umpan balik 8 menit). Peran berganti setiap putaran.", "Menit 75-90: refleksi bersama."],
      alat: "Kartu skenario, rubrik observasi", ukur: "Setiap peserta memegang ketiga peran. Minimal 75 persen peserta mendapat skor rubrik 3 dari 5 atau lebih.", out: "Rubrik observasi terisi dan umpan balik" },
    { no: 8, judul: "Supervised Data Clinic", tujuan: "Peserta mampu merekonsiliasi data wawancara dengan dokumen pendukung pada kasus nyata.",
      materi: ["Pengisian worksheet dari sampel kasus PIM", "Rekonsiliasi dokumen dan catatan wawancara", "Pencatatan selisih dalam issue log"],
      akt: ["Menit 0-15: fasilitator mengantar kasus yang sudah disiapkan bersama PIM dan dianonimkan.", "Menit 15-70: kelompok mengisi worksheet dari dokumen dan catatan wawancara, lalu mencatat selisih.", "Menit 70-90: fasilitator memeriksa dan memberi catatan."],
      alat: "Paket kasus, worksheet, issue log", ukur: "Setiap kelompok mencatat sedikitnya tiga selisih data beserta cara menyelesaikannya.", out: "Worksheet praktik dan issue log" },
  ] },
  { judul: "Hari 3. Kualitas Data, Kalibrasi, dan Tindak Lanjut", modul: [
    { no: 9, judul: "Quality Assurance Data", tujuan: "Peserta mampu memeriksa kualitas data dengan daftar cek dan menemukan kesalahan umum.",
      materi: ["Completeness check", "Duplikasi aktivitas dan unit tidak cocok", "Outlier dan double counting", "Reasonableness dan audit trail"],
      akt: ["Menit 0-25: fasilitator menjelaskan daftar cek QA.", "Menit 25-65: tiap kelompok menerima worksheet dengan enam kesalahan yang sengaja disisipkan dan mencarinya.", "Menit 65-90: pembahasan temuan."],
      alat: "Worksheet berisi kesalahan, QA checklist", ukur: "Setiap kelompok menemukan sedikitnya 5 dari 6 kesalahan.", out: "QA checklist dan data issue log" },
    { no: 10, judul: "Review dan Kalibrasi", tujuan: "Peserta mampu mengevaluasi hasil ABK, memberi confidence rating, dan mencatat keputusan.",
      materi: ["Review lintas peran", "Manager calibration", "Confidence rating dan sensitivitas asumsi", "Pencatatan keputusan"],
      akt: ["Menit 0-20: fasilitator menjelaskan alur kalibrasi.", "Menit 20-70: peserta berperan sebagai manajer unit dan HR, mereview satu paket hasil, memberi confidence rating, menguji satu asumsi (jam kerja efektif), dan mencatat keputusan.", "Menit 70-90: pembahasan hasil."],
      alat: "Paket hasil ABK, template decision log", ukur: "Setiap penyesuaian di decision log memuat keputusan, alasan, dan pemilik.", out: "Sampel tervalidasi dan decision log" },
    { no: 11, judul: "Tindak Lanjut dan RACI", tujuan: "Peserta mampu memilih tindak lanjut atas temuan ABK dan menguji dampaknya pada RACI.",
      materi: ["Simplifikasi, redistribusi, redesign, otomasi, reskill, penambahan kapasitas", "Uji kapasitas pada penugasan Responsible dan Accountable"],
      akt: ["Menit 0-20: fasilitator menjelaskan pilihan tindak lanjut.", "Menit 20-70: kelompok memetakan tiga temuan ke tindakan dan menguji siapa yang menjadi R dan A baru serta kecukupan kapasitasnya.", "Menit 70-90: presentasi singkat tiap kelompok."],
      alat: "Template action matrix, contoh RACI", ukur: "Setiap tindakan memiliki pemilik dan hasil uji kapasitas.", out: "Prioritized action matrix" },
    { no: 12, judul: "Rencana Tindak dan Penutupan", tujuan: "Peserta mampu menyusun rencana 30-60-90 hari untuk unitnya.",
      materi: ["Pemilik, target, dan indikator", "Tata kelola peninjauan", "Post-test dan komitmen unit"],
      akt: ["Menit 0-35: tiap peserta menyusun rencana 30-60-90 hari unitnya dengan template.", "Menit 35-60: berbagi dalam kelompok dan saling memberi masukan.", "Menit 60-80: post-test (20 menit) dan kuesioner.", "Menit 80-90: komitmen unit dan pembagian sertifikat elektronik."],
      alat: "Template rencana 30-60-90, soal post-test", ukur: "Skor post-test minimal 75 dan peningkatan rata-rata minimal 20 persen dari baseline.", out: "Rencana 30-60-90 hari, hasil post-test, komitmen peserta" },
  ] },
];
function kurikulum(hari) {
  const widths = [2300, 3350, 1200, 1500, 1288];
  const rows = hari.modul.map((m) => [
    [`**Modul ${m.no}. ${m.judul}**`, `Tujuan: ${m.tujuan}`, `Materi: ${m.materi.map((t, i) => (i && /^[A-Z][a-z]/.test(t) && !/^(Root|Quality)/.test(t) ? t[0].toLowerCase() + t.slice(1) : t)).join(", ")}.`],
    m.akt, m.alat, m.ukur, m.out,
  ]);
  return table(widths, ["Topik atau modul", "Aktivitas pembelajaran (andragogi)", "Alat atau media", "Indikator capaian peserta", "Output konkret"], rows, { size: 15 });
}

// ---------- isi ----------
const surat = [
  new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "SURAT PENGANTAR", bold: true, size: 30, color: NAVY })] }),
  p("Nomor: [ISI TIM: nomor surat]", { after: 20 }),
  p("Tanggal: [ISI TIM: tanggal surat]", { after: 200 }),
  p("Kepada Yth.", { after: 20 }),
  p("Manajemen PT Pupuk Iskandar Muda", { after: 20, run: { bold: true } }),
  p("[ISI TIM: nama dan jabatan penerima]", { after: 200 }),
  p("Dengan hormat,"),
  p("Sanggabiz menyampaikan proposal lanjutan Analisis Beban Kerja (ABK) fungsi Umum PT Pupuk Iskandar Muda (PIM), disertai pelatihan tiga hari bagi tim PIM. Proposal ini melanjutkan studi Workforce Analysis PIM yang sudah berjalan. Baseline data, istilah jabatan, logika perhitungan FTE, dan pola temuan sudah tersedia, sehingga pekerjaan lanjutan dapat langsung menghasilkan keputusan kebutuhan tenaga kerja."),
  p("Kegiatan ini mengukur kebutuhan dan distribusi beban kerja 112 posisi di Direktorat Utama (26 posisi) dan Direktorat Keuangan dan Umum (86 posisi). Hasilnya berupa analisis dan rekomendasi perbaikan pada tiga sisi: orang (people), proses (process), dan perangkat (platform atau tools)."),
  p("Kami memahami bahwa PIM adalah industri proses dengan tuntutan keselamatan, kepatuhan, kesinambungan operasi, dan koordinasi lintas fungsi yang tinggi. Karena itu setiap angka kami dasarkan pada bukti dari lebih dari satu sumber, dikerjakan dengan pendampingan tenaga senior, dan divalidasi bersama pemilik proses. Rekomendasi disahkan melalui forum manajemen PIM agar dapat dipertanggungjawabkan dan dijalankan."),
  p("Hormat kami,", { before: 160, after: 700 }),
  p("[ISI TIM: tanda tangan, nama, dan jabatan penandatangan]", { after: 20 }),
  p("Sanggabiz", { run: { bold: true } }),
];

const daftarIsi = [
  h1("DAFTAR ISI"),
  ...["Ringkasan Eksekutif", "1. Konteks dan Urgensi", "2. Tujuan, Prinsip, dan Indikator Keberhasilan", "3. Ruang Lingkup", "4. Analisis Beban Kerja Fungsi Umum", "5. Metode Pelaksanaan", "6. Rencana Kerja dan Jadwal", "7. Data, Wawancara, dan Workshop", "8. Pelatihan dan Transfer Kemampuan ABK", "9. Luaran dan Kriteria Penerimaan", "10. Tim, Tata Kelola, dan Quality Assurance", "11. Risiko Pelaksanaan dan Mitigasi", "12. Anggaran", "13. Asumsi, Batasan, dan Ketentuan", "14. Profil Sanggabiz", "Lampiran A. Struktur Workbook ABK", "Lampiran B. Scorecard Manfaat", "Lampiran C. Rujukan"]
    .map((t) => new Paragraph({ spacing: { after: 90 }, border: { bottom: { style: BorderStyle.DOTTED, size: 4, color: LINE, space: 3 } }, children: [new TextRun({ text: t, size: 20 })] })),
];

const ringkasan = [
  h1("RINGKASAN EKSEKUTIF"),
  callout("STATUS PROPOSAL", "Proposal ini melanjutkan studi Workforce Analysis PIM yang sudah ada. Data baseline dipakai kembali selama masih berlaku. Pengumpulan ulang hanya untuk data yang belum lengkap, posisi baru, perubahan struktur, dan verifikasi."),
  spacer(140),
  p("Studi Workforce Analysis PIM sebelumnya menemukan ketidakseimbangan beban kerja: 47 persen posisi berada pada kategori stretched dan 18 persen pada kategori overload. Di level manajerial, 91 persen Manager dan 86 persen Vice President mengalami kelebihan beban [1]. Temuan ini menunjukkan bahwa persoalan PIM bukan hanya jumlah tenaga kerja yang kurang atau lebih. Pembagian pekerjaan, tingkat delegasi, desain keputusan, dan kepemilikan proses lintas unit ikut menentukan."),
  p("Skala PIM mempertegas kebutuhan itu. Informasi publik PIM mencatat 628 pekerja dan Direksi yang terdiri atas Direktur Utama, Direktur Operasi dan Produksi, Direktur Keuangan dan Umum, serta Direktur Manajemen Risiko [2][3]. Nilai ekonomi yang dihasilkan PIM pada 2024 dilaporkan sekitar Rp 5,75 triliun [4]. Organisasi dengan skala ini memerlukan satu acuan yang sama untuk kebutuhan FTE, kepemilikan proses, dan hak keputusan."),
  spacer(60),
  callout("REKOMENDASI KAMI", "Program terpadu 16 sampai 18 minggu: (1) ABK untuk 112 posisi di Direktorat Utama dan Direktorat Keuangan dan Umum, (2) validasi berjenjang bersama pemilik proses, dan (3) pelatihan tiga hari secara luring di Medan untuk tim PIM, lengkap dengan perangkat kerja yang diserahterimakan. Anggaran Rp 165.000.000, sudah termasuk pajak."),
  spacer(140),
  table([1700, 4300, 3638], ["Elemen", "Yang kami kerjakan", "Nilai bagi PIM"], [
    ["Analisis beban kerja", "Perhitungan FTE, BEI, workload heatmap, kesenjangan tenaga kerja, dan usulan perancangan ulang peran", "Dasar keputusan kebutuhan dan pembagian ulang SDM"],
    ["Pemetaan RACI", "RACI untuk posisi dalam lingkup, disusun dari SOP yang sudah ada", "Pembagian peran yang jelas tanpa tumpang tindih"],
    ["Uji kapasitas", "Pemeriksaan apakah penugasan R dan A pada RACI masih sesuai kapasitas pelaksana", "RACI yang realistis dan rekomendasi tenaga kerja yang dapat dijalankan"],
    ["Pelatihan dan serah terima", "Pelatihan tiga hari di Medan, workshop validasi, dan perangkat kerja untuk dipakai sendiri oleh PIM", "Keputusan diterima, terdokumentasi, dan dapat diperbarui oleh tim PIM"],
  ], { firstBold: true }),
];

const bab1 = [
  h1("1. Konteks dan Urgensi"),
  h2("1.1 Dari diagnosis beban kerja ke perbaikan cara kerja organisasi"),
  p("Analisis beban kerja menjawab dua pertanyaan: berapa kapasitas yang dibutuhkan dan di mana kesenjangannya. Angka FTE belum menjawab siapa yang mengerjakan, siapa yang memutuskan, siapa yang wajib dimintai pendapat, dan siapa yang cukup diberi tahu. Tanpa jawaban itu, PIM berisiko menambah orang pada proses yang masih tumpang tindih, atau membiarkan beban Manager dan VP tetap tinggi karena delegasi dan hak keputusan belum dibenahi."),
  p("ISO 9001 meminta organisasi menetapkan pemilik proses, tanggung jawab, wewenang, dan informasi terdokumentasi sebagai dasar pendekatan proses. RACI membantu menerapkan prinsip itu pada setiap proses atau SOP, sedangkan ABK memastikan beban yang dibagikan masih masuk akal terhadap kapasitas yang tersedia [5][6]."),
  h2("1.2 Alasan fungsi Umum didahulukan"),
  bullet("Fungsi Umum melayani hampir semua direktorat. Hambatan pada layanan umum memperpanjang waktu proses di banyak unit sekaligus."),
  bullet("Layanan, administrasi, fasilitas, hubungan industrial, pengadaan pendukung, komunikasi korporat, dan keamanan punya permintaan yang naik turun dan banyak pekerjaan mendadak."),
  bullet("Beban administratif yang tidak tertata cenderung naik ke Manager dan VP. Pola ini sejalan dengan overload level manajerial pada studi sebelumnya."),
  bullet("Pengukuran FTE perlu dilengkapi pemilahan aktivitas menjadi value-adding, mandatory-control, necessary-support, dan non-value-adding. Dengan begitu rekomendasi tidak berhenti pada penambahan SDM."),
];

const bab2 = [
  h1("2. Tujuan, Prinsip, dan Indikator Keberhasilan"),
  h2("2.1 Tujuan program"),
  num("Menghitung beban kerja aktual dan kebutuhan FTE per posisi pada fungsi Umum dengan cara yang konsisten dengan baseline studi PIM sebelumnya.", "tuj"),
  num("Mengenali posisi underload, normal, stretched, dan overload beserta akar penyebabnya: permintaan kerja, metode, struktur, kompetensi, sistem, atau tata kelola.", "tuj"),
  num("Menyusun RACI untuk posisi dalam lingkup dari SOP yang sudah ada, dengan satu penanggung jawab akhir (Accountable) per aktivitas kunci.", "tuj"),
  num("Menyusun rekomendasi tenaga kerja, pembagian ulang pekerjaan, perancangan ulang peran, penyederhanaan proses, dan kesiapan digitalisasi, lengkap dengan urutan prioritas.", "tuj"),
  num("Membekali tim Human Capital, tim Tata Kelola atau Mutu, dan pemilik proses dengan kemampuan dan perangkat untuk menjaga ABK dan RACI tetap mutakhir.", "tuj"),
  h2("2.2 Prinsip desain"),
  table([2900, 6738], ["Prinsip", "Penerapan dalam proyek"], [
    ["Satu aktivitas, satu Accountable", "Setiap aktivitas kunci punya tepat satu A. Pengecualian wajib diberi alasan."],
    ["Mengacu struktur resmi", "Nama unit dan jabatan mengikuti struktur resmi terakhir. Jabatan lama dicatat sebagai selisih data."],
    ["Peka terhadap kontrol", "Pemisahan tugas (segregation of duties), kontrol wajib, dan batas persetujuan tidak disederhanakan tanpa kajian risiko."],
    ["Konsultasi seperlunya", "C dan I dibatasi pada pihak yang benar-benar dibutuhkan untuk mutu keputusan atau kepatuhan."],
    ["Bukti di atas persepsi", "Survei, wawancara, workload diary, dokumen, dan data transaksi dibandingkan satu sama lain."],
    ["Disusun bersama, disahkan manajemen", "Kami memfasilitasi. Pemilik proses dan manajemen PIM mengesahkan kepemilikan akhir."],
  ], { firstBold: true }),
  h2("2.3 Indikator keberhasilan"),
  table([2900, 6738], ["Indikator", "Target penerimaan"], [
    ["Cakupan ABK", "100 persen posisi dalam lingkup punya profil beban kerja dan status FTE yang tervalidasi"],
    ["Kualitas akuntabilitas", "Minimal 95 persen aktivitas kunci punya tepat satu A. Pengecualian punya alasan tertulis"],
    ["Kejelasan peran", "Semua konflik atau tumpang tindih peran tercatat, diputuskan, atau punya pemilik tindakan"],
    ["Validasi", "Seluruh unit dalam lingkup mengikuti validasi. Sponsor dan pemilik proses yang ditunjuk mengesahkan"],
    ["Keberlanjutan", "Perangkat kerja, pemilik, jadwal peninjauan, dan prosedur perubahan RACI diserahterimakan"],
  ], { firstBold: true }),
];

const bab3 = [
  h1("3. Ruang Lingkup"),
  table([2900, 6738], ["Dimensi", "Ketentuan"], [
    ["Objek", "112 posisi: 26 posisi di Direktorat Utama dan 86 posisi di Direktorat Keuangan dan Umum. Daftar posisi dikunci saat kick-off."],
    ["Satuan analisis", "Posisi, aktivitas, frekuensi, volume, waktu, FTE"],
    ["Metode", "FTE, workload diary, BEI, wawancara terstruktur, data transaksi"],
    ["Output inti", "Workbook FTE, heatmap, rekomendasi tenaga kerja, perancangan ulang peran, RACI"],
    ["Pelatihan", "Satu angkatan, 3 hari, maksimal 30 peserta, di Medan"],
    ["Uji kapasitas", "Pemeriksaan penugasan R dan A pada RACI terhadap kapasitas pelaksana"],
  ], { firstBold: true }),
  h2("3.1 Yang dikerjakan"),
  bullet("Menelaah struktur organisasi, job description, KPI, daftar pekerja, dan baseline ABK terdahulu."),
  bullet("Mengukur aktivitas berkala, insidental, berbasis proyek, kontrol rutin, dan koordinasi."),
  bullet("Memetakan persetujuan, serah terima pekerjaan, eskalasi, pengecualian, konsultasi wajib, dan aliran informasi."),
  bullet("Memvalidasi hasil per unit, mengalibrasi lintas fungsi, memperoleh pengesahan manajemen, dan menyerahterimakan perangkat kerja."),
  h2("3.2 Di luar lingkup dasar"),
  bullet("Menulis ulang seluruh naskah SOP, flowchart rinci, formulir, atau work instruction."),
  bullet("Audit kepatuhan hukum, audit sertifikasi ISO, dan pendapat hukum."),
  bullet("Penilaian kinerja individu, job grading, desain remunerasi, atau talent assessment."),
  bullet("Pembangunan aplikasi workflow, HRIS, atau dashboard tingkat produksi."),
  bullet("Studi waktu dan gerak 24 jam di seluruh lokasi. Observasi intensif dapat ditambahkan untuk posisi kritis."),
  spacer(100),
  callout("PENGUNCIAN LINGKUP", "RACI disusun dari SOP aktif yang diserahkan PIM untuk 112 posisi dalam lingkup. Dokumen yang duplikat, kedaluwarsa, masih berupa draf, berupa kebijakan tanpa alur aktivitas, atau berupa instruksi kerja yang sangat rinci dicatat dengan status tersendiri dan tidak dihitung sebagai satu SOP."),
];

const bab4 = [
  h1("4. Analisis Beban Kerja Fungsi Umum"),
  h2("4.1 Pertanyaan yang dijawab"),
  bullet("Berapa jam kerja efektif dan kebutuhan FTE ideal untuk setiap posisi?"),
  bullet("Posisi mana yang underload, normal, stretched, atau overload, dan seberapa besar selisih tenaga kerjanya?"),
  bullet("Apakah beban tinggi berasal dari permintaan kerja yang nyata, pekerjaan manual, pengerjaan ulang, lapisan persetujuan, penugasan sementara, atau peran yang tumpang tindih?"),
  bullet("Aktivitas mana yang dapat dihentikan, disederhanakan, diotomasi, dipindahkan, digabung, atau dilatihkan silang (cross-skill)?"),
  h2("4.2 Komponen analisis"),
  table([1900, 4400, 3338], ["Komponen", "Yang kami lakukan", "Output"], [
    ["Daftar posisi", "Mengunci daftar posisi, incumbent, unit, atasan, status, lokasi, dan shift", "Master position list dan catatan data yang kurang"],
    ["Inventaris aktivitas", "Memetakan aktivitas, output, pemicu, volume, frekuensi, waktu, dan puncak musiman", "Kamus aktivitas per posisi"],
    ["Perhitungan FTE", "Menghitung beban berkala dan tidak berkala terhadap jam kerja efektif", "FTE aktual, FTE ideal, selisih, dan confidence level"],
    ["Validasi BEI", "Mewawancarai kejadian kritis, hambatan, pengerjaan ulang, ketergantungan, dan hidden work", "Bukti kualitatif dan catatan akar masalah"],
    ["Pandangan proses", "Mengelompokkan aktivitas: value-adding, control, support, waiting, duplikasi", "Peluang penyederhanaan dan otomasi"],
    ["Skenario tenaga kerja", "Menyusun model pertahankan, bagi ulang, rancang ulang, rekrut, latih ulang", "Rekomendasi tenaga kerja berurutan prioritas"],
    ["Validasi manajemen", "Berkalibrasi dengan pimpinan unit, HR, dan sponsor", "Workbook tervalidasi dan decision log"],
  ], { firstBold: true }),
  h2("4.3 Rumus dan klasifikasi"),
  p("FTE dihitung dari total jam beban kerja tahunan dibagi jam kerja efektif tahunan. Jam kerja efektif dikunci bersama PIM dengan memperhitungkan hari kerja, cuti, hari libur, pelatihan wajib, briefing, dan allowance yang relevan. Agar sejalan dengan studi sebelumnya, ambang awal memakai kategori berikut [1]. Ambang dapat dikalibrasi mengikuti kebijakan PIM."),
  table([1500, 1700, 3000, 3438], ["Status", "Rentang awal", "Tafsiran", "Arah tindak lanjut"], [
    ["Underload", "< 0,85 FTE", "Kapasitas berlebih atau permintaan kerja belum penuh", "Bagi ulang, cross-skill, perkaya peran"],
    ["Normal", "0,85-1,15 FTE", "Beban relatif seimbang", "Pertahankan, pantau puncak beban dan ketergantungan"],
    ["Stretched", "> 1,15-1,30 FTE", "Beban tinggi, rentan saat puncak atau saat ada yang absen", "Sederhanakan, siapkan pengganti, atur ulang prioritas"],
    ["Overload", "> 1,30 FTE", "Beban tidak berkelanjutan", "Rancang ulang, otomasi, bagi ulang, atau rekrut"],
  ], { firstBold: true }),
  h2("4.4 Output keputusan"),
  bullet("Workload heatmap per posisi, level, dan unit."),
  bullet("Selisih kebutuhan SDM: kondisi saat ini, kebutuhan hasil hitung, kebutuhan dibulatkan, surplus atau kekurangan, dan confidence level."),
  bullet("Matriks prioritas dari tingkat FTE, kritisnya bisnis, dan dampak pada keselamatan atau kepatuhan."),
  bullet("Rekomendasi per posisi: pertahankan, bagi ulang, rancang ulang, rekrut, latih ulang, otomasi, atau pantau."),
  bullet("Daftar hidden work dan penugasan tidak resmi yang tidak tercatat di job description."),
  bullet("Rencana tindak tenaga kerja untuk 0-3 bulan, 3-6 bulan, dan 6-12 bulan."),
];

const bab5 = [
  h1("5. Metode Pelaksanaan"),
  p("Kami memakai metode yang lazim di bidangnya dan sesuai dengan masalah PIM. ABK mengikuti logika Pedoman Analisis Jabatan dan Analisis Beban Kerja pada Permen PANRB Nomor 1 Tahun 2020, yang kami pakai sebagai acuan metode dan tidak sebagai aturan yang otomatis mengikat PIM [7]. Pelatihan mengikuti ADDIE dan prinsip belajar orang dewasa. Angka-angka diuji dari lebih dari satu sumber agar hasil tidak bergantung pada satu laporan mandiri."),
  table([2300, 3700, 3638], ["Metode", "Alasan dipilih", "Muncul di bagian"], [
    ["Pedoman ABK Permen PANRB 1/2020 sebagai acuan [7]", "Memberi kerangka baku perhitungan beban kerja dan kebutuhan pegawai yang dikenal pegawai BUMN dan pemerintah", "Bab 4 dan Bab 5.2"],
    ["Triangulasi bukti", "Mengurangi bias laporan mandiri", "Bab 5.3"],
    ["ADDIE [9]", "Memberi alur pelatihan dari analisis kebutuhan sampai evaluasi", "Bab 8.1"],
    ["Andragogi (Knowles) [10]", "Peserta adalah pegawai dewasa yang belajar dari kasus kerjanya sendiri", "Bab 8.3"],
    ["Taksonomi Bloom revisi [11]", "Membuat tujuan pembelajaran terukur dengan kata kerja yang jelas", "Bab 8.2 dan 8.4"],
    ["Kirkpatrick (4 level) [12]", "PIM meminta bukti bahwa pelatihan mengubah cara kerja", "Bab 8.5"],
    ["Backward design", "Pelatihan dimulai dari hasil yang diharapkan dan bukti keberhasilan, lalu materi dirancang dari pemetaan awal peserta", "Bab 8.1"],
    ["PDCA", "Menjaga ABK dan RACI tetap mutakhir setelah proyek selesai", "Lampiran B"],
  ], { firstBold: true }),
  h2("5.1 Fase pelaksanaan"),
  table([2300, 1000, 3900, 2438], ["Fase", "Minggu", "Kegiatan utama", "Gerbang atau output"], [
    ["0. Mobilisasi", "1-2", "Kick-off, data room, penguncian lingkup, daftar posisi, pendataan SOP, pemetaan pemangku kepentingan", "Laporan awal dan baseline terkonfirmasi"],
    ["1. Diagnosis", "2-6", "Telaah data, survei dan workload diary, wawancara, telaah proses, hipotesis awal", "Kelengkapan data minimal 90 persen dan daftar isu v1"],
    ["2. Analisis dan draf", "5-10", "Perhitungan FTE, heatmap, penguraian proses, normalisasi peran, draf RACI", "Draf ABK dan cakupan RACI minimal 80 persen"],
    ["3. Validasi", "9-14", "Workshop unit, telaah pemilik proses, kalibrasi lintas fungsi, penyelesaian data menyimpang", "Matriks tervalidasi dan decision log"],
    ["4. Integrasi dan rekomendasi", "13-16", "RACI yang memperhitungkan kapasitas, skenario tenaga kerja, prioritas penyederhanaan, peta jalan", "Rekomendasi terintegrasi"],
    ["5. Pengesahan dan serah terima", "16-18", "Telaah eksekutif, revisi, pengesahan akhir, pelatihan perangkat kerja, penutupan", "Paket akhir dan serah terima"],
  ], { firstBold: true }),
  h2("5.2 Pengujian rekomendasi"),
  p("Setiap rekomendasi posisi atau proses diuji dengan lima pertanyaan:", { after: 60 }),
  num("Apakah aktivitas wajib dilakukan karena produksi, keselamatan, regulasi, kontrak, atau tata kelola?", "uji"),
  num("Apakah aktivitas memberi nilai, mengendalikan risiko, atau hanya muncul karena pengerjaan ulang, menunggu, dan duplikasi?", "uji"),
  num("Apakah peran yang menjalankannya adalah peran yang paling tepat menurut RACI dan kompetensinya?", "uji"),
  num("Apakah kapasitas FTE peran itu cukup setelah semua penugasan dihitung?", "uji"),
  num("Apakah jalan keluar terbaik berupa penyederhanaan, pembagian ulang, perancangan ulang, otomasi, pelatihan ulang, atau penambahan SDM?", "uji"),
  h2("5.3 Triangulasi bukti"),
  table([2300, 3900, 3438], ["Sumber bukti", "Dipakai untuk", "Pengendali bias"], [
    ["Dokumen formal", "Struktur, JD, KPI, SOP, batas wewenang, kebijakan, formulir, titik kontrol", "Pemeriksaan versi dan status dokumen"],
    ["Survei dan workload diary", "Volume, frekuensi, waktu, musiman, hidden work", "Uji data menyimpang dan kalibrasi atasan"],
    ["BEI dan wawancara terstruktur", "Kejadian kritis, hambatan, ketergantungan, praktik nyata", "Pemicu bukti dan perbandingan antarwawancara"],
    ["Data transaksi", "Jumlah tiket, PO, invoice, surat, permintaan, work order, persetujuan", "Periode sampel dan rekonsiliasi"],
    ["Workshop", "Pembagian peran, penyelesaian konflik, penanganan pengecualian", "Decision log yang difasilitasi dan pengesahan"],
    ["Baseline ABK lama", "Kesinambungan, pembanding, dan penyegaran terarah", "Validasi ulang bila struktur atau permintaan kerja berubah"],
  ], { firstBold: true }),
];

const bab6 = [
  h1("6. Rencana Kerja dan Jadwal"),
  callout("DURASI PROGRAM", "Durasi dasar program adalah **16 sampai 18 minggu**. Jadwal bergantung pada ketersediaan dokumen, respons survei, dan kecepatan penjadwalan workshop. Pelatihan tiga hari berlangsung pada November 2026 di Medan. Kami mengusulkan pelatihan selesai sebelum workshop validasi unit dimulai pada minggu ke-9."),
  spacer(140),
  gantt(),
  p("Keterangan: V menandai periode dua minggu saat kegiatan berjalan. M1-2 berarti minggu pertama dan kedua.", { before: 100, run: { size: 16, color: MUTED } }),
  h2("6.1 Gerbang manajemen"),
  table([2900, 1700, 5038], ["Gerbang", "Waktu", "Keputusan yang diperlukan"], [
    ["Gerbang 1: Lingkup dan data", "Akhir minggu 2", "Daftar posisi, baseline SOP, pemilik, lokasi, dan asumsi jam kerja"],
    ["Gerbang 2: Temuan dan konflik", "Minggu 10-12", "FTE menyimpang, aktivitas tanpa A atau dengan A ganda, tabrakan peran, dan kebijakan penyelesaian konflik"],
    ["Gerbang 3: Rekomendasi dan pengesahan", "Minggu 16-18", "Tindakan tenaga kerja, RACI akhir, pemilik pelaksanaan, tanggal berlaku"],
  ], { firstBold: true }),
];

const bab7 = [
  h1("7. Data, Wawancara, dan Workshop"),
  h2("7.1 Permintaan data awal"),
  table([1800, 5800, 2038], ["Kelompok", "Dokumen atau data", "Waktu"], [
    ["Organisasi", "Struktur terakhir, daftar posisi dan incumbent, garis pelaporan, lokasi dan shift, posisi kosong", "Hari 1-5"],
    ["Peran", "JD, KPI, profil jabatan, profil kompetensi, memo penugasan, keanggotaan komite", "Hari 1-10"],
    ["Beban kerja", "Baseline ABK, data volume dan transaksi, antrean pekerjaan, lembur, kalender kerja, tingkat layanan", "Hari 1-10"],
    ["Proses", "Daftar SOP, SOP aktif, kebijakan, instruksi kerja, formulir, peta proses", "Hari 1-10"],
    ["Tata kelola", "Delegation of Authority, matriks persetujuan, persyaratan GCG, risiko, dan kontrol, temuan audit", "Hari 1-12"],
    ["Sistem", "Hak akses aplikasi, tangkapan layar alur kerja, log tiket atau transaksi yang tersedia", "Hari 5-15"],
  ], { firstBold: true }),
  h2("7.2 Pelibatan pemangku kepentingan"),
  table([2700, 2200, 1500, 3238], ["Pemangku kepentingan", "Bentuk", "Perkiraan", "Tujuan"], [
    ["Direksi atau sponsor", "Kick-off dan tiga gerbang", "4 sesi", "Arah, penyelesaian konflik, pengesahan"],
    ["HR atau Organisasi", "Sesi kerja mingguan", "8-10 sesi", "Data, kalibrasi ABK, kebijakan tenaga kerja"],
    ["Mutu, GCG, Risiko, Audit Internal", "Workshop terstruktur", "4-6 sesi", "Status SOP, kontrol, kepatuhan, pemisahan tugas"],
    ["VP dan Manager fungsi Umum", "Wawancara dan validasi", "12-18 sesi", "Permintaan kerja, peran, beban kerja, rekomendasi"],
    ["Pemilik proses", "Workshop per kelompok proses", "12-20 sesi", "Draf RACI dan serah terima lintas fungsi"],
    ["Incumbent terpilih", "Wawancara BEI dan beban kerja", "25-40 orang", "Pekerjaan nyata, hidden work, penentu waktu"],
  ], { firstBold: true }),
  spacer(100),
  callout("CARA PELAKSANAAN", "Analisis dan sesi kerja dilakukan daring lebih dahulu. Kami merencanakan paling banyak tiga kunjungan langsung: kick-off dan validasi data, kalibrasi lintas fungsi, dan pengesahan eksekutif. Biaya perjalanan untuk tiga kunjungan itu sudah termasuk dalam anggaran."),
];

const bab8 = [
  h1("8. Pelatihan dan Transfer Kemampuan ABK"),
  p("Pelatihan berlangsung tiga hari berturut-turut pada November 2026 di Medan, dengan 24 jam pelajaran (JPL) atau 18 jam efektif. Satu JPL setara 45 menit. Baseline pelaksanaan maksimal 30 peserta dengan satu Lead Facilitator dan satu Co-Facilitator."),
  table([2400, 2200, 2400, 2638], ["Durasi", "Lokasi", "Peserta", "Format"], [
    ["3 hari, 24 JPL", "Medan", "Maksimal 30 orang", "Tatap muka, praktik, diskusi kelompok"],
  ], { align: [AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] }),
  h2("8.1 Alur pelatihan dan pemetaan awal"),
  p("Alur mengikuti ADDIE. Pada tahap analisis, kami memetakan peserta lebih dulu lewat konfirmasi kebutuhan, profil peserta, dan pre-test, sesuai prinsip backward design. Materi hari kedua dan ketiga disesuaikan dengan hasil pre-test."),
  table([1900, 2300, 5438], ["Tahap ADDIE", "Waktu", "Kegiatan"], [
    ["Analisis", "H-14 sampai H-7", "Konfirmasi kebutuhan, profil peserta, kasus yang dipakai, bahan bacaan, dan pre-test baseline"],
    ["Rancangan dan pengembangan", "H-14 sampai H-3", "Rencana pelajaran akhir, materi fasilitator, workbook peserta, template, dan paket kasus"],
    ["Pelaksanaan", "3 hari, November 2026", "24 JPL sesuai kurikulum di Medan"],
    ["Evaluasi", "Hari 1 sampai D+90", "Evaluasi harian, post-test, clinic D+30, instrumen Level 4 pada D+60 sampai D+90"],
  ], { firstBold: true }),
  h2("8.2 Tujuan pembelajaran"),
  p("Setelah pelatihan, peserta mampu:", { after: 60 }),
  num("Menjelaskan prinsip ABK, FTE, jam kerja efektif, volume, frekuensi, waktu penyelesaian, dan asumsi annualisasi secara konsisten.", "tpl"),
  num("Menggunakan activity inventory, workload worksheet, interview guide, dan evidence checklist untuk mengumpulkan data yang dapat diaudit.", "tpl"),
  num("Mewawancarai secara terstruktur untuk menggali hidden work, musiman, pengerjaan ulang, waktu menunggu, pekerjaan mendadak, dan aktivitas pengendalian.", "tpl"),
  num("Memeriksa kualitas data lewat uji kewajaran dan uji data menyimpang, serta merekonsiliasi laporan mandiri, dokumen, dan data transaksi.", "tpl"),
  num("Menganalisis heatmap, menelusuri akar penyebab, lalu memilih tindak lanjut: penyederhanaan, pembagian ulang, perancangan ulang, otomasi, pelatihan ulang, atau penambahan kapasitas.", "tpl"),
  num("Menguji apakah penugasan Responsible dan Accountable pada RACI menimbulkan overload, duplikasi, atau hambatan baru.", "tpl"),
  h2("8.3 Peserta dan pendekatan belajar"),
  p("Peserta diprioritaskan dari Tim Human Capital atau Organization Development, perwakilan supervisor dan manajer fungsi Umum, dan pemilik proses lintas direktorat yang terlibat dalam validasi ABK dan RACI. Porsi belajar sekitar 25 persen konsep, 55 persen praktik dan simulasi, dan 20 persen review dan rencana tindak. Pendekatan andragogi dipakai di setiap modul: peserta bekerja dengan kasus dari unitnya sendiri, belajar dengan memecahkan masalah, dan fasilitator berperan memandu."),
  h2("8.4 Kurikulum tiga hari"),
  p("Satu modul berdurasi 2 JPL atau 90 menit. Kata kerja pada tujuan mengikuti taksonomi Bloom revisi. Rincian waktu di kolom aktivitas dihitung per modul.", { after: 80 }),
  table([1500, 1900, 6238], ["Hari", "Waktu", "Kegiatan"], [
    ["Hari 1", "08.30 - 09.00", "Registrasi dan pembukaan (pre-test sudah dikerjakan sebelum hari pertama)"],
    ["", "09.00 - 10.30", "Modul 1"],
    ["", "10.45 - 12.15", "Modul 2"],
    ["", "13.15 - 14.45", "Modul 3"],
    ["", "15.00 - 16.30", "Modul 4"],
    ["Hari 2", "09.00 - 10.30", "Modul 5"],
    ["", "10.45 - 12.15", "Modul 6"],
    ["", "13.15 - 14.45", "Modul 7"],
    ["", "15.00 - 16.30", "Modul 8"],
    ["Hari 3", "09.00 - 10.30", "Modul 9"],
    ["", "10.45 - 12.15", "Modul 10"],
    ["", "13.15 - 14.45", "Modul 11"],
    ["", "15.00 - 16.30", "Modul 12, termasuk post-test dan penutupan"],
  ], { align: [AlignmentType.CENTER, AlignmentType.CENTER], firstBold: true }),
  p("Waktu istirahat 15 menit pada pagi dan sore hari serta istirahat makan siang pukul 12.15 sampai 13.15 WIB. Jam pelaksanaan dapat disesuaikan dengan jam kerja PIM.", { before: 100, run: { size: 17, color: MUTED } }),
  ...HARI.flatMap((h) => [h3(h.judul), kurikulum(h)]),
  h2("8.5 Evaluasi pembelajaran dengan Kirkpatrick"),
  p("Evaluasi tidak berhenti pada kepuasan peserta. Kami menilai empat tingkat supaya PIM dapat melihat apakah pemahaman berubah menjadi perilaku kerja dan perbaikan mutu ABK. Target Level 3 dan Level 4 adalah indikator bersama, karena juga dipengaruhi ketersediaan data, dukungan atasan, dan tindak lanjut yang dijalankan PIM."),
  table([1300, 2300, 2800, 3238], ["Level", "Pertanyaan evaluasi", "Instrumen dan waktu", "Indikator target"], [
    ["1 Reaksi", "Apakah materi relevan, jelas, dapat diterapkan, dan difasilitasi dengan baik?", "Pulse survey harian dan evaluasi akhir Hari 3", "Rata-rata minimal 4,2 dari 5 dan minimal 85 persen peserta menilai baik atau sangat baik"],
    ["2 Pembelajaran", "Apakah peserta memahami konsep dan mampu memakai instrumen dasar?", "Pre-test, post-test, latihan FTE, role play wawancara, kasus QA", "Skor post-test minimal 75 dan peningkatan rata-rata minimal 20 persen dari baseline"],
    ["3 Perilaku", "Apakah peserta memakai template, teknik wawancara, dan QA checklist dalam pekerjaan?", "Daftar cek atasan, bukti penggunaan, dan implementation clinic daring pada D+30", "Minimal 75 persen peserta sasaran memakainya pada sedikitnya satu aktivitas ABK atau validasi"],
    ["4 Hasil", "Apakah kualitas data dan kecepatan tindak lanjut membaik?", "Tinjauan indikator pada D+60 sampai D+90 dengan perangkat yang diserahterimakan", "Kelengkapan data minimal 90 persen, isu berulang turun, dan rencana tindak punya pemilik dan target waktu"],
  ], { firstBold: true, size: 16 }),
  h3("Rangkaian evaluasi dan tindak lanjut"),
  table([2300, 7338], ["Waktu", "Kegiatan"], [
    ["Sebelum Hari 1", "Konfirmasi kebutuhan, profil peserta, bahan bacaan, dan pre-test baseline"],
    ["Hari 1 sampai 3", "Observasi praktik, penilaian latihan, pulse survey harian, post-test, dan rencana tindak"],
    ["D+30", "Satu implementation clinic daring maksimal dua jam untuk hambatan penerapan dan tinjauan Level 3"],
    ["D+60 sampai D+90", "Instrumen Level 4 dan panduan tinjauan manajemen diserahkan agar PIM dapat menilai hasil awal secara konsisten"],
  ], { firstBold: true }),
  h2("8.6 Hasil yang dibawa pulang"),
  table([4400, 2900, 2338], ["Hasil", "Dihasilkan di", "Pemegang"], [
    ["Peta isu awal unit", "Modul 1", "Kelompok"],
    ["Lembar latihan hitung FTE", "Modul 2", "Peserta"],
    ["Analisis kasus dan pohon akar masalah", "Modul 3", "Kelompok"],
    ["Daftar hipotesis dan pertanyaan validasi", "Modul 4", "Kelompok"],
    ["Activity inventory satu posisi", "Modul 5", "Peserta"],
    ["Interview guide terisi", "Modul 6", "Pasangan"],
    ["Rubrik observasi dan umpan balik", "Modul 7", "Peserta"],
    ["Worksheet praktik dan issue log", "Modul 8", "Kelompok"],
    ["QA checklist dan data issue log", "Modul 9", "Kelompok"],
    ["Sampel tervalidasi dan decision log", "Modul 10", "Kelompok"],
    ["Prioritized action matrix", "Modul 11", "Kelompok"],
    ["Rencana 30-60-90 hari unit, hasil post-test, dan komitmen", "Modul 12", "Peserta"],
    ["Materi fasilitator, workbook digital, rencana pelajaran akhir", "Disiapkan sebelum pelatihan", "PIM"],
    ["Daftar hadir, sertifikat elektronik, dan training report (maksimal 7 hari kerja setelah pelatihan)", "Setelah pelatihan", "PIM"],
    ["Instrumen Level 3 dan Level 4 serta panduan tinjauan manajemen", "Clinic D+30 dan D+60 sampai D+90", "PIM"],
  ], { align: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER] }),
  h2("8.7 Kebutuhan tempat"),
  p("PIM menyediakan ruang pelatihan di Medan. Kami memerlukan ruang yang memuat 30 peserta dalam lima meja kelompok (susunan cluster), dengan layar atau proyektor, tata suara, koneksi internet stabil, dan ruang cukup untuk papan flipchart di dinding. Ruang sebaiknya berada di lokasi yang mudah dijangkau peserta dari kantor mereka. [ISI TIM: nama tempat, alamat, nama ruang, ukuran ruang, dan kapasitas bila PIM sudah menetapkan tempatnya]"),
];

const bab9 = [
  h1("9. Luaran dan Kriteria Penerimaan"),
  table([600, 3000, 1700, 4338], ["No.", "Luaran", "Format", "Kriteria penerimaan"], [
    ["1", "Laporan awal dan baseline lingkup", "DOCX atau PDF", "Posisi, SOP yang dipakai, pemilik, metode, jadwal, dan data yang kurang disetujui"],
    ["2", "Master posisi dan inventaris aktivitas", "XLSX", "100 persen posisi dalam lingkup, nomenklatur dan garis pelaporan tervalidasi"],
    ["3", "Workbook perhitungan ABK", "XLSX", "Rumus dapat diaudit. Volume, frekuensi, waktu, FTE, selisih, dan confidence level tersedia"],
    ["4", "Laporan ABK dan heatmap", "DOCX, PDF, atau PPTX", "Temuan, akar masalah, skenario tenaga kerja, dan tindakan prioritas"],
    ["5", "RACI untuk posisi dalam lingkup", "XLSX", "Disusun dari SOP aktif. Minimal 95 persen aktivitas kunci punya tepat satu A"],
    ["6", "Catatan selisih, konflik, dan perubahan", "XLSX", "Isu, dampak, keputusan, pemilik, target tanggal, dan status tersedia"],
    ["7", "Rekomendasi terintegrasi dan peta jalan", "DOCX, PDF, atau PPTX", "ABK dan RACI terhubung. Ada kemenangan cepat dan peta jalan 12 bulan"],
    ["8", "Pelatihan dan transfer kemampuan", "Pelatihan tiga hari, perangkat kerja, laporan, clinic D+30", "Pelaksanaan sesuai kurikulum, seluruh perangkat belajar tersedia, hasil Level 1 dan 2 dilaporkan, instrumen Level 3 dan 4 diserahterimakan"],
  ], { align: [AlignmentType.CENTER], split: true }),
  h2("9.1 Prosedur penerimaan"),
  bullet("Setiap luaran punya satu jendela telaah terkonsolidasi paling lama lima hari kerja."),
  bullet("PIM menunjuk satu konsolidator agar komentar tidak saling bertentangan."),
  bullet("Komentar dikelompokkan menjadi koreksi, klarifikasi, perlu keputusan, atau perubahan lingkup."),
  bullet("Luaran dianggap diterima bila komentar material sudah ditutup dan sponsor atau pemilik proses mengesahkan, atau tidak ada komentar baru dalam periode telaah yang disepakati."),
  bullet("Perubahan struktur, tambahan posisi atau SOP, atau permintaan menulis ulang SOP setelah baseline dikunci diproses lewat permintaan perubahan."),
];

const bab10 = [
  h1("10. Tim, Tata Kelola, dan Quality Assurance"),
  h2("10.1 Susunan tim"),
  table([2900, 4600, 2138], ["Peran", "Tanggung jawab", "Keterlibatan"], [
    ["Direktur Proyek (Project Director)", "Arahan eksekutif, keselarasan pemangku kepentingan, gerbang mutu, rekomendasi akhir", "Paruh waktu sepanjang proyek"],
    ["Manajer Proyek (Engagement Manager)", "Kepemimpinan harian, integrasi ABK dan RACI, penyelesaian isu, koordinasi dengan PIM", "Intensif, memimpin rapat mingguan"],
    ["Pemimpin ABK dan Organisasi", "Metode FTE, BEI, analisis beban kerja, skenario tenaga kerja, perancangan ulang peran", "Pemimpin analisis beban kerja"],
    ["Pemimpin Tata Kelola Proses dan RACI", "Penelaahan SOP, desain RACI, kontrol, dan hak keputusan", "Pemimpin pemetaan RACI"],
    ["Konsultan senior", "Wawancara, workshop, analisis, penyelesaian isu lintas fungsi", "Penuh atau paruh waktu sesuai fase"],
    ["Analis 1-2 orang", "Pembersihan data, workbook, draf RACI, visualisasi, dokumentasi", "Penuh waktu"],
    ["PMO dan QA", "Jadwal, pengendalian dokumen, decision log, versi, pelacakan penerimaan", "Paruh waktu, puncak di gerbang"],
    ["Lead Facilitator dan Co-Facilitator pelatihan", "Menyiapkan dan memandu pelatihan tiga hari, menyusun laporan, memandu clinic D+30", "Sesuai jadwal pelatihan"],
  ], { firstBold: true }),
  p("[ISI TIM: nama, jabatan, dan pengalaman anggota tim kunci. Lampirkan CV ringkas]", { before: 100 }),
  h2("10.2 Forum tata kelola"),
  table([2400, 3000, 1800, 2438], ["Forum", "Peserta", "Frekuensi", "Output"], [
    ["Steering Committee", "Sponsor, Direksi terkait, pimpinan HR, Direktur Proyek", "3 gerbang dan bila ada pengecualian", "Keputusan dan pengesahan"],
    ["Rapat manajemen proyek", "PM PIM, Manajer Proyek, pemimpin tiap bagian", "Mingguan", "Kemajuan, risiko, data kurang, langkah berikut"],
    ["Sesi kerja", "Pemilik materi dan konsultan", "1-3 kali per minggu", "Draf analisis dan penutupan isu"],
    ["Telaah mutu", "Direktur Proyek dan penelaah independen", "Sebelum setiap gerbang", "Pemeriksaan metode, konsistensi, dan bukti"],
  ], { firstBold: true }),
  h2("10.3 Quality assurance"),
  bullet("Pemeriksaan konsistensi metode untuk rumus FTE, pembulatan, jam kerja efektif, dan klasifikasi status."),
  bullet("Pemeriksaan logika RACI: tanpa A, A ganda, tanpa R, R/C/I berlebihan, konflik pemisahan tugas, dan nama yang tidak cocok."),
  bullet("Keterlacakan bukti: setiap rekomendasi material ditautkan pada data, wawancara, SOP, atau decision log."),
  bullet("Pemeriksaan konsistensi antar-SOP untuk peran dan titik kontrol yang berulang."),
  bullet("Penyusunan laporan eksekutif yang membedakan fakta, hipotesis, keputusan, dan rekomendasi."),
  bullet("Pengendalian versi dan pembatasan akses pada data tenaga kerja dan dokumen internal."),
];

const bab11 = [
  h1("11. Risiko Pelaksanaan dan Mitigasi"),
  table([2300, 1800, 3300, 2238], ["Risiko", "Dampak", "Mitigasi", "Pemilik bersama"], [
    ["Daftar SOP tidak lengkap atau ada versi ganda", "Cakupan dan beban kerja tidak pasti", "Pendataan SOP, penandaan versi, penguncian baseline", "Mutu atau GCG PIM dan PM"],
    ["Laporan beban kerja terlalu tinggi atau terlalu rendah", "FTE bias", "Data transaksi, uji data menyimpang, kalibrasi atasan", "HR dan Pemimpin ABK"],
    ["Nama jabatan tidak konsisten", "RACI sulit disahkan", "Normalisasi peran dan tabel padanan jabatan lama", "HR Organisasi"],
    ["Konflik kepemilikan antarunit", "Workshop berulang", "Decision log, kertas isu, eskalasi ke gerbang", "Sponsor dan Direktur Proyek"],
    ["Manager menjadi A atau R bawaan", "RACI memperkuat overload", "Pemeriksaan kapasitas dan perancangan ulang delegasi", "Direktorat terkait"],
    ["Perubahan struktur saat proyek", "Pengerjaan ulang", "Tanggal batas dan prosedur pengendalian perubahan", "Steering Committee"],
    ["Data atau jadwal terlambat", "Jadwal bergeser", "Pemilik data, pemantauan umur data mingguan, penjadwalan ulang milestone", "PM PIM dan Manajer Proyek"],
    ["Peserta pelatihan tidak sesuai profil atau berhalangan", "Kemampuan tidak sampai ke pemakai", "Konfirmasi peserta sebelum H-7 dan peserta pengganti dari daftar cadangan", "HR PIM dan Sanggabiz"],
  ], { firstBold: true }),
  h2("11.1 Kerahasiaan data"),
  p("Data individu, daftar jabatan, workload diary, dan dokumen SOP diperlakukan sebagai informasi rahasia. Akses diberikan menurut kebutuhan. Data mentah tidak dimuat dalam laporan eksekutif kecuali diperlukan. File kerja final diserahkan kepada PIM, sedangkan salinan kerja Sanggabiz dihapus atau diarsipkan sesuai masa simpan yang disepakati."),
];

const bab12 = [
  h1("12. Anggaran"),
  p("Kami mengajukan satu paket untuk 112 posisi, tiga hari pelatihan, dan seluruh luaran pada Bab 9. Nilai total sudah termasuk pajak."),
  table([2300, 4938, 2400], ["Komponen", "Cakupan", "Nilai"], [
    ["Jasa konsultansi", "ABK untuk lingkup yang disepakati, RACI dari SOP yang sudah ada, validasi, rekomendasi, dan seluruh luaran akhir", { text: fmt(KONSULTANSI), bold: false }],
    ["Jasa pelatihan", "Kurikulum, tiga hari luring di Medan, dua fasilitator, perangkat kerja, evaluasi Kirkpatrick, laporan, dan clinic D+30", { text: fmt(PELATIHAN), bold: false }],
    [{ text: "Total (sudah termasuk pajak)", bold: true, bg: "F2F5FA" }, { text: "Satu rangkaian program ABK, RACI, dan transfer kemampuan", bold: true, bg: "F2F5FA" }, { text: fmt(TOTAL), bold: true, bg: "F2F5FA" }],
  ], { align: [AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.RIGHT], firstBold: true }),
  p("Harga tidak termasuk tempat, konsumsi, dan transportasi peserta selama pelatihan.", { before: 100, run: { size: 17, color: MUTED } }),
  h2("12.1 Ketentuan biaya dan pelaksanaan"),
  bullet("Nilai di atas sudah termasuk pajak. Bila ada pemotongan PPh, bukti potong disampaikan kepada Sanggabiz."),
  bullet("Komponen pelatihan mencakup honor fasilitator, pengembangan materi, perjalanan dan akomodasi dua fasilitator untuk satu angkatan tiga hari di Medan, serta luaran pada Bab 8.6."),
  bullet("PIM menyediakan ruang pelatihan, layar atau proyektor, tata suara, koneksi internet, konsumsi peserta, serta perjalanan dan akomodasi peserta."),
  bullet("Baseline peserta maksimal 30 orang. Tambahan peserta, hari pelatihan, angkatan, lokasi, pencetakan khusus, atau tindak lanjut di luar D+30 diproses lewat persetujuan perubahan lingkup."),
  bullet("Hari 2 berisi praktik pengumpulan data dan wawancara dengan sampel kasus terpilih. Kegiatan ini tidak menggantikan pengumpulan data, validasi, atau wawancara proyek."),
  bullet("Pelatihan tidak mencakup penyusunan, penulisan ulang, atau revisi SOP. SOP yang sudah ada dipakai sebagai acuan pemetaan aktivitas dan RACI."),
  h2("12.2 Termin pembayaran"),
  table([2300, 2400, 4938], ["Komponen", "Termin", "Pemicu"], [
    ["Konsultansi", "25 / 35 / 30 / 10 persen", "Mobilisasi, draf analisis, validasi selesai, luaran akhir diserahkan"],
    ["Pelatihan", "50 / 50 persen", "Kurikulum dan jadwal final, lalu pelatihan selesai dan training report diserahkan"],
  ], { firstBold: true }),
  h2("12.3 Tarif perubahan lingkup"),
  table([3200, 2800, 3638], ["Perubahan", "Tarif", "Ketentuan"], [
    ["Tambahan posisi ABK di atas 112", "Rp 1.750.000 per posisi", "Bila instrumen dan metode sama. Tidak termasuk observasi khusus"],
    ["Observasi langsung tambahan", "Biaya riil ditambah hari kerja profesional yang disepakati", "Disetujui tertulis sebelum perjalanan"],
    ["Penulisan ulang SOP", "Dihitung terpisah", "Di luar pemetaan RACI dasar"],
  ], { firstBold: true }),
];

const bab13 = [
  h1("13. Asumsi, Batasan, dan Ketentuan"),
  table([2300, 7338], ["Kategori", "Ketentuan"], [
    ["Baseline posisi", "112 posisi: 26 di Direktorat Utama dan 86 di Direktorat Keuangan dan Umum. Incumbent tidak dihitung sebagai posisi terpisah bila pekerjaannya identik."],
    ["Baseline SOP", "Mengacu pada SOP aktif dan unik yang diserahkan PIM. Duplikat, kedaluwarsa, draf, kebijakan tanpa alur, dan instruksi kerja rinci diberi status tersendiri."],
    ["Siklus telaah", "Satu telaah terkonsolidasi dan satu siklus revisi utama per luaran. Koreksi fakta tetap ditutup."],
    ["PM dari PIM", "PIM menunjuk sponsor, manajer proyek, pemilik data, dan pemilik proses yang berwenang mengambil keputusan."],
    ["Kualitas data", "Kami menguji kewajaran data, sedangkan ketepatan data sumber tetap menjadi tanggung jawab penyedia data."],
    ["Jadwal", "Jadwal mengasumsikan data dan pemangku kepentingan tersedia sesuai kalender. Keterlambatan dari sisi PIM menggeser milestone secara sebanding."],
    ["Kunjungan langsung", "Baseline paling banyak tiga kunjungan, biaya sudah termasuk. Kunjungan tambahan dibayar sesuai biaya riil dan disetujui sebelum pemesanan."],
    ["Kepemilikan", "Luaran akhir menjadi milik PIM setelah pembayaran. Metodologi, template umum, dan pengetahuan kerja Sanggabiz tetap milik Sanggabiz."],
    ["Masa berlaku", "Proposal berlaku 30 hari kalender sejak tanggal dokumen."],
  ], { firstBold: true }),
  h2("13.1 Pemicu perubahan lingkup"),
  bullet("Jumlah posisi atau SOP melampaui baseline."),
  bullet("Penambahan lokasi, studi shift, atau observasi intensif di luar baseline."),
  bullet("Permintaan menulis ulang SOP, flowchart, formulir, KPI, job description, atau Delegation of Authority."),
  bullet("Perubahan struktur organisasi setelah tanggal batas yang menuntut pengerjaan ulang material."),
  bullet("Tambahan workshop atau siklus telaah akibat perubahan keputusan yang sudah disahkan."),
];

const bab14 = [
  h1("14. Profil Sanggabiz"),
  p("Sanggabiz adalah penyedia layanan bisnis virtual yang menjalankan fungsi departemen sebagai layanan, serta konsultansi dan pelatihan. Menurut company profile kami, Sanggabiz telah bekerja dengan 6.000+ mitra engagement dari sektor swasta dan pemerintah, menangani 400+ proyek dan layanan, dan didukung 500+ ahli dan pelatih [8]."),
  h2("14.1 Relevansi untuk PIM"),
  bullet("Kesinambungan: kami sudah mengenal baseline, istilah, pola data, dan temuan studi Workforce Analysis PIM sebelumnya. [ISI TIM: konfirmasi bahwa Sanggabiz yang mengerjakan studi sebelumnya dan sebutkan tahun serta ruang lingkupnya]"),
  bullet("Keterpaduan organisasi dan proses: tenaga kerja tidak dipisahkan dari kejelasan peran, kepemilikan proses, dan tata kelola."),
  bullet("Pendampingan tenaga senior: isu material dan konflik lintas direktorat ditangani Direktur Proyek atau Manajer Proyek."),
  bullet("Berbasis data: workbook yang dapat diaudit, keterlacakan bukti, heatmap, penanda pengecualian, dan decision log."),
  bullet("Berorientasi penerapan: setiap luaran dilengkapi pemilik, jadwal, perangkat kerja, serah terima, dan irama tinjauan."),
  bullet("Pengalaman pelatihan: kami menyelenggarakan pelatihan di sembilan kota dan merancang modul serta bahan ajar interaktif untuk berbagai klien [8]."),
  h2("14.2 Kapabilitas inti"),
  table([2900, 6738], ["Pilar", "Kapabilitas"], [
    ["Strategi dan organisasi", "Operating model, desain organisasi, perencanaan tenaga kerja, analisis jabatan dan peran, KPI"],
    ["Proses dan tata kelola", "SOP, RACI, hak keputusan, keselarasan risiko dan kontrol, service excellence"],
    ["Data dan dashboard", "Analitik beban kerja, dashboard eksekutif, pemantauan kinerja, pelaporan manajemen"],
    ["Penerapan dan PMO", "Peta jalan perubahan, PMO, workshop, coaching, transfer kemampuan"],
    ["Pelatihan", "Pengembangan kurikulum dan modul, pelatihan luring dan daring, evaluasi pembelajaran"],
  ], { firstBold: true }),
  p("[ISI TIM: tambahkan daftar pekerjaan konsultansi dan ABK yang relevan beserta klien dan tahunnya]", { before: 100 }),
  spacer(80),
  callout("PANDANGAN KAMI TENTANG ORGANISASI RAMPING", "Organisasi yang ramping tidak berarti organisasi dengan orang sesedikit mungkin. Organisasi yang ramping punya kapasitas yang cukup, satu pemilik keputusan yang jelas, serah terima pekerjaan yang sederhana, kontrol yang sepadan, dan data yang memungkinkan manajemen bertindak sebelum overload menjadi risiko operasi."),
];

const lampiranA = [
  h1("Lampiran A. Struktur Workbook ABK"),
  table([3200, 6438], ["Kolom", "Definisi"], [
    ["Direktorat, unit, posisi", "Struktur dan nama jabatan resmi"],
    ["ID aktivitas dan aktivitas", "Aktivitas unik dan uraian output"],
    ["Pemicu volume", "Transaksi, permintaan, peralatan, rapat, proyek, atau siklus wajib"],
    ["Frekuensi dan volume", "Frekuensi dan volume per periode"],
    ["Waktu per kejadian", "Waktu normal untuk satu kejadian"],
    ["Jam beban kerja tahunan", "Volume × frekuensi × waktu + allowance yang relevan"],
    ["Jam efektif tahunan", "Jam kerja efektif tahunan yang disepakati"],
    ["FTE hasil hitung", "Jam beban kerja tahunan ÷ jam efektif tahunan"],
    ["HC saat ini, kebutuhan dibulatkan, selisih", "Kondisi saat ini, kebutuhan, dan surplus atau kekurangan"],
    ["Confidence dan bukti", "Tinggi, sedang, rendah, dan sumber verifikasi"],
    ["Rekomendasi", "Pertahankan, bagi ulang, rancang ulang, rekrut, latih ulang, otomasi, pantau"],
  ], { firstBold: true }),
  h1("Lampiran B. Scorecard Manfaat", false),
  p("Scorecard ini dipakai PIM setelah proyek selesai. Peninjauan berkala mengikuti siklus PDCA: rencanakan, jalankan, periksa, perbaiki.", { after: 100 }),
  table([1800, 2400, 2800, 1100, 1538], ["Tuas manfaat", "Baseline", "Target atau ukuran", "Pemilik", "Tinjauan"], [
    ["Keseimbangan beban kerja", "Persen underload, normal, stretched, overload", "Posisi di atas 1,30 FTE berkurang. Tindakan atas selisih ditutup", "HR", "Per kuartal"],
    ["Kapasitas manajerial", "Persen aktivitas operasional pada Manager dan VP", "Tindakan delegasi dan perancangan ulang dijalankan", "Direktorat", "Per kuartal"],
    ["Kejelasan peran", "Aktivitas tanpa A, A ganda, tanpa R, konflik pemisahan tugas", "Minimal 95 persen aktivitas bersih. Pengecualian punya pemilik", "Mutu atau GCG", "Bulanan saat penerapan"],
    ["Kecepatan keputusan", "Waktu proses persetujuan dan pengecualian", "Berkurang setelah RACI berlaku", "Pemilik proses", "Bulanan"],
    ["Kesehatan SOP", "Persen SOP aktif dan terkini dibanding perlu pembaruan atau usang", "SOP aktif dan terkini bertambah. Duplikat dihentikan", "Mutu", "Per kuartal"],
    ["Penerapan", "Persen SOP dan RACI yang diperbarui setelah perubahan organisasi", "100 persen permintaan perubahan diproses dalam SLA", "HR dan Mutu", "Per kuartal"],
  ], { firstBold: true, size: 16 }),
];

const rujukan = [
  "Dokumen internal PIM: Report Analisis Beban Kerja, Workforce Analysis, Pupuk Iskandar Muda (Revised with Non-TKNO). Temuan 47 persen stretched, 18 persen overload, dan pola overload Manager dan VP. Rahasia.",
  "PT Pupuk Iskandar Muda, \"Manajemen\" (struktur Direksi terkini). https://www.pim.co.id/perusahaan/manajemen",
  "PT Pupuk Iskandar Muda, Sustainability Corner, Social: Human Capital. Data 628 karyawan, rekrutmen, turnover, dan penilaian kinerja. https://www.pim.co.id/hubungan-investor/sustainability-corner/social/",
  "PT Pupuk Iskandar Muda, Sustainability Report 2024. Nilai ekonomi Rp 5,75 triliun, juga dirangkum TopBusiness, 18 Maret 2026. https://www.pim.co.id/download/SR_PIM_2024.pdf dan https://www.topbusiness.id/125536/perkuat-pemberdayaan-strategi-csr-pim-dorong-nilai-bisnis-berkelanjutan.html",
  "ISO/TC 176, The Process Approach in ISO 9001:2015, panduan tentang kepemilikan proses, akuntabilitas, peran, tanggung jawab, dan wewenang. https://www.iso.org/iso/iso9001_2015_process_approach.pdf",
  "International Organization for Standardization, \"ISO 9001 explained\". https://www.iso.org/home/insights-news/resources/iso-9001-explained.html",
  "Peraturan Menteri PANRB Nomor 1 Tahun 2020 tentang Pedoman Analisis Jabatan dan Analisis Beban Kerja. Dipakai sebagai acuan metode, bukan ketentuan yang otomatis mengikat PIM. https://peraturan.bpk.go.id/Details/143660/permen-pan-rb-no-1-tahun-2020",
  "Sanggabiz, Company Profile: Training Portfolio Sanggabiz 2026 (angka mitra, proyek, tim, dan pengalaman pelatihan).",
  "Model ADDIE. Jurnal Ilmiah Pendidikan Agama Islam, UIN Sunan Gunung Djati. https://journal.uinsgd.ac.id/index.php/jipai/article/download/11042/pdf",
  "Konsep Andragogi dalam Pembelajaran Kursus dan Pelatihan Jarak Jauh: Studi Literatur. Jurnal Pendidikan Luar Sekolah, Universitas Negeri Surabaya. https://ejournal.unesa.ac.id/index.php/jurnal-pendidikan-luar-sekolah/article/view/42812",
  "Taksonomi Bloom Revisi. Universitas Atma Jaya Yogyakarta. https://dpm.uajy.ac.id/wp-content/uploads/2011/08/taksonomi_Bloom_revisi.pdf",
  "Penerapan Model Kirkpatrick untuk Evaluasi Pelatihan. Jurnal Diklat Teknis. https://media.neliti.com/media/publications/275133-penerapan-model-kirkpatrick-untuk-evalua-ced0cdda.pdf",
];
const lampiranC = [
  h1("Lampiran C. Rujukan"),
  ...rujukan.map((t, i) => new Paragraph({ spacing: { after: 70, line: 260 }, indent: { left: 420, hanging: 420 }, children: [new TextRun({ text: `[${i + 1}]\t${t}`, size: 16 })], tabStops: [{ type: TabStopType.LEFT, position: 420 }] })),
  p("Catatan: angka perusahaan dan struktur organisasi perlu dikonfirmasi lagi pada inception meeting karena dapat berubah setelah tanggal proposal. Rujukan eksternal dipakai untuk konteks dan standar metode. Keputusan akhir selalu mengikuti dokumen resmi PIM.", { before: 140, run: { size: 17, color: MUTED } }),
];

// ---------- sampul ----------
const none = (c) => ({ style: BorderStyle.NONE, size: 0, color: c });
const cover = [
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({
      height: { value: 12300, rule: HeightRule.EXACT },
      children: [new TableCell({
        width: { size: W, type: WidthType.DXA }, shading: { fill: NAVY, type: ShadingType.CLEAR, color: "auto" },
        borders: { top: none(NAVY), bottom: none(NAVY), left: none(NAVY), right: none(NAVY) },
        margins: { top: 500, bottom: 400, left: 560, right: 560 },
        children: [
          new Paragraph({ spacing: { after: 2400 }, children: [img("logo_white.png", 150, 40)] }),
          new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "PROPOSAL KONSULTANSI DAN PELATIHAN", font: FONT, bold: true, size: 20, color: GOLD, characterSpacing: 60 })] }),
          new Paragraph({ spacing: { after: 240, line: 400, lineRule: "atLeast" }, children: [new TextRun({ text: "Analisis Beban Kerja Fungsi Umum dan Pelatihan Transfer Kemampuan", font: FONT, bold: true, size: 52, color: "FFFFFF" })] }),
          new Paragraph({ spacing: { after: 2200 }, border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: GOLD, space: 14 } }, children: [new TextRun({ text: "Lanjutan studi Workforce Analysis PT Pupuk Iskandar Muda", font: FONT, size: 24, color: "DCE6F2" })] }),
          new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Disiapkan untuk PT Pupuk Iskandar Muda", font: FONT, size: 22, color: "DCE6F2" })] }),
          new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Diajukan oleh Sanggabiz", font: FONT, bold: true, size: 22, color: "FFFFFF" })] }),
          new Paragraph({ children: [new TextRun({ text: "Oktober 2026", font: FONT, size: 22, color: "DCE6F2" })] }),
        ],
      })],
    })],
  }),
  spacer(300),
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [3212, 3213, 3213],
    rows: [new TableRow({ children: [
      ["112 posisi", "Direktorat Utama dan Direktorat Keuangan dan Umum"],
      ["16-18 minggu", "Analisis, validasi, dan pengesahan oleh manajemen"],
      ["3 hari pelatihan", "Medan, November 2026, maksimal 30 peserta"],
    ].map(([a, b], i) => new TableCell({
      width: { size: 3212 + (i ? 1 : 0), type: WidthType.DXA },
      borders: { top: { style: BorderStyle.SINGLE, size: 18, color: GOLD }, bottom: none("FFFFFF"), left: none("FFFFFF"), right: none("FFFFFF") },
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
      new TextRun({ text: "Sanggabiz  |  Analisis Beban Kerja Fungsi Umum PT Pupuk Iskandar Muda", size: 14, color: MUTED }),
      new TextRun({ text: "\t", size: 14 }),
      new TextRun({ children: ["Halaman ", PageNumber.CURRENT], size: 14, color: MUTED }),
    ],
  })],
});

const page = { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } };
const numRef = (ref) => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 340 } } } }] });
const doc = new Document({
  creator: "Sanggabiz", title: "Analisis Beban Kerja Fungsi Umum PT Pupuk Iskandar Muda",
  styles: {
    default: { document: { run: { font: FONT, size: 19, color: INK } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: NAVY },
        paragraph: { spacing: { before: 0, after: 200 }, outlineLevel: 0, border: { bottom: { style: BorderStyle.SINGLE, size: 14, color: GOLD, space: 6 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 22, bold: true, color: NAVY }, paragraph: { spacing: { before: 240, after: 110 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 260 } } } }] },
    ...["num", "tuj", "uji", "tpl"].map(numRef),
  ] },
  sections: [
    { properties: { page: { ...page, margin: { top: 700, bottom: 700, left: 1134, right: 1134 } } }, children: cover },
    { properties: { page }, footers: { default: footer }, children: [
      ...surat, ...daftarIsi, ...ringkasan, ...bab1, ...bab2, ...bab3, ...bab4, ...bab5, ...bab6, ...bab7, ...bab8, ...bab9,
      ...bab10, ...bab11, ...bab12, ...bab13, ...bab14, ...lampiranA, ...lampiranC,
    ] },
  ],
});

const outDir = path.join(__dirname, "..", "proposal");
fs.mkdirSync(outDir, { recursive: true });
const out = path.join(outDir, "v2 Proposal ABK Fungsi Umum dan Pelatihan - PT Pupuk Iskandar Muda.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ditulis:", out, "total", fmt(TOTAL)); });
