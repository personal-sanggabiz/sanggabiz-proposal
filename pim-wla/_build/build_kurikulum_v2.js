// Membuat dokumen kurikulum pelatihan tiga hari v2 (pengganti bagian 7.4 proposal v3). Tabel enam kolom, tanpa menit per langkah.
// Jalankan: NODE_PATH=$(npm root -g) node build_kurikulum.js. Versi 1 dibuat oleh build_kurikulum_v1.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, Footer,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageNumber, LevelFormat,
  VerticalAlign, TabStopType, PageOrientation,
} = require("docx");

const NAVY = "0B396D", GOLD = "FFC00C", INK = "1D1D1F", MUTED = "5C5F66", LIGHT = "EEF3F9", LINE = "C9D2DE";
const FONT = "Montserrat";
const WP = 9638;   // lebar isi A4 tegak, margin 2 cm
const WL = 14570;  // lebar isi A4 mendatar, margin 2 cm
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
  spacing: { after: o.after ?? 110, before: o.before ?? 0, line: o.line ?? 290 }, keepNext: o.keepNext,
});
const h1 = (text, pb = true) => new Paragraph({ heading: "Heading1", children: [new TextRun({ text })], pageBreakBefore: pb, keepNext: true });
const h2 = (text) => new Paragraph({ heading: "Heading2", children: [new TextRun({ text })], keepNext: true });
const num = (text) => new Paragraph({ numbering: { reference: "num", level: 0 }, children: runs(text), spacing: { after: 80, line: 290 } });
const spacer = (n = 120) => new Paragraph({ spacing: { after: n }, children: [] });

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
      alignment: o.align || AlignmentType.LEFT, spacing: { after: o.gap ?? 40, line: o.lineH ?? 260 },
      children: runs(t, { size: o.size || 17, bold: o.bold, color: o.color || INK }),
    })),
  });
}
function table(widths, header, rows, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const head = new TableRow({
    tableHeader: true, cantSplit: true,
    children: header.map((t, i) => cell(t, widths[i], { fill: NAVY, color: "FFFFFF", bold: true, size: o.headSize || 16, align: o.headAlign && o.headAlign[i] })),
  });
  const body = rows.map((r, ri) => {
    if (r && r.group) return new TableRow({ cantSplit: true, children: [cell(r.group, total, { fill: "DCE6F2", bold: true, span: widths.length })] });
    return new TableRow({
      cantSplit: !o.split,
      children: r.map((c, i) => cell(c, widths[i], {
        fill: (o.firstBold || o.firstFill) && i === 0 ? LIGHT : (ri % 2 ? "F7F9FC" : undefined),
        align: o.align && o.align[i], size: o.size, bold: o.firstBold && i === 0, gap: o.gap, lineH: o.lineH,
      })),
    });
  });
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: [head, ...body] });
}

// ---------- data kurikulum ----------
// tag: K = konsep, P = praktik dan simulasi, R = review dan rencana tindak
const LABEL = { K: "konsep", P: "praktik", R: "review" };
const HARI = [
  { judul: "Hari 1. Memahami Ukuran dan Merancang Pengumpulan Data", modul: [
    { no: 1, judul: "Gambaran ABK dan hasil akhir yang dituju",
      tujuan: "menjelaskan enam langkah ABK dan membedakan beban kerja, kinerja, dan jumlah pegawai.",
      outline: ["Hasil akhir ABK: lembar FTE satu posisi yang sudah jadi", "Enam langkah ABK: inventaris aktivitas, pengumpulan data, uji mutu, hitung FTE, membaca hasil, tindak lanjut", "Beban kerja, kinerja, dan jumlah pegawai adalah tiga hal yang berbeda", "Ringkasan temuan studi PIM sebelumnya dari PIM"],
      akt: [
        [0, 15, "P", "Peserta menerima deskripsi satu posisi fungsi Umum berisi 8 aktivitas lengkap dengan volume tahunan. Tanpa alat hitung dan dalam waktu singkat, tiap peserta menulis perkiraan FTE posisi itu pada kartu, lalu menempelkannya di garis angka flipchart (0,5 sampai 2,0)."],
        [15, 50, "K", "Fasilitator menunjukkan sebaran kartu, lalu membuka lembar FTE yang sudah jadi untuk posisi yang sama. Fasilitator membongkar lembar itu dari belakang menjadi enam langkah ABK dan menjelaskan tiga istilah dengan satu contoh tiap istilah."],
        [50, 80, "P", "Lima kelompok (masing-masing 6 orang) menerima paket kasus tiga posisi. Tiap kelompok menulis sedikitnya lima isu beban kerja pada kartu dan menandai langkah ABK yang akan menjawab tiap isu."],
        [80, 90, "R", "Tiap kelompok menyebut satu isu. Peserta mengerjakan kuis tiga butir tentang tiga istilah, lalu fasilitator membagikan ringkasan hasil pre-test."],
      ],
      alat: "Slide, deskripsi satu posisi, flipchart dengan garis angka, kartu, spidol, paket kasus kelompok",
      ukur: "Sebaran perkiraan FTE terpajang dan dibahas. Setiap kelompok menulis sedikitnya lima isu dan mengaitkan tiap isu ke satu langkah ABK. Minimal 80 persen peserta menjawab benar 2 dari 3 butir kuis.",
      out: "Peta isu awal per kelompok" },
    { no: 2, judul: "Menyusun inventaris aktivitas",
      tujuan: "menyusun inventaris aktivitas satu posisi lengkap dengan output, pemicu, satuan volume, dan frekuensi.",
      outline: ["Aktivitas, tugas, dan SOP: tingkat rincian yang tepat, yaitu satu aktivitas menghasilkan satu output yang bisa dihitung", "Output, pemicu, dan satuan volume", "Frekuensi dan puncak musiman", "Pekerjaan di luar uraian jabatan", "Sumber bukti tiap aktivitas"],
      akt: [
        [0, 25, "K", "Fasilitator menjelaskan tiap kolom template dengan dua baris contoh yang tepat dan dua baris yang terlalu luas. Peserta menandai baris yang tepat dan menyebut alasannya."],
        [25, 75, "P", "Tiap peserta mengisi template untuk satu posisi dari paket kelompoknya, sedikitnya 8 aktivitas, dengan merujuk pada SOP yang sudah ada. Fasilitator berkeliling dan memberi umpan balik."],
        [75, 90, "R", "Peserta bertukar template secara berpasangan, memeriksanya dengan daftar cek lima butir, lalu memperbaikinya."],
      ],
      alat: "Template inventaris aktivitas (lembar hitung), daftar cek lima butir, SOP dalam paket kasus",
      ukur: "Minimal 80 persen peserta mengisi sedikitnya 8 aktivitas dengan semua kolom wajib terisi. Skor daftar cek minimal 4 dari 5.",
      out: "Inventaris aktivitas satu posisi per peserta" },
    { no: 3, judul: "Menghitung beban kerja dan FTE",
      tujuan: "menghitung FTE satu posisi dari volume, frekuensi, waktu per kejadian, dan jam kerja efektif.",
      outline: ["Beban jam satu aktivitas: volume tahunan dikali waktu per kejadian", "Jam kerja efektif tahunan: hari kerja, cuti, hari libur, pelatihan wajib, dan kelonggaran", "FTE: total beban jam tahunan dibagi jam kerja efektif tahunan", "Kategori underload, normal, stretched, dan overload", "Tingkat keyakinan (confidence level) data"],
      akt: [
        [0, 25, "K", "Fasilitator menghitung satu aktivitas, lalu jam kerja efektif tahunan, langkah demi langkah di layar."],
        [25, 70, "P", "Tiap peserta menghitung FTE dari inventaris yang ia isi di Modul 2. Volume dan waktu diambil dari paket kasus yang masih berupa perkiraan. Peserta memakai lembar hitung FTE baku dan menentukan kategorinya."],
        [70, 90, "R", "Anggota kelompok membandingkan hasil dan menelusuri sebab selisih. Tiap kelompok menulis tiga data yang masih perlu dibuktikan. Daftar itu menjadi bahan Modul 4."],
      ],
      alat: "Lembar hitung FTE baku, kalkulator, kunci jawaban kasus, tabel kategori",
      ukur: "Minimal 80 persen peserta menghasilkan FTE dengan selisih paling banyak 0,10 dari kunci. Setiap kelompok mencatat sedikitnya 3 data yang perlu dibuktikan.",
      out: "FTE awal berbasis perkiraan dan daftar data yang perlu dibuktikan" },
    { no: 4, judul: "Merancang rencana pengumpulan data",
      tujuan: "memilih sumber data yang sesuai untuk tiap aktivitas dan menyusun rencana pengumpulan data yang dapat diperiksa.",
      outline: ["Empat sumber data: wawancara, catatan kerja harian (workload diary), pengamatan langsung, dan data transaksi", "Kekuatan dan batas tiap sumber", "Triangulasi: aktivitas dengan jam terbesar memakai sedikitnya dua sumber", "Memilih sumber menurut bobot jam aktivitas", "Kerahasiaan data pribadi dan tujuan pengumpulan, yaitu bukan penilaian kinerja individu"],
      akt: [
        [0, 20, "K", "Fasilitator menjelaskan empat sumber data dengan satu aktivitas sebagai contoh. Pada pengurusan surat masuk, jumlah surat diambil dari register, lama pengerjaan dari catatan harian, dan hambatan dari wawancara."],
        [20, 70, "P", "Tiap kelompok menandai aktivitas yang menyumbang jam terbesar pada FTE awalnya, memilih sumber data untuk tiap aktivitas, dan mengisi template rencana data: siapa, kapan, sumber, dan bukti."],
        [70, 85, "R", "Keliling antarkelompok (gallery walk). Tiap kelompok menilai rencana kelompok lain dengan tiga pertanyaan pada kartu penilaian, lalu memperbaiki rencananya."],
        [85, 90, "R", "Peserta mengisi pulse survey hari pertama."],
      ],
      alat: "Template rencana data, FTE awal kelompok, kartu penilaian galeri dengan tiga pertanyaan: apakah aktivitas berjam terbesar punya dua sumber, apakah tiap baris punya pemilik dan tanggal, dan apakah buktinya dapat diperiksa orang lain",
      ukur: "Setiap aktivitas berjam terbesar punya sedikitnya dua sumber data. Semua baris rencana memuat pemilik, tanggal, dan bukti.",
      out: "Rencana pengumpulan data per kelompok" },
  ] },
  { judul: "Hari 2. Mengumpulkan dan Menguji Data", modul: [
    { no: 5, judul: "Teknik wawancara ABK",
      tujuan: "menyusun panduan wawancara terstruktur yang menggali pekerjaan di luar uraian jabatan, pengerjaan ulang, waktu tunggu, dan puncak beban.",
      outline: ["Struktur wawancara: pembukaan, alur kerja satu hari, kejadian kritis, penutup", "Pertanyaan terbuka dan pertanyaan lanjutan (probing)", "Menggali pekerjaan di luar uraian jabatan, pengerjaan ulang, waktu tunggu, pekerjaan mendadak, dan puncak musiman", "Bias wawancara: melebihkan, melupakan, dan menjawab sesuai harapan", "Meminta contoh nyata dan angka: kejadian terakhir, berapa kali dalam sebulan, berapa lama"],
      akt: [
        [0, 25, "K", "Fasilitator menjelaskan struktur wawancara, jenis pertanyaan lanjutan, dan tiga bentuk bias dengan contoh kalimat."],
        [25, 45, "P", "Dua fasilitator memperagakan dua wawancara pendek untuk aktivitas yang sama. Wawancara pertama memakai pertanyaan menggiring dan yang kedua netral. Peserta menilai keduanya dengan rubrik pengamatan."],
        [45, 55, "R", "Peserta membandingkan skor rubrik dan menyebut kalimat yang menggiring."],
        [55, 90, "P", "Peserta berpasangan menyusun panduan wawancara untuk satu posisi dari paket kelompok. Fasilitator berkeliling dan memberi umpan balik."],
      ],
      alat: "Rubrik pengamatan wawancara, skenario peragaan, template panduan wawancara",
      ukur: "Panduan memuat sedikitnya 10 pertanyaan terbuka, termasuk 3 pertanyaan untuk pekerjaan di luar uraian jabatan. Tidak ada pertanyaan menggiring menurut rubrik.",
      out: "Panduan wawancara terisi per pasangan" },
    { no: 6, judul: "Praktik wawancara dengan bermain peran",
      tujuan: "mewawancarai dengan pertanyaan netral, mencatat bukti secara tertib, dan menilai wawancara rekan dengan rubrik.",
      outline: ["Peran pewawancara, narasumber, dan pengamat", "Mencatat angka dan bukti serta memisahkan fakta dari dugaan", "Pertanyaan lanjutan yang netral", "Menanyakan kejadian terakhir, bukan rata-rata, untuk memperkirakan waktu", "Memberi umpan balik dengan rubrik"],
      akt: [
        [0, 15, "K", "Fasilitator menjelaskan skenario, cara memakai kartu peran, dan rubrik. Narasumber memegang tiga fakta tersembunyi yang hanya dijawab bila ditanya dengan tepat."],
        [15, 75, "P", "Peserta membentuk 10 kelompok tiga orang. Ada tiga putaran. Tiap putaran terdiri atas wawancara dan umpan balik singkat, dan peran bergilir sehingga tiap peserta memegang ketiga peran."],
        [75, 90, "R", "Tiap kelompok tiga orang menulis dua hal yang akan diubah dalam wawancaranya. Fasilitator merangkum pola yang paling sering muncul."],
      ],
      alat: "Kartu peran berisi fakta tersembunyi, rubrik pengamatan, lembar catatan, penghitung waktu",
      ukur: "Setiap peserta memegang ketiga peran. Minimal 75 persen peserta menggali sedikitnya 2 dari 3 fakta tersembunyi pada putaran terakhir. Skor rubrik minimal 3 dari 5.",
      out: "Catatan wawancara satu putaran dan rubrik terisi" },
    { no: 7, judul: "Mengisi lembar kerja dan mencocokkan data",
      tujuan: "mengisi lembar kerja ABK dari catatan wawancara dan mencocokkannya dengan catatan harian serta data transaksi.",
      outline: ["Dari catatan wawancara ke lembar kerja: satuan, volume tahunan, waktu per kejadian", "Mengisi dan membaca catatan kerja harian", "Data transaksi (register, tiket, nomor surat) sebagai pembanding", "Mencatat selisih di issue log dan menetapkan sumber yang dipakai", "Annualisasi aktivitas musiman"],
      akt: [
        [0, 15, "K", "Fasilitator membagikan paket data: catatan wawancara terisi, catatan harian lima hari, dan cuplikan register. Fasilitator menjelaskan cara menghitung volume tahunan dari data musiman."],
        [15, 75, "P", "Tiap kelompok mengisi lembar kerja untuk tiga posisi dari ketiga sumber itu. Setiap selisih lebih dari 20 persen antara dua sumber dicatat di issue log beserta sumber yang dipilih dan alasannya. Ambang 20 persen berlaku untuk latihan. Ambang untuk studi ditetapkan bersama PIM."],
        [75, 90, "R", "Fasilitator memeriksa satu issue log tiap kelompok bersama kelompok itu dan memberi catatan."],
      ],
      alat: "Paket data (catatan wawancara, catatan harian, cuplikan register), lembar kerja ABK, issue log",
      ukur: "Setiap volume punya satuan dan sumber. Setiap kelompok mencatat sedikitnya 3 selisih beserta sumber terpilih dan alasannya.",
      out: "Lembar kerja terisi untuk tiga posisi dan issue log" },
    { no: 8, judul: "Uji mutu data",
      tujuan: "memeriksa mutu data dengan daftar cek, menemukan kesalahan umum, dan memutuskan perbaikannya.",
      outline: ["Kelengkapan data", "Aktivitas ganda", "Satuan dan periode tidak cocok, misalnya bulan dengan tahun", "Nilai menyimpang", "Hitung ganda antarposisi", "Uji kewajaran: jam yang dilaporkan per orang tidak melebihi jam yang tersedia"],
      akt: [
        [0, 20, "K", "Fasilitator menjelaskan daftar cek mutu enam butir dengan satu contoh tiap butir."],
        [20, 45, "P", "Tiap kelompok memeriksa lembar uji buatan fasilitator yang berisi enam kesalahan sisipan, lalu mencatat temuannya."],
        [45, 70, "P", "Kelompok bertukar lembar kerja Modul 7, memeriksanya dengan daftar cek, dan menulis temuan di issue log."],
        [70, 85, "R", "Lembar kembali ke pemiliknya. Pemilik memutuskan menerima atau menolak tiap temuan dengan alasan tertulis, lalu memperbaiki lembarnya."],
        [85, 90, "R", "Peserta mengisi pulse survey hari kedua."],
      ],
      alat: "Lembar uji berisi enam kesalahan sisipan, daftar cek mutu, issue log",
      ukur: "Setiap kelompok menemukan sedikitnya 5 dari 6 kesalahan sisipan. Setiap temuan di issue log punya keputusan dan alasan.",
      out: "Issue log yang sudah diputuskan dan lembar kerja yang diperbaiki" },
  ] },
  { judul: "Hari 3. Menghitung, Menafsirkan, dan Menindaklanjuti", modul: [
    { no: 9, judul: "Menghitung ulang FTE dan menguji asumsi",
      tujuan: "menghitung FTE dari data yang sudah diuji dan menganalisis pengaruh asumsi terhadap kategori beban.",
      outline: ["Hitung ulang FTE dengan lembar kerja yang sudah diperbaiki", "Selisih FTE awal dan FTE akhir serta sebabnya", "Uji sensitivitas: tiga skenario jam kerja efektif, kelonggaran, dan waktu per kejadian pada aktivitas terbesar", "Rentang FTE dan kategori", "Tingkat keyakinan data: tinggi bila dua sumber cocok, sedang bila satu sumber ditambah perkiraan, rendah bila hanya perkiraan"],
      akt: [
        [0, 25, "K", "Fasilitator memperlihatkan satu contoh latihan: beban 1.600 jam per tahun dengan tiga skenario jam kerja efektif, yaitu 1.250, 1.400, dan 1.500 jam. FTE menjadi 1,28, 1,14, dan 1,07, sehingga kategori bergeser dari stretched ke normal. Angka ini contoh latihan dan bukan angka PIM."],
        [25, 70, "P", "Tiap kelompok menghitung ulang FTE tiga posisi, membuat tabel tiga skenario, dan memberi tingkat keyakinan pada tiap posisi."],
        [70, 90, "R", "Tiap kelompok melaporkan satu posisi yang kategorinya berubah antarskenario. Peserta membahas keputusan yang aman dibuat dan keputusan yang menunggu data tambahan."],
      ],
      alat: "Lembar hitung FTE dengan tabel skenario, kunci jawaban kasus uji, tabel kategori",
      ukur: "Setiap posisi punya FTE pada tiga skenario, kategori, dan tingkat keyakinan beserta alasannya. Pada kasus uji, selisih hitung paling banyak 0,05 dari kunci.",
      out: "FTE akhir tiga posisi dengan rentang skenario dan tingkat keyakinan" },
    { no: 10, judul: "Membaca heatmap dan menelusuri akar masalah",
      tujuan: "menafsirkan heatmap, menelusuri akar penyebab beban tinggi, dan mempertahankan hasilnya dalam kalibrasi dengan manajer.",
      outline: ["Membaca heatmap per posisi dan unit", "Gejala dan akar masalah: permintaan nyata, pekerjaan manual, pengerjaan ulang, lapisan persetujuan, penugasan sementara, tumpang tindih peran", "Pohon akar masalah", "Kalibrasi dengan manajer: pertanyaan yang akan muncul dan bukti yang disiapkan", "Catatan keputusan: keputusan, alasan, pemilik, tanggal"],
      akt: [
        [0, 25, "K", "Fasilitator menunjukkan contoh heatmap dan satu pohon akar masalah, lalu menjelaskan cara menyiapkan diri untuk kalibrasi."],
        [25, 55, "P", "Tiap kelompok menyusun heatmap dari hasil Modul 9 dan pohon akar masalah untuk satu posisi stretched atau overload."],
        [55, 80, "R", "Simulasi kalibrasi. Peserta kelompok lain berperan sebagai panel manajer (Manager, VP, dan HR). Tiap kelompok memaparkan hasilnya secara singkat, lalu menjawab pertanyaan panel dengan bukti dan mencatat keputusan."],
        [80, 90, "P", "Kelompok melengkapi catatan keputusan berdasarkan hasil panel."],
      ],
      alat: "Heatmap (lembar hitung dengan pewarnaan bersyarat), kartu pertanyaan panel, template catatan keputusan, flipchart",
      ukur: "Setiap pohon akar memuat sedikitnya dua akar penyebab beserta buktinya. Setiap catatan keputusan memuat alasan dan pemilik.",
      out: "Heatmap kelompok, pohon akar masalah, dan catatan keputusan" },
    { no: 11, judul: "Memilih tindak lanjut dan menguji dampaknya pada RACI",
      tujuan: "memilih tindak lanjut yang sesuai akar masalah dan menguji dampaknya pada penugasan Responsible dan Accountable.",
      outline: ["Pilihan tindak lanjut: penyederhanaan, pembagian ulang, perancangan ulang, otomasi, pelatihan ulang, penambahan kapasitas, dan pemantauan", "Mencocokkan akar masalah dengan tindakan", "Uji RACI: siapa menjadi R dan A baru, apakah beban penerima melewati 1,15 FTE, apakah ada dua A pada satu aktivitas atau aktivitas tanpa A", "Prioritas menurut dampak FTE, kritikalitas, serta dampak kepatuhan atau keselamatan", "Batas pekerjaan: SOP yang ada menjadi rujukan dan tidak direvisi dalam pelatihan"],
      akt: [
        [0, 20, "K", "Fasilitator menjelaskan tabel pencocokan akar masalah dengan tindakan dan menguji satu aktivitas pada RACI sebagai contoh."],
        [20, 70, "P", "Tiap kelompok memilih tiga temuan dari Modul 10, menetapkan tindakan, membuat potongan RACI (4 sampai 6 aktivitas) sebelum dan sesudah tindakan, menghitung ulang FTE penerima peran R yang baru, dan mengisi matriks prioritas."],
        [70, 90, "R", "Kelompok bertukar matriks. Kelompok penilai mencari satu masalah, yaitu A ganda atau penerima R yang melewati 1,15 FTE, lalu memberi umpan balik."],
      ],
      alat: "Tabel pencocokan akar masalah dengan tindakan, template RACI, matriks prioritas, lembar hitung FTE",
      ukur: "Setiap tindakan punya pemilik, perkiraan perubahan FTE, dan hasil uji RACI. Tidak ada aktivitas dengan A ganda atau tanpa A, atau ada penjelasan tertulis bila tetap demikian.",
      out: "Matriks tindak lanjut berprioritas dan potongan RACI sebelum dan sesudah" },
    { no: 12, judul: "Rencana 30-60-90 hari dan penutupan",
      tujuan: "menyusun rencana 30-60-90 hari untuk unitnya dan menunjukkan penguasaan lewat post-test.",
      outline: ["Rencana 30-60-90: pekerjaan unit sendiri dengan perangkat yang dibawa pulang", "Pemilik, target, dan indikator yang dapat diukur", "Tinjauan manajemen dan kaitannya dengan online clinic D+30 serta tinjauan D+60 sampai D+90", "Komitmen penerapan"],
      akt: [
        [0, 15, "K", "Fasilitator menjelaskan cara menulis target yang terukur dengan dua contoh yang tepat dan dua yang kabur, lalu menunjukkan hubungan rencana dengan clinic D+30."],
        [15, 50, "P", "Tiap peserta menyusun rencana 30-60-90 hari untuk unitnya dengan template, sedikitnya tiga langkah per tahap, masing-masing dengan pemilik dan indikator."],
        [50, 60, "R", "Peserta berbagi bertiga dan memberi dua masukan pada rencana rekan."],
        [60, 80, "R", "Post-test 15 butir (12 pilihan ganda dan 3 hitungan singkat) dengan kisi-kisi yang sama dengan pre-test."],
        [80, 90, "R", "Peserta mengisi evaluasi akhir, menulis satu komitmen penerapan, dan menerima sertifikat elektronik."],
      ],
      alat: "Template rencana 30-60-90, soal post-test, formulir evaluasi akhir",
      ukur: "Setiap rencana memuat sedikitnya tiga langkah per tahap dengan pemilik dan indikator. Skor post-test minimal 75 dan peningkatan rata-rata minimal 20 persen dari baseline.",
      out: "Rencana 30-60-90 hari, hasil post-test, dan komitmen peserta" },
  ] },
];

// ---------- pemeriksaan rancangan: bobot waktu tiap langkah hanya untuk menghitung porsi dan tidak ditampilkan ----------
const TOT = { K: 0, P: 0, R: 0 };
HARI.forEach((h) => h.modul.forEach((m) => {
  let t = 0;
  m.akt.forEach(([a, b, tag]) => {
    if (a !== t) throw new Error(`Modul ${m.no}: celah waktu di menit ${t}`);
    TOT[tag] += b - a; t = b;
  });
  if (t !== 90) throw new Error(`Modul ${m.no}: total ${t} menit`);
}));
const ALL = TOT.K + TOT.P + TOT.R;
const pct = (n) => Math.round((n * 100) / ALL);
console.log("Porsi", TOT, ALL, pct(TOT.K), pct(TOT.P), pct(TOT.R));
if (ALL !== 1080) throw new Error("Total harus 1080 menit (24 JPL)");

// ---------- tabel kurikulum format standar tipe 1 (lanskap) ----------
const cap = (t) => t[0].toUpperCase() + t.slice(1);
function kurikulum(hari) {
  const widths = [2700, 3000, 3900, 1700, 1900, 1370];
  const rows = hari.modul.map((m) => [
    [`**Modul ${m.no}. ${m.judul}**`, `Tujuan: peserta mampu ${m.tujuan}`],
    m.outline.map((t, i) => `${i + 1}. ${t}`),
    m.akt.map(([, , tag, t], i) => `${i + 1}. **${cap(LABEL[tag])}.** ${t}`),
    m.alat, m.ukur, m.out,
  ]);
  return table(widths, ["Topik atau modul", "Outline materi", "Aktivitas pembelajaran (andragogi)", "Alat atau media pelatihan", "Indikator capaian peserta", "Output konkret"], rows, { size: 16, gap: 80, lineH: 270, firstFill: true });
}

// ---------- isi ----------
const logo = new ImageRun({ type: "png", data: fs.readFileSync(A("logo_navy.png")), transformation: { width: 150, height: 30 } });
const judul = [
  new Paragraph({ spacing: { after: 360 }, children: [logo] }),
  new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: "KURIKULUM PELATIHAN TIGA HARI", bold: true, size: 20, color: "B58900", characterSpacing: 60 })] }),
  new Paragraph({ spacing: { after: 120, line: 380, lineRule: "atLeast" }, children: [new TextRun({ text: "Transfer Kemampuan Analisis Beban Kerja", bold: true, size: 44, color: NAVY })] }),
  new Paragraph({ spacing: { after: 160 }, border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: GOLD, space: 10 } }, children: [new TextRun({ text: "PT Pupuk Iskandar Muda, Medan", size: 22, color: MUTED })] }),
  p("Dokumen ini adalah usulan pengganti bagian 7.4 (Kurikulum Pelatihan Tiga Hari) pada Proposal Analisis Beban Kerja Fungsi Umum dan Rencana RACI Enterprise. Bagian 7.5 (evaluasi Kirkpatrick) dan bagian lain tidak berubah. Diajukan oleh Sanggabiz. Draf 14 Oktober 2026.", { run: { size: 18, color: MUTED }, before: 120, after: 200 }),
];

const ringkasan = [
  h2("1. Ringkasan kurikulum"),
  p("Pelatihan berlangsung tiga hari dengan 12 modul, masing-masing 2 JPL (90 menit). Total 24 JPL, yaitu 18 jam efektif dengan 1 JPL setara 45 menit. Peserta bekerja dalam lima kelompok dengan satu kasus yang sama dari Modul 1 sampai Modul 12. Dari kasus itu mereka membuat inventaris aktivitas, rencana data, panduan wawancara, lembar kerja, hitungan FTE, heatmap, matriks tindak lanjut, dan rencana 30-60-90 hari."),
  table([2300, 2400, 2400, 2538], ["Durasi", "Peserta", "Fasilitator", "Porsi belajar"], [
    ["3 hari, 24 JPL", "Maksimal 30 orang, lima kelompok", "Satu Lead Facilitator dan satu Co-Facilitator", `Sekitar ${pct(TOT.K)} persen konsep, ${pct(TOT.P)} persen praktik dan simulasi, ${pct(TOT.R)} persen review dan rencana tindak`],
  ], { align: [AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] }),
  spacer(160),
  new Paragraph({ children: [new TextRun({ text: "Daftar modul", bold: true, size: 22, color: NAVY })], pageBreakBefore: true, spacing: { after: 100 }, keepNext: true }),
  table([750, 3750, 5138], ["Modul", "Topik", "Output utama"],
    HARI.flatMap((h) => [{ group: h.judul }, ...h.modul.map((m) => [String(m.no), m.judul, m.out])]),
    { align: [AlignmentType.CENTER], size: 16 }),
  h2("Alur kasus: output satu modul menjadi bahan modul berikutnya"),
  table([3200, 1100, 5338], ["Langkah ABK", "Modul", "Yang dikerjakan kelompok pada kasusnya"], [
    ["Merumuskan isu", "1", "Peta isu awal untuk tiga posisi"],
    ["1. Inventaris dan hitungan awal", "2 sampai 3", "Inventaris 8 aktivitas per posisi dan FTE berbasis perkiraan"],
    ["2. Rencana dan pengumpulan data", "4 sampai 7", "Rencana data, panduan wawancara, wawancara, dan lembar kerja dari tiga sumber"],
    ["3. Uji mutu data", "8", "Pemeriksaan silang antarkelompok dan keputusan perbaikan"],
    ["4. Hitung FTE", "9", "FTE akhir pada tiga skenario dan tingkat keyakinan"],
    ["5. Membaca hasil", "10", "Heatmap, pohon akar masalah, dan kalibrasi di depan panel"],
    ["6. Tindak lanjut dan uji RACI", "11", "Tindakan, RACI sebelum dan sesudah, dan prioritas"],
    ["Rencana penerapan", "12", "Rencana 30-60-90 hari per unit"],
  ], { align: [AlignmentType.LEFT, AlignmentType.CENTER], firstBold: true, size: 16 }),
];

const dasar = [
  h1("2. Dasar rancangan"),
  p("Enam keputusan rancangan berikut membedakan kurikulum ini dari susunan modul yang biasa."),
  num("**Urutan mengikuti urutan kerja ABK.** Peserta lebih dulu merencanakan dan mengumpulkan data, lalu mengolah, memverifikasi, dan menindaklanjuti. Urutan ini sejalan dengan tahapan penyusunan ABK pada pedoman nasional: pengumpulan data, pengolahan data, serta verifikasi dan validasi [1][2]. Pedoman itu mengatur instansi pemerintah. Kami memakainya sebagai rujukan urutan kerja, sedangkan ambang kategori FTE tetap mengikuti studi PIM sebelumnya."),
  num("**Dimulai dari hasil akhir.** Pada Modul 1 peserta melihat lembar FTE yang sudah jadi, lalu membongkarnya mundur menjadi enam langkah. Cara ini mengikuti backward design: hasil akhir ditetapkan lebih dulu, materi menyusul."),
  num("**Pengalaman lebih dulu, penjelasan sesudahnya.** Peserta menebak FTE tanpa alat hitung sebelum menerima rumus. Penjelasan baru datang setelah mereka melihat sebaran tebakan yang lebar. Pola ini mengikuti siklus belajar Kolb, yaitu pengalaman, refleksi, konsep, lalu penerapan, dengan instruksi seminimal mungkin sebelum pengalaman pertama dan waktu latihan yang cukup [3]."),
  num("**Satu kasus dari awal sampai akhir.** Setiap kelompok memegang tiga posisi fungsi Umum yang dipilih bersama PIM pada H-14 sampai H-7, dengan data yang dianonimkan. Output satu modul menjadi bahan modul berikutnya. Bila PIM belum dapat membagikan data, kami menyiapkan kasus simulasi yang ditandai jelas sebagai simulasi."),
  num("**FTE dihitung dua kali.** Hitungan pertama memakai perkiraan (Modul 3), hitungan kedua memakai data yang sudah diuji (Modul 9). Selisihnya memperlihatkan pengaruh mutu data pada kategori beban."),
  num("**Data tidak bergantung pada satu sumber.** Perkiraan jam kerja dari wawancara atau survei dicocokkan dengan catatan harian dan data transaksi. Kajian pemakaian waktu memakai catatan harian sebagai pembanding bagi perkiraan umum jam kerja [4], dan validasi catatan harian pada umumnya bertumpu pada kecocokan beberapa penanda [5]."),
  p("Tujuan tiap modul memakai kata kerja taksonomi Bloom revisi: menjelaskan, menyusun, menghitung, memilih, memeriksa, menafsirkan, dan menguji. Pelatihan tidak merevisi SOP. SOP yang ada dipakai sebagai rujukan pemetaan aktivitas dan RACI.", { before: 100 }),
];

const jadwal = [
  h1("3. Jadwal harian"),
  p("Tiap hari terdiri atas empat blok 90 menit. Istirahat 15 menit pada pagi dan sore hari serta istirahat makan siang pukul 12.15 sampai 13.15 WIB. Jam dapat disesuaikan dengan jam kerja PIM. Pre-test dikerjakan sebelum Hari 1, sehingga tidak memakai waktu kelas.", { after: 100 }),
  table([1500, 2000, 6138], ["Hari", "Waktu", "Kegiatan"], [
    ["Hari 1", "08.30 - 09.00", "Registrasi dan pembukaan (di luar 24 JPL)"],
    ["", "09.00 - 10.30", "Modul 1. Gambaran ABK dan hasil akhir yang dituju"],
    ["", "10.45 - 12.15", "Modul 2. Menyusun inventaris aktivitas"],
    ["", "13.15 - 14.45", "Modul 3. Menghitung beban kerja dan FTE"],
    ["", "15.00 - 16.30", "Modul 4. Merancang rencana pengumpulan data"],
    ["Hari 2", "08.30 - 09.00", "Ulasan Hari 1: tiap kelompok menyebut satu hal yang dipakai dan satu pertanyaan (di luar 24 JPL)"],
    ["", "09.00 - 10.30", "Modul 5. Teknik wawancara ABK"],
    ["", "10.45 - 12.15", "Modul 6. Praktik wawancara dengan bermain peran"],
    ["", "13.15 - 14.45", "Modul 7. Mengisi lembar kerja dan mencocokkan data"],
    ["", "15.00 - 16.30", "Modul 8. Uji mutu data"],
    ["Hari 3", "08.30 - 09.00", "Ulasan Hari 2: tiap kelompok menyebut satu temuan mutu data dari lembarnya (di luar 24 JPL)"],
    ["", "09.00 - 10.30", "Modul 9. Menghitung ulang FTE dan menguji asumsi"],
    ["", "10.45 - 12.15", "Modul 10. Membaca heatmap dan menelusuri akar masalah"],
    ["", "13.15 - 14.45", "Modul 11. Memilih tindak lanjut dan menguji dampaknya pada RACI"],
    ["", "15.00 - 16.30", "Modul 12. Rencana 30-60-90 hari dan penutupan, termasuk post-test"],
  ], { align: [AlignmentType.CENTER, AlignmentType.CENTER], firstBold: true }),
];

const hasil = [
  h1("5. Hasil yang dibawa pulang"),
  p("Semua hasil dikumpulkan di workbook digital peserta. Template dan lembar hitung diserahkan kepada PIM sebagai bagian dari perangkat pelatihan.", { after: 100 }),
  table([4300, 1200, 1700, 2438], ["Hasil", "Modul", "Pemegang", "Dipakai untuk"], [
    ["Peta isu awal", "1", "Kelompok", "Bahan rencana 30-60-90 hari"],
    ["Inventaris aktivitas satu posisi", "2", "Peserta", "Dasar ABK posisi lain di unit"],
    ["FTE awal dan daftar data yang perlu dibuktikan", "3", "Peserta", "Pembanding hasil akhir"],
    ["Rencana pengumpulan data", "4", "Kelompok", "Pengumpulan data nyata di unit"],
    ["Panduan wawancara terisi", "5", "Pasangan", "Wawancara ABK berikutnya"],
    ["Catatan wawancara dan rubrik terisi", "6", "Peserta", "Latihan mandiri dan umpan balik"],
    ["Lembar kerja ABK tiga posisi dan issue log", "7", "Kelompok", "Contoh lembar kerja siap pakai"],
    ["Issue log yang diputuskan dan daftar cek mutu", "8", "Kelompok", "Pemeriksaan data sebelum dihitung"],
    ["FTE akhir pada tiga skenario beserta tingkat keyakinan", "9", "Kelompok", "Dasar keputusan kebutuhan pegawai"],
    ["Heatmap, pohon akar masalah, dan catatan keputusan", "10", "Kelompok", "Bahan kalibrasi dengan manajemen"],
    ["Matriks tindak lanjut berprioritas dan potongan RACI", "11", "Kelompok", "Usulan tindakan dan penugasan peran"],
    ["Rencana 30-60-90 hari, hasil post-test, dan komitmen", "12", "Peserta", "Pegangan unit dan bahan clinic D+30"],
    ["Lembar hitung FTE baku, seluruh template, rubrik, dan kunci kasus", "Sebelum pelatihan", "PIM", "Penerapan mandiri setelah pelatihan"],
  ], { align: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.LEFT] }),
  h2("Rujukan"),
  ...[
    "Peraturan Menteri PANRB Nomor 1 Tahun 2020 tentang Pedoman Analisis Jabatan dan Analisis Beban Kerja. JDIH BPK RI. https://peraturan.bpk.go.id/Home/Details/143660/permen-pan-rb-no-1-tahun-2020",
    "Kementerian PANRB, Langkah Penyusunan Anjab dan ABK dalam Penentuan Jabatan. https://www.menpan.go.id/site/berita-terkini/langkah-penyusunan-anjab-dan-abk-dalam-penentuan-jabatan",
    "UF/IFAS Extension (EDIS), publikasi tentang penyelenggaraan workshop untuk peserta dewasa dan siklus belajar Kolb. https://journals.flvc.org/edis/article/download/116843/115007/169081",
    "Improving Stylised Working Time Estimates with Time Diary Data: A Multi Study Assessment for the UK. https://www.ncbi.nlm.nih.gov/pmc/articles/PMC6609582/",
    "Harms dkk. (2019), A validation study of the Eurostat harmonised European time use study (HETUS) diary using wearable technology. https://pmc.ncbi.nlm.nih.gov/articles/PMC6546612",
  ].map((t, i) => p(`[${i + 1}] ${t}`, { after: 60, run: { size: 16 }, line: 260 })),
];

// ---------- perakitan ----------
const footerTegak = new Footer({ children: [new Paragraph({
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: LINE, space: 6 } },
  tabStops: [{ type: TabStopType.RIGHT, position: WP }],
  children: [
    new TextRun({ text: "Sanggabiz  |  Kurikulum Pelatihan ABK PT Pupuk Iskandar Muda", size: 14, color: MUTED }),
    new TextRun({ text: "\t", size: 14 }),
    new TextRun({ children: ["Halaman ", PageNumber.CURRENT], size: 14, color: MUTED }),
  ],
})] });
const footerMendatar = new Footer({ children: [new Paragraph({
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: LINE, space: 6 } },
  tabStops: [{ type: TabStopType.RIGHT, position: WL }],
  children: [
    new TextRun({ text: "Sanggabiz  |  Kurikulum Pelatihan ABK PT Pupuk Iskandar Muda", size: 14, color: MUTED }),
    new TextRun({ text: "\t", size: 14 }),
    new TextRun({ children: ["Halaman ", PageNumber.CURRENT], size: 14, color: MUTED }),
  ],
})] });

const margin = { top: 1134, bottom: 1134, left: 1134, right: 1134 };
const tegak = { size: { width: 11906, height: 16838 }, margin };
const mendatar = { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin };
const numRef = (ref) => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 340 } } } }] });

const kurikulumDoc = [
  h1("4. Kurikulum per hari", false),
  p("Tabel mengikuti format standar kurikulum Sanggabiz dengan enam kolom. Kata kerja pada tujuan mengikuti taksonomi Bloom revisi. Tiap modul berdurasi 2 JPL (90 menit). Langkah aktivitas ditulis berurutan tanpa pembagian menit, sehingga fasilitator dapat menyesuaikan waktu tiap langkah dengan kondisi kelas. Label konsep, praktik, dan review menunjukkan jenis kegiatan.", { after: 100 }),
  ...HARI.flatMap((h, i) => [
    new Paragraph({ children: [new TextRun({ text: h.judul, bold: true, size: 22, color: NAVY })], spacing: { before: i ? 0 : 120, after: 100 }, pageBreakBefore: i > 0, keepNext: true }),
    kurikulum(h),
  ]),
];

const doc = new Document({
  creator: "Sanggabiz", title: "Kurikulum Pelatihan ABK PT Pupuk Iskandar Muda",
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
  numbering: { config: [numRef("num")] },
  sections: [
    { properties: { page: tegak }, footers: { default: footerTegak }, children: [...judul, ...ringkasan, ...dasar, ...jadwal] },
    { properties: { page: mendatar }, footers: { default: footerMendatar }, children: kurikulumDoc },
    { properties: { page: tegak }, footers: { default: footerTegak }, children: hasil },
  ],
});

const outDir = path.join(__dirname, "..", "kurikulum");
fs.mkdirSync(outDir, { recursive: true });
const out = path.join(outDir, "v2 Kurikulum Pelatihan Tiga Hari ABK - PT Pupuk Iskandar Muda.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ditulis:", out); });
