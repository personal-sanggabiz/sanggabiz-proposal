// Membuat dokumen kurikulum pelatihan tiga hari v3 (siap kirim, lampiran proposal). Tanpa sampul, tabel enam kolom, tanpa menit per langkah, bahasa awam.
// Jalankan: NODE_PATH=$(npm root -g) node build_kurikulum.js. Versi 1 dan 2 dibuat oleh build_kurikulum_v1.js dan build_kurikulum_v2.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, Footer,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageNumber, LevelFormat,
  VerticalAlign, TabStopType, PageOrientation, HeightRule,
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
// tag: K = penjelasan, P = praktik dan simulasi, R = review dan rencana tindak
const LABEL = { K: "penjelasan", P: "praktik", R: "review" };
const HARI = [
  { judul: "Hari 1. Memahami Ukuran dan Merancang Pengumpulan Data", modul: [
    { no: 1, judul: "Gambaran ABK dan hasil akhir yang dituju",
      tujuan: "menjelaskan langkah-langkah ABK dan membedakan beban kerja, kinerja, dan jumlah pegawai.",
      outline: ["Hasil akhir ABK: satu lembar hitung beban kerja untuk satu posisi yang sudah jadi", "Enam langkah ABK: menyusun daftar pekerjaan, mengumpulkan data, memeriksa mutu data, menghitung FTE, membaca hasil, dan menentukan tindak lanjut", "Perbedaan beban kerja (banyaknya pekerjaan), kinerja (mutu hasil kerja), dan jumlah pegawai", "Ringkasan temuan studi PIM sebelumnya"],
      akt: [
        [0, 15, "P", "Peserta menerima uraian satu posisi fungsi Umum berisi 8 pekerjaan lengkap dengan jumlahnya per tahun. Tanpa alat hitung dan dalam waktu singkat, tiap peserta menuliskan perkiraan FTE posisi itu pada kartu, lalu menempelkannya di garis angka pada flipchart (0,5 sampai 2,0)."],
        [15, 50, "K", "Fasilitator menunjukkan sebaran kartu yang terkumpul, lalu membuka lembar hitung yang sudah jadi untuk posisi yang sama. Fasilitator menguraikan lembar itu dari hasil akhir mundur ke enam langkah ABK, dan menjelaskan tiga istilah (beban kerja, kinerja, jumlah pegawai) dengan satu contoh untuk tiap istilah."],
        [50, 80, "P", "Lima kelompok (masing-masing 6 orang) menerima paket kasus berisi tiga posisi. Tiap kelompok menuliskan sedikitnya lima masalah beban kerja pada kartu dan menandai langkah ABK yang akan menjawab tiap masalah."],
        [80, 90, "R", "Tiap kelompok menyebut satu masalah. Peserta mengerjakan kuis tiga butir tentang tiga istilah tadi, lalu fasilitator membagikan ringkasan hasil pre-test."],
      ],
      alat: "Slide, uraian satu posisi, flipchart dengan garis angka, kartu, spidol, paket kasus kelompok",
      ukur: "Sebaran perkiraan FTE terpajang dan dibahas bersama. Setiap kelompok menulis sedikitnya lima masalah dan mengaitkan tiap masalah dengan satu langkah ABK. Minimal 80 persen peserta menjawab benar 2 dari 3 butir kuis.",
      out: "Daftar masalah beban kerja awal tiap kelompok" },
    { no: 2, judul: "Menyusun daftar pekerjaan satu posisi",
      tujuan: "menyusun daftar pekerjaan satu posisi lengkap dengan hasil kerja, pemicu, satuan jumlah, dan frekuensi.",
      outline: ["Pekerjaan, tugas, dan SOP: tingkat rincian yang tepat, yaitu satu pekerjaan menghasilkan satu hasil kerja yang bisa dihitung", "Hasil kerja, pemicu, dan satuan jumlah, misalnya surat, permohonan, atau dokumen", "Frekuensi dan masa sibuk dalam setahun", "Pekerjaan yang dilakukan tetapi tidak tertulis di uraian jabatan", "Bukti untuk tiap pekerjaan, misalnya register atau laporan"],
      akt: [
        [0, 25, "K", "Fasilitator menjelaskan tiap kolom formulir dengan dua baris contoh yang tepat dan dua baris yang terlalu luas. Peserta menandai baris yang tepat dan menyebut alasannya."],
        [25, 75, "P", "Tiap peserta mengisi formulir untuk satu posisi dari paket kelompoknya, sedikitnya 8 pekerjaan, dengan merujuk pada SOP yang sudah ada. Fasilitator berkeliling dan memberi masukan."],
        [75, 90, "R", "Peserta bertukar formulir secara berpasangan, memeriksanya dengan daftar periksa lima butir, lalu memperbaikinya."],
      ],
      alat: "Formulir daftar pekerjaan (lembar hitung), daftar periksa lima butir, SOP dalam paket kasus",
      ukur: "Minimal 80 persen peserta mengisi sedikitnya 8 pekerjaan dengan semua kolom wajib terisi. Skor daftar periksa minimal 4 dari 5.",
      out: "Daftar pekerjaan satu posisi untuk tiap peserta" },
    { no: 3, judul: "Menghitung beban kerja dan FTE",
      tujuan: "menghitung FTE satu posisi dari jumlah pekerjaan, frekuensi, waktu pengerjaan, dan jam kerja efektif.",
      outline: ["Beban jam satu pekerjaan: jumlah per tahun dikali waktu pengerjaan satu kali", "Jam kerja efektif setahun: jam kerja yang tersedia setelah dikurangi cuti, hari libur, kegiatan wajib, dan waktu longgar yang wajar", "FTE: total beban jam setahun dibagi jam kerja efektif setahun", "Empat kategori beban: kurang beban, normal, padat, dan berlebih", "Tingkat keyakinan atas data: tinggi, sedang, atau rendah"],
      akt: [
        [0, 25, "K", "Fasilitator menghitung satu pekerjaan, lalu jam kerja efektif setahun, langkah demi langkah di layar."],
        [25, 70, "P", "Tiap peserta menghitung FTE dari daftar pekerjaan yang ia isi di Modul 2. Jumlah dan waktu pengerjaan diambil dari paket kasus yang masih berupa perkiraan. Peserta memakai lembar hitung FTE baku dan menentukan kategori bebannya."],
        [70, 90, "R", "Anggota kelompok membandingkan hasil dan mencari sebab selisih. Tiap kelompok menuliskan tiga data yang masih perlu dibuktikan. Daftar itu menjadi bahan Modul 4."],
      ],
      alat: "Lembar hitung FTE baku, kalkulator, kunci jawaban kasus, tabel kategori beban",
      ukur: "Minimal 80 persen peserta menghasilkan FTE dengan selisih paling banyak 0,10 dari kunci. Setiap kelompok mencatat sedikitnya 3 data yang perlu dibuktikan.",
      out: "FTE awal berdasarkan perkiraan dan daftar data yang perlu dibuktikan" },
    { no: 4, judul: "Merencanakan pengumpulan data",
      tujuan: "memilih sumber data yang sesuai untuk tiap pekerjaan dan menyusun rencana pengumpulan data yang dapat diperiksa.",
      outline: ["Empat sumber data: wawancara, catatan kerja harian, pengamatan langsung, dan data transaksi", "Kekuatan dan batas tiap sumber", "Mencocokkan data dengan sedikitnya dua sumber untuk pekerjaan yang memakan jam terbanyak", "Memilih sumber menurut besarnya jam tiap pekerjaan", "Kerahasiaan data pribadi. Tujuannya menghitung beban kerja, bukan menilai kinerja perorangan"],
      akt: [
        [0, 20, "K", "Fasilitator menjelaskan empat sumber data dengan satu pekerjaan sebagai contoh. Pada pengurusan surat masuk, jumlah surat diambil dari register, lama pengerjaan dari catatan harian, dan hambatan dari wawancara."],
        [20, 70, "P", "Tiap kelompok menandai pekerjaan yang menyumbang jam terbanyak pada FTE awalnya, memilih sumber data untuk tiap pekerjaan, dan mengisi formulir rencana data: siapa yang mengumpulkan, kapan, dari sumber apa, dan buktinya."],
        [70, 85, "R", "Tiap kelompok berkeliling melihat rencana kelompok lain dan menilainya dengan tiga pertanyaan pada kartu penilaian, lalu memperbaiki rencananya sendiri."],
        [85, 90, "R", "Peserta mengisi survei singkat hari pertama."],
      ],
      alat: "Formulir rencana data, FTE awal kelompok, kartu penilaian dengan tiga pertanyaan: apakah pekerjaan berjam terbanyak punya dua sumber data, apakah tiap baris punya penanggung jawab dan tanggal, dan apakah buktinya dapat diperiksa orang lain",
      ukur: "Setiap pekerjaan berjam terbanyak punya sedikitnya dua sumber data. Semua baris rencana memuat penanggung jawab, tanggal, dan bukti.",
      out: "Rencana pengumpulan data tiap kelompok" },
  ] },
  { judul: "Hari 2. Mengumpulkan dan Memeriksa Data", modul: [
    { no: 5, judul: "Teknik wawancara ABK",
      tujuan: "menyusun panduan wawancara yang menggali pekerjaan di luar uraian jabatan, pekerjaan yang diulang, waktu menunggu, dan masa sibuk.",
      outline: ["Susunan wawancara: pembukaan, gambaran satu hari kerja, kejadian penting, dan penutup", "Pertanyaan terbuka dan pertanyaan lanjutan untuk menggali jawaban", "Menggali pekerjaan di luar uraian jabatan, pekerjaan yang diulang karena ada kesalahan, waktu menunggu, pekerjaan mendadak, dan masa sibuk", "Kekeliruan yang sering terjadi dalam wawancara: melebihkan, lupa, dan menjawab sesuai harapan pewawancara", "Meminta contoh nyata dan angka: kejadian terakhir, berapa kali dalam sebulan, dan berapa lama"],
      akt: [
        [0, 25, "K", "Fasilitator menjelaskan susunan wawancara, jenis pertanyaan lanjutan, dan tiga bentuk kekeliruan tadi dengan contoh kalimat."],
        [25, 45, "P", "Dua fasilitator memperagakan dua wawancara pendek tentang pekerjaan yang sama. Wawancara pertama memakai pertanyaan yang menggiring jawaban, wawancara kedua memakai pertanyaan netral. Peserta menilai keduanya dengan lembar pengamatan."],
        [45, 55, "R", "Peserta membandingkan nilai mereka dan menyebut kalimat yang menggiring."],
        [55, 90, "P", "Peserta berpasangan menyusun panduan wawancara untuk satu posisi dari paket kelompok. Fasilitator berkeliling dan memberi masukan."],
      ],
      alat: "Lembar pengamatan wawancara, skenario peragaan, formulir panduan wawancara",
      ukur: "Panduan memuat sedikitnya 10 pertanyaan terbuka, termasuk 3 pertanyaan tentang pekerjaan di luar uraian jabatan. Tidak ada pertanyaan yang menggiring jawaban menurut lembar pengamatan.",
      out: "Panduan wawancara tiap pasangan" },
    { no: 6, judul: "Praktik wawancara dengan bermain peran",
      tujuan: "mewawancarai dengan pertanyaan netral, mencatat bukti dengan tertib, dan menilai wawancara rekan memakai lembar pengamatan.",
      outline: ["Tiga peran: pewawancara, narasumber, dan pengamat", "Mencatat angka dan bukti, serta memisahkan fakta dari dugaan", "Pertanyaan lanjutan yang netral", "Menanyakan kejadian terakhir, bukan rata-rata, untuk memperkirakan waktu pengerjaan", "Memberi umpan balik memakai lembar pengamatan"],
      akt: [
        [0, 15, "K", "Fasilitator menjelaskan skenario, cara memakai kartu peran, dan lembar pengamatan. Narasumber memegang tiga fakta yang disembunyikan dan hanya menjawabnya bila ditanya dengan tepat."],
        [15, 75, "P", "Peserta membentuk 10 kelompok beranggota tiga orang. Ada tiga putaran. Tiap putaran terdiri atas wawancara dan umpan balik singkat, dan peran bergilir sehingga tiap peserta memegang ketiga peran."],
        [75, 90, "R", "Tiap kelompok tiga orang menuliskan dua hal yang akan diubah dalam cara mewawancarai. Fasilitator merangkum pola yang paling sering muncul."],
      ],
      alat: "Kartu peran berisi fakta yang disembunyikan, lembar pengamatan, lembar catatan, penghitung waktu",
      ukur: "Setiap peserta memegang ketiga peran. Minimal 75 persen peserta berhasil menggali sedikitnya 2 dari 3 fakta tersembunyi pada putaran terakhir. Nilai lembar pengamatan minimal 3 dari 5.",
      out: "Catatan wawancara satu putaran dan lembar pengamatan terisi" },
    { no: 7, judul: "Mengisi lembar kerja dan mencocokkan data",
      tujuan: "mengisi lembar kerja ABK dari catatan wawancara dan mencocokkannya dengan catatan harian serta data transaksi.",
      outline: ["Dari catatan wawancara ke lembar kerja: satuan, jumlah per tahun, dan waktu pengerjaan satu kali", "Mengisi dan membaca catatan kerja harian", "Data transaksi (register, tiket, nomor surat) sebagai pembanding", "Mencatat selisih antarsumber di daftar masalah data dan menetapkan sumber yang dipakai", "Menghitung jumlah setahun untuk pekerjaan musiman"],
      akt: [
        [0, 15, "K", "Fasilitator membagikan paket data berisi catatan wawancara yang sudah terisi, catatan harian lima hari, dan cuplikan register. Fasilitator menjelaskan cara menghitung jumlah setahun dari pekerjaan yang hanya ramai pada musim tertentu."],
        [15, 75, "P", "Tiap kelompok mengisi lembar kerja untuk tiga posisi dari ketiga sumber itu. Selisih lebih dari 20 persen antara dua sumber dicatat di daftar masalah data, lengkap dengan sumber yang dipilih dan alasannya. Batas 20 persen berlaku untuk latihan. Batas untuk studi ditetapkan bersama PIM."],
        [75, 90, "R", "Fasilitator memeriksa satu daftar masalah data tiap kelompok bersama kelompok itu dan memberi catatan."],
      ],
      alat: "Paket data (catatan wawancara, catatan harian, cuplikan register), lembar kerja ABK, formulir daftar masalah data",
      ukur: "Setiap jumlah punya satuan dan sumber. Setiap kelompok mencatat sedikitnya 3 selisih beserta sumber yang dipilih dan alasannya.",
      out: "Lembar kerja terisi untuk tiga posisi dan daftar masalah data" },
    { no: 8, judul: "Memeriksa mutu data",
      tujuan: "memeriksa mutu data dengan daftar periksa, menemukan kesalahan yang umum, dan memutuskan perbaikannya.",
      outline: ["Kelengkapan data", "Pekerjaan yang tercatat dua kali", "Satuan atau periode yang tidak cocok, misalnya data bulanan tercampur dengan data tahunan", "Angka yang janggal", "Pekerjaan yang dihitung ganda pada dua posisi", "Uji kewajaran: jam yang dilaporkan satu orang tidak boleh melebihi jam yang tersedia"],
      akt: [
        [0, 20, "K", "Fasilitator menjelaskan daftar periksa enam butir dengan satu contoh untuk tiap butir."],
        [20, 45, "P", "Tiap kelompok memeriksa lembar latihan buatan fasilitator yang berisi enam kesalahan yang sengaja disisipkan, lalu mencatat temuannya."],
        [45, 70, "P", "Kelompok bertukar lembar kerja dari Modul 7, memeriksanya dengan daftar periksa, dan menuliskan temuan di daftar masalah data."],
        [70, 85, "R", "Lembar kembali ke pemiliknya. Pemilik memutuskan menerima atau menolak tiap temuan dengan alasan tertulis, lalu memperbaiki lembarnya."],
        [85, 90, "R", "Peserta mengisi survei singkat hari kedua."],
      ],
      alat: "Lembar latihan berisi enam kesalahan sisipan, daftar periksa mutu data, formulir daftar masalah data",
      ukur: "Setiap kelompok menemukan sedikitnya 5 dari 6 kesalahan sisipan. Setiap temuan di daftar masalah data punya keputusan dan alasan.",
      out: "Daftar masalah data yang sudah diputuskan dan lembar kerja yang diperbaiki" },
  ] },
  { judul: "Hari 3. Menghitung, Membaca Hasil, dan Menindaklanjuti", modul: [
    { no: 9, judul: "Menghitung ulang FTE dan menguji asumsi",
      tujuan: "menghitung FTE dari data yang sudah diperiksa dan menganalisis bagaimana asumsi mengubah kategori beban.",
      outline: ["Menghitung ulang FTE dengan lembar kerja yang sudah diperbaiki", "Selisih FTE awal dan FTE akhir serta sebabnya", "Menguji akibat perubahan asumsi: tiga pilihan jam kerja efektif, waktu longgar, dan waktu pengerjaan pada pekerjaan terbesar", "Rentang FTE dan kategori beban", "Tingkat keyakinan atas data: tinggi bila dua sumber cocok, sedang bila satu sumber ditambah perkiraan, rendah bila hanya perkiraan"],
      akt: [
        [0, 25, "K", "Fasilitator memperlihatkan satu contoh latihan: beban kerja 1.600 jam per tahun dengan tiga pilihan jam kerja efektif, yaitu 1.250, 1.400, dan 1.500 jam. FTE menjadi 1,28, 1,14, dan 1,07, sehingga kategorinya bergeser dari padat ke normal. Angka ini contoh latihan dan bukan angka PIM."],
        [25, 70, "P", "Tiap kelompok menghitung ulang FTE tiga posisi, membuat tabel tiga pilihan jam kerja efektif, dan memberi tingkat keyakinan pada tiap posisi."],
        [70, 90, "R", "Tiap kelompok melaporkan satu posisi yang kategorinya berubah antarpilihan. Peserta membahas keputusan yang aman diambil sekarang dan keputusan yang menunggu data tambahan."],
      ],
      alat: "Lembar hitung FTE dengan tabel tiga pilihan asumsi, kunci jawaban kasus latihan, tabel kategori beban",
      ukur: "Setiap posisi punya FTE pada tiga pilihan asumsi, kategori beban, dan tingkat keyakinan beserta alasannya. Pada kasus latihan, selisih hitung paling banyak 0,05 dari kunci.",
      out: "FTE akhir tiga posisi dengan rentang hasil dan tingkat keyakinan" },
    { no: 10, judul: "Membaca heatmap dan menelusuri akar masalah",
      tujuan: "membaca heatmap, menelusuri penyebab beban tinggi, dan mempertahankan hasilnya saat dicocokkan dengan manajer.",
      outline: ["Membaca heatmap per posisi dan per unit", "Gejala dan akar masalah: permintaan yang memang tinggi, pekerjaan manual, pekerjaan yang diulang, banyaknya lapisan persetujuan, penugasan sementara, dan peran yang tumpang tindih", "Pohon akar masalah", "Pencocokan dengan manajer: pertanyaan yang akan muncul dan bukti yang disiapkan", "Catatan keputusan: keputusan, alasan, penanggung jawab, dan tanggal"],
      akt: [
        [0, 25, "K", "Fasilitator menunjukkan contoh heatmap dan satu pohon akar masalah, lalu menjelaskan cara menyiapkan diri untuk pencocokan dengan manajer."],
        [25, 55, "P", "Tiap kelompok menyusun heatmap dari hasil Modul 9 dan pohon akar masalah untuk satu posisi yang padat atau berlebih."],
        [55, 80, "R", "Simulasi pencocokan dengan manajer. Peserta dari kelompok lain berperan sebagai panel manajer (Manager, VP, dan SDM). Tiap kelompok memaparkan hasilnya secara singkat, lalu menjawab pertanyaan panel dengan bukti dan mencatat keputusan."],
        [80, 90, "P", "Kelompok melengkapi catatan keputusan berdasarkan hasil panel."],
      ],
      alat: "Heatmap (lembar hitung dengan warna otomatis), kartu pertanyaan panel, formulir catatan keputusan, flipchart",
      ukur: "Setiap pohon akar masalah memuat sedikitnya dua penyebab beserta buktinya. Setiap catatan keputusan memuat alasan dan penanggung jawab.",
      out: "Heatmap kelompok, pohon akar masalah, dan catatan keputusan" },
    { no: 11, judul: "Memilih tindak lanjut dan menguji dampaknya pada RACI",
      tujuan: "memilih tindak lanjut yang sesuai dengan akar masalah dan menguji dampaknya pada pembagian peran pelaksana (Responsible) dan penanggung jawab akhir (Accountable).",
      outline: ["Pilihan tindak lanjut: menyederhanakan pekerjaan, membagi ulang, merancang ulang, memakai otomasi, melatih ulang, menambah kapasitas, atau memantau", "Mencocokkan akar masalah dengan tindakan", "Menguji RACI: siapa menjadi pelaksana dan penanggung jawab akhir yang baru, apakah beban penerimanya melewati 1,15 FTE, dan apakah ada dua penanggung jawab akhir pada satu pekerjaan atau pekerjaan tanpa penanggung jawab akhir", "Urutan prioritas menurut dampak pada FTE, tingkat kepentingan, serta dampak pada kepatuhan atau keselamatan", "Batas pekerjaan: SOP yang ada menjadi rujukan dan tidak direvisi dalam pelatihan"],
      akt: [
        [0, 20, "K", "Fasilitator menjelaskan tabel yang mencocokkan akar masalah dengan tindakan, lalu menguji satu pekerjaan pada RACI sebagai contoh."],
        [20, 70, "P", "Tiap kelompok memilih tiga temuan dari Modul 10, menetapkan tindakan, membuat potongan RACI (4 sampai 6 pekerjaan) sebelum dan sesudah tindakan, menghitung ulang FTE penerima peran pelaksana yang baru, dan mengisi matriks prioritas."],
        [70, 90, "R", "Kelompok bertukar matriks. Kelompok penilai mencari satu masalah, yaitu dua penanggung jawab akhir pada satu pekerjaan atau penerima peran pelaksana yang melewati 1,15 FTE, lalu memberi masukan."],
      ],
      alat: "Tabel pencocokan akar masalah dengan tindakan, formulir RACI, matriks prioritas, lembar hitung FTE",
      ukur: "Setiap tindakan punya penanggung jawab, perkiraan perubahan FTE, dan hasil uji RACI. Tidak ada pekerjaan dengan dua penanggung jawab akhir atau tanpa penanggung jawab akhir, atau ada penjelasan tertulis bila tetap demikian.",
      out: "Matriks tindak lanjut berprioritas dan potongan RACI sebelum dan sesudah tindakan" },
    { no: 12, judul: "Rencana 30-60-90 hari dan penutupan",
      tujuan: "menyusun rencana 30-60-90 hari untuk unitnya dan menunjukkan penguasaan lewat post-test.",
      outline: ["Rencana 30-60-90 hari: pekerjaan yang akan dilakukan unit sendiri dengan perangkat yang dibawa pulang", "Penanggung jawab, target, dan ukuran keberhasilan yang dapat diukur", "Tinjauan oleh manajemen, serta kaitannya dengan klinik daring 30 hari setelah pelatihan dan tinjauan 60 sampai 90 hari setelah pelatihan", "Komitmen penerapan"],
      akt: [
        [0, 15, "K", "Fasilitator menjelaskan cara menulis target yang terukur dengan dua contoh yang tepat dan dua yang kabur, lalu menunjukkan hubungan rencana dengan klinik daring 30 hari setelah pelatihan."],
        [15, 50, "P", "Tiap peserta menyusun rencana 30-60-90 hari untuk unitnya dengan formulir, sedikitnya tiga langkah untuk tiap tahap, masing-masing dengan penanggung jawab dan ukuran keberhasilan."],
        [50, 60, "R", "Peserta berbagi bertiga dan memberi dua masukan pada rencana rekannya."],
        [60, 80, "R", "Post-test 15 butir (12 pilihan ganda dan 3 hitungan singkat) dengan kisi-kisi yang sama dengan pre-test."],
        [80, 90, "R", "Peserta mengisi evaluasi akhir, menuliskan satu komitmen penerapan, dan menerima sertifikat elektronik."],
      ],
      alat: "Formulir rencana 30-60-90 hari, soal post-test, formulir evaluasi akhir",
      ukur: "Setiap rencana memuat sedikitnya tiga langkah untuk tiap tahap dengan penanggung jawab dan ukuran keberhasilan. Skor post-test minimal 75 dan peningkatan rata-rata minimal 20 persen dari nilai awal.",
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
  return table(widths, ["Topik atau modul", "Outline materi", "Aktivitas pembelajaran", "Alat atau media pelatihan", "Indikator capaian peserta", "Output konkret"], rows, { size: 16, gap: 80, lineH: 270, firstFill: true });
}

// ---------- isi ----------
const logo = new ImageRun({ type: "png", data: fs.readFileSync(A("logo_navy.png")), transformation: { width: 150, height: 30 } });
const judul = [
  new Paragraph({ spacing: { after: 160 }, children: [logo] }),
  new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: "LAMPIRAN PROPOSAL", bold: true, size: 20, color: "B58900", characterSpacing: 60 })] }),
  new Paragraph({ spacing: { after: 120, line: 380, lineRule: "atLeast" }, children: [new TextRun({ text: "Kurikulum Pelatihan Transfer Kemampuan Analisis Beban Kerja", bold: true, size: 44, color: NAVY })] }),
  new Paragraph({ spacing: { after: 120 }, border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: GOLD, space: 10 } }, children: [new TextRun({ text: "Lampiran Proposal Analisis Beban Kerja Fungsi Umum dan Rencana RACI Enterprise", size: 21, color: MUTED })] }),
  p("Disiapkan untuk PT Pupuk Iskandar Muda. Diajukan oleh Sanggabiz. Oktober 2026.", { run: { size: 18, color: MUTED }, before: 80, after: 120 }),
];

const ringkasan = [
  h1("1. Ringkasan kurikulum", false),
  p("Pelatihan berlangsung tiga hari di Medan dengan 12 modul. Tiap modul berdurasi 2 jam pelajaran (JPL), yaitu 90 menit, dengan 1 JPL setara 45 menit. Totalnya 24 JPL atau 18 jam efektif."),
  p("Peserta bekerja dalam lima kelompok dengan satu kasus yang sama dari modul pertama sampai terakhir. Dari kasus itu peserta menghasilkan daftar pekerjaan, rencana pengumpulan data, panduan wawancara, lembar kerja, hitungan FTE, heatmap, daftar tindakan, dan rencana 30-60-90 hari. Dengan perangkat yang dibawa pulang, peserta diharapkan dapat menjalankan ABK sederhana di unitnya."),
  table([1700, 2200, 2500, 3238], ["Durasi", "Peserta", "Fasilitator", "Porsi belajar"], [
    ["3 hari, 24 JPL", "Maksimal 30 orang, lima kelompok", "Satu fasilitator utama dan satu fasilitator pendamping", `Sekitar ${pct(TOT.K)} persen penjelasan, ${pct(TOT.P)} persen praktik dan simulasi, ${pct(TOT.R)} persen review dan rencana tindak`],
  ], { align: [AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] }),
  h2("Istilah yang dipakai"),
  table([2300, 7338], ["Istilah", "Artinya"], [
    ["ABK", "Analisis Beban Kerja: cara menghitung banyaknya pekerjaan satu posisi dan berapa orang yang dibutuhkan untuk mengerjakannya."],
    ["FTE", "Full-Time Equivalent: beban kerja dalam satuan pegawai penuh waktu. FTE 1,0 berarti pekerjaan itu setara satu orang yang bekerja penuh. FTE 1,3 berarti setara 1,3 orang."],
    ["Jam kerja efektif", "Jam kerja setahun yang benar-benar tersedia untuk mengerjakan tugas, setelah dikurangi cuti, hari libur, kegiatan wajib, dan waktu longgar yang wajar."],
    ["Kategori beban", "Kurang beban (underload): di bawah 0,85 FTE. Normal: 0,85 sampai 1,15. Padat (stretched): di atas 1,15 sampai 1,30. Berlebih (overload): di atas 1,30. Batas awal mengikuti studi PIM sebelumnya."],
    ["Heatmap", "Tabel berwarna yang menunjukkan posisi mana yang kurang beban, normal, padat, atau berlebih."],
    ["RACI", "Pembagian peran untuk tiap pekerjaan: pelaksana (Responsible), penanggung jawab akhir (Accountable), pihak yang dimintai pendapat (Consulted), dan pihak yang diberi tahu (Informed)."],
    ["Catatan kerja harian dan data transaksi", "Catatan kerja harian diisi pegawai setiap hari tentang pekerjaan dan lamanya (workload diary). Data transaksi adalah catatan yang sudah ada di kantor, seperti register surat, tiket layanan, atau nomor dokumen."],
  ], { firstBold: true, size: 16 }),
  new Paragraph({ children: [new TextRun({ text: "Daftar modul", bold: true, size: 22, color: NAVY })], pageBreakBefore: true, spacing: { after: 100 }, keepNext: true }),
  table([750, 3750, 5138], ["Modul", "Topik", "Output utama"],
    HARI.flatMap((h) => [{ group: h.judul }, ...h.modul.map((m) => [String(m.no), m.judul, m.out])]),
    { align: [AlignmentType.CENTER], size: 16 }),
];

const dasar = [
  h1("2. Dasar rancangan"),
  p("Kurikulum ini disusun dengan enam pertimbangan berikut."),
  num("**Urutan mengikuti urutan kerja ABK.** Peserta lebih dulu merencanakan dan mengumpulkan data, lalu mengolah, memeriksa, dan menindaklanjuti. Urutan ini sejalan dengan tahapan penyusunan ABK pada pedoman nasional: pengumpulan data, pengolahan data, serta verifikasi dan validasi [1][2]. Pedoman itu mengatur instansi pemerintah. Kami memakainya sebagai rujukan urutan kerja, sedangkan kategori FTE tetap mengikuti studi PIM sebelumnya."),
  num("**Dimulai dari hasil akhir.** Pada Modul 1 peserta melihat lembar hitung yang sudah jadi, lalu menguraikannya mundur menjadi enam langkah. Hasil akhir ditetapkan lebih dulu, baru materi disusun untuk mencapainya. Cara ini dikenal sebagai backward design [3]."),
  num("**Mengalami dulu, menjelaskan sesudahnya.** Peserta menebak FTE tanpa alat hitung sebelum menerima rumus. Penjelasan datang setelah mereka melihat sendiri betapa lebar sebaran tebakan. Pola ini mengikuti siklus belajar Kolb: mengalami, merenungkan, memahami konsep, lalu menerapkan [4]."),
  num("**Satu kasus dari awal sampai akhir.** Setiap kelompok memegang tiga posisi fungsi Umum yang dipilih bersama PIM satu sampai dua minggu sebelum pelatihan, dengan data yang namanya sudah dihapus. Hasil satu modul menjadi bahan modul berikutnya. Bila PIM belum dapat membagikan data, kami menyiapkan kasus simulasi yang ditandai jelas sebagai simulasi."),
  num("**FTE dihitung dua kali.** Hitungan pertama memakai perkiraan (Modul 3), hitungan kedua memakai data yang sudah diperiksa (Modul 9). Selisihnya memperlihatkan pengaruh mutu data pada kategori beban."),
  num("**Data tidak bergantung pada satu sumber.** Perkiraan jam kerja dari wawancara atau survei dicocokkan dengan catatan harian dan data transaksi. Kajian tentang pemakaian waktu memakai catatan harian sebagai pembanding bagi perkiraan umum jam kerja [5], dan pemeriksaan keabsahan catatan harian umumnya bertumpu pada kecocokan beberapa penanda [6]."),
  p("Tujuan tiap modul memakai kata kerja taksonomi Bloom revisi [7], misalnya menjelaskan, menyusun, menghitung, memilih, memeriksa, membaca, dan menguji. Pelatihan tidak merevisi SOP. SOP yang ada dipakai sebagai rujukan untuk memetakan pekerjaan dan RACI.", { before: 100 }),
  h2("Alur kasus: hasil satu modul menjadi bahan modul berikutnya"),
  table([3200, 1100, 5338], ["Langkah ABK", "Modul", "Yang dikerjakan kelompok pada kasusnya"], [
    ["Merumuskan masalah", "1", "Daftar masalah beban kerja untuk tiga posisi"],
    ["1. Daftar pekerjaan dan hitungan awal", "2 sampai 3", "Daftar 8 pekerjaan per posisi dan FTE berdasarkan perkiraan"],
    ["2. Rencana dan pengumpulan data", "4 sampai 7", "Rencana data, panduan wawancara, wawancara, dan lembar kerja dari tiga sumber"],
    ["3. Pemeriksaan mutu data", "8", "Pemeriksaan silang antarkelompok dan keputusan perbaikan"],
    ["4. Menghitung FTE", "9", "FTE akhir dengan tiga pilihan asumsi dan tingkat keyakinan"],
    ["5. Membaca hasil", "10", "Heatmap, pohon akar masalah, dan pencocokan dengan panel manajer"],
    ["6. Tindak lanjut dan uji RACI", "11", "Tindakan, RACI sebelum dan sesudah, dan urutan prioritas"],
    ["Rencana penerapan", "12", "Rencana 30-60-90 hari untuk tiap unit"],
  ], { align: [AlignmentType.LEFT, AlignmentType.CENTER], firstBold: true, size: 16 }),
];

const jadwal = [
  h1("3. Jadwal harian"),
  p("Tiap hari terdiri atas empat blok 90 menit. Istirahat 15 menit pada pagi dan sore hari serta istirahat makan siang pukul 12.15 sampai 13.15 WIB. Jam dapat disesuaikan dengan jam kerja PIM. Pre-test dikerjakan sebelum Hari 1, sehingga tidak memakai waktu kelas.", { after: 100 }),
  table([1500, 2000, 6138], ["Hari", "Waktu", "Kegiatan"], [
    ["Hari 1", "08.30 - 09.00", "Registrasi dan pembukaan (di luar 24 JPL)"],
    ["", "09.00 - 10.30", "Modul 1. Gambaran ABK dan hasil akhir yang dituju"],
    ["", "10.45 - 12.15", "Modul 2. Menyusun daftar pekerjaan satu posisi"],
    ["", "13.15 - 14.45", "Modul 3. Menghitung beban kerja dan FTE"],
    ["", "15.00 - 16.30", "Modul 4. Merencanakan pengumpulan data"],
    ["Hari 2", "08.30 - 09.00", "Ulasan Hari 1: tiap kelompok menyebut satu hal yang akan dipakai dan satu pertanyaan (di luar 24 JPL)"],
    ["", "09.00 - 10.30", "Modul 5. Teknik wawancara ABK"],
    ["", "10.45 - 12.15", "Modul 6. Praktik wawancara dengan bermain peran"],
    ["", "13.15 - 14.45", "Modul 7. Mengisi lembar kerja dan mencocokkan data"],
    ["", "15.00 - 16.30", "Modul 8. Memeriksa mutu data"],
    ["Hari 3", "08.30 - 09.00", "Ulasan Hari 2: tiap kelompok menyebut satu temuan masalah data dari lembarnya (di luar 24 JPL)"],
    ["", "09.00 - 10.30", "Modul 9. Menghitung ulang FTE dan menguji asumsi"],
    ["", "10.45 - 12.15", "Modul 10. Membaca heatmap dan menelusuri akar masalah"],
    ["", "13.15 - 14.45", "Modul 11. Memilih tindak lanjut dan menguji dampaknya pada RACI"],
    ["", "15.00 - 16.30", "Modul 12. Rencana 30-60-90 hari dan penutupan, termasuk post-test"],
  ], { align: [AlignmentType.CENTER, AlignmentType.CENTER], firstBold: true }),
];

const hasil = [
  h1("5. Hasil yang dibawa pulang"),
  p("Semua hasil dikumpulkan di buku kerja digital peserta. Formulir dan lembar hitung diserahkan kepada PIM sebagai bagian dari perangkat pelatihan.", { after: 100 }),
  table([4500, 1000, 1400, 2738], ["Hasil", "Modul", "Pemegang", "Dipakai untuk"], [
    ["Daftar masalah beban kerja awal", "1", "Kelompok", "Bahan rencana 30-60-90 hari"],
    ["Daftar pekerjaan satu posisi", "2", "Peserta", "Dasar ABK untuk posisi lain di unit"],
    ["FTE awal dan daftar data yang perlu dibuktikan", "3", "Peserta", "Pembanding hasil akhir"],
    ["Rencana pengumpulan data", "4", "Kelompok", "Pengumpulan data nyata di unit"],
    ["Panduan wawancara terisi", "5", "Pasangan", "Wawancara ABK berikutnya"],
    ["Catatan wawancara dan lembar pengamatan terisi", "6", "Peserta", "Bahan berlatih dan umpan balik"],
    ["Lembar kerja ABK tiga posisi dan daftar masalah data", "7", "Kelompok", "Contoh lembar kerja siap pakai"],
    ["Daftar masalah data yang sudah diputuskan dan daftar periksa mutu data", "8", "Kelompok", "Pemeriksaan data sebelum dihitung"],
    ["FTE akhir dengan tiga pilihan asumsi dan tingkat keyakinan", "9", "Kelompok", "Dasar keputusan kebutuhan pegawai"],
    ["Heatmap, pohon akar masalah, dan catatan keputusan", "10", "Kelompok", "Bahan pencocokan dengan manajemen"],
    ["Matriks tindak lanjut berprioritas dan potongan RACI", "11", "Kelompok", "Usulan tindakan dan pembagian peran"],
    ["Rencana 30-60-90 hari, hasil post-test, dan komitmen", "12", "Peserta", "Pegangan unit dan bahan klinik 30 hari setelah pelatihan"],
    ["Lembar hitung FTE baku, seluruh formulir, lembar pengamatan, dan kunci kasus", "Awal", "PIM", "Dipakai sendiri setelah pelatihan"],
  ], { align: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.LEFT], size: 16 }),
  h2("Rujukan"),
  ...[
    "Peraturan Menteri PANRB Nomor 1 Tahun 2020 tentang Pedoman Analisis Jabatan dan Analisis Beban Kerja. JDIH BPK RI. https://peraturan.bpk.go.id/Home/Details/143660/permen-pan-rb-no-1-tahun-2020",
    "Kementerian PANRB, Langkah Penyusunan Anjab dan ABK dalam Penentuan Jabatan. https://www.menpan.go.id/site/berita-terkini/langkah-penyusunan-anjab-dan-abk-dalam-penentuan-jabatan",
    "Wiggins, G. dan McTighe, J. (2005). Understanding by Design (edisi kedua). ASCD.",
    "Kolb, D. A. (1984). Experiential Learning: Experience as the Source of Learning and Development. Prentice Hall.",
    "Improving Stylised Working Time Estimates with Time Diary Data: A Multi Study Assessment for the UK. https://www.ncbi.nlm.nih.gov/pmc/articles/PMC6609582/",
    "Harms dkk. (2019). A validation study of the Eurostat harmonised European time use study (HETUS) diary using wearable technology. https://pmc.ncbi.nlm.nih.gov/articles/PMC6546612",
    "Anderson, L. W. dan Krathwohl, D. R. (Ed.) (2001). A Taxonomy for Learning, Teaching, and Assessing: A Revision of Bloom's Taxonomy of Educational Objectives. Longman.",
  ].map((t, i) => p(`[${i + 1}] ${t}`, { after: 40, run: { size: 15 }, line: 250 })),
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
  p("Tiap modul berdurasi 2 JPL (90 menit) dan memuat tujuan, outline materi, aktivitas, alat atau media, indikator capaian, dan output. Langkah aktivitas ditulis berurutan, dan fasilitator menyesuaikan waktu tiap langkah dengan kondisi kelas. Label penjelasan, praktik, dan review menunjukkan jenis kegiatan.", { after: 100 }),
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
const out = path.join(outDir, "v3 Kurikulum Pelatihan Tiga Hari ABK - PT Pupuk Iskandar Muda.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ditulis:", out); });
