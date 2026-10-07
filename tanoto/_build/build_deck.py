import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import *
import budget as B

OUT = os.path.normpath(os.path.join(HERE, "..", "Proposal Teknis - Sanggabiz - Tanoto PPG LPTK.pptx"))
P = lambda n: os.path.join(ASSETS, n)

prs = new_prs()
N = [0]


def new(bg="grad", foot=True):
    s = blank(prs, bg)
    N[0] += 1
    if foot:
        footer(s, N[0])
    return s


# =============== 1. SAMPUL ===============
def s_cover():
    s = new(foot=False)
    logo(s, 0.7, 0.55, 0.42)
    text(s, 0.7, 1.55, 6.9, 0.35, "PROPOSAL TEKNIS", size=14, bold=True, color=YELLOW, caps=True, tracking=300)
    text(s, 0.7, 2.0, 6.9, 2.6, "Mitra Pelaksana Program Pendampingan Teknis LPTK Penyelenggara PPG",
         size=34, color=WHITE, font=F_XB, spacing=1.0)
    text(s, 0.7, 4.55, 6.6, 0.7, [[("Menuju Akreditasi Unggul: dari rencana tindak lanjut ke bukti yang siap diakreditasi", {})]],
         size=16, color=WHITE, font=F_M)
    pill(s, 0.7, 5.5, 6.4, 0.6, fill=NAVY_DD, line=TEAL, lw=1.5)
    text(s, 0.7, 5.5, 6.4, 0.6, "12 LPTK  ·  5 Provinsi  ·  6 Oktober – 31 Desember 2026", size=14, bold=True,
         color=WHITE, align="c", anchor="m")
    text(s, 0.7, 6.3, 6.6, 0.7, [[("Untuk: ", {"color": GREY}), ("Tanoto Foundation", {"bold": True})],
                                 [("RFP No. 0648/TF-JKT/VIII/2026  ·  Diajukan 30 September 2026", {"color": GREY, "size": 11})]],
         size=13, after=2)
    # kolase foto miring (gaya sampul compro)
    photo(s, P("image3.jpeg"), 8.3, 0.35, 2.3, 3.0, rot=6, focus=(0.5, 0.3))
    photo(s, P("image37.jpeg"), 10.55, 0.75, 2.3, 2.9, rot=-5)
    photo(s, P("image124.jpeg"), 8.0, 3.6, 2.5, 2.9, rot=-6)
    photo(s, P("image6.jpeg"), 10.4, 3.95, 2.5, 2.9, rot=5)


# =============== 2. RINGKASAN EKSEKUTIF ===============
def s_exec():
    s = new()
    header(s, "Ringkasan Eksekutif", "Dari rencana tindak lanjut menjadi bukti yang siap akreditasi")
    a = B.A
    paras = [
        [("Kegiatan.  ", {"bold": True, "color": YELLOW}),
         ("Sanggabiz mengoordinasikan pendampingan luring Tim Pendamping Nasional Direktorat PPG di 12 LPTK untuk periode Oktober sampai Desember 2026. Kedua belas LPTK ini bukan kelompok baru. Mereka sudah dibina Tanoto Foundation sejak Supervisi 2025, jauh sebelum kami terlibat, sehingga program ini tidak dimulai dari nol. Kami melanjutkan apa yang sudah dirintis: menelaah pelaksanaan tindakan perbaikan, memperkuat kesiapan akreditasi, dan membuka peluang peningkatan mutu PPL, sesuai tujuan yang ditetapkan Tanoto Foundation sendiri.", {})],
        [("Hasil.  ", {"bold": True, "color": YELLOW}),
         ("Tiap LPTK menerima Laporan Pendampingan dan Paket Perbaikan Akreditasi paling lambat 14 hari setelah kunjungan: pemetaan kualifikasi dosen, telaah kurikulum OBE, pemetaan awal PPL, celah bukti akreditasi, dan RTL ber-penanggung jawab dan bertenggat. Di akhir program: studi komparatif terhadap kelompok LPTK pembanding, analisis lintas LPTK, dan rekomendasi pedoman pendampingan teknis (final 7 Desember 2026).", {})],
        [("Kebaruan.  ", {"bold": True, "color": YELLOW}),
         ("Lima fokus 2026 dan empat temuan Supervisi 2025 milik Tanoto Foundation sendiri menjadi kerangka kerja utama kami, diterjemahkan ke rubrik akreditasi LAMDIK/BAN-PT sebagai bahasa teknis. Evidence Register menelusuri setiap bukti. Rute berklaster menghemat biaya perjalanan sehingga LPTK tambahan dapat dilayani.", {})],
        [("Dana yang diminta.  ", {"bold": True, "color": YELLOW}),
         (f"{B.rp(a['total'])} (termasuk management fee dan pajak) untuk 12 LPTK inti + 2 LPTK tambahan; atau {B.rp(B.B['total'])} bila hanya 12 LPTK.", {"bold": True})],
    ]
    text(s, ML, 1.72, 7.95, 5.1, paras, size=13.5, spacing=1.1, after=9)
    stats = [("12", "LPTK di 5 provinsi, lanjutan\nSupervisi 2025"), ("5", "fokus 2026 Tanoto\nFoundation"),
             ("≤14", "hari: laporan tiap LPTK\nsetelah kunjungan"), ("7 Des", "Paket Laporan Akhir\n(final, 2026)")]
    for i, (n, l) in enumerate(stats):
        x, y = 8.95 + (i % 2) * 2.05, 1.85 + (i // 2) * 2.45
        card(s, x, y, 1.9, 2.25, fill=TEAL, line=WHITE)
        text(s, x, y + 0.22, 1.9, 0.9, n, size=32, color=WHITE, font=F_XB, align="c", anchor="m")
        text(s, x + 0.1, y + 1.2, 1.7, 0.9, l, size=11.5, bold=True, color=WHITE, align="c")


# =============== 3. PETA KESESUAIAN RFP ===============
def s_compliance():
    s = new()
    header(s, "Kemudahan penilaian", "Peta kesesuaian proposal dengan RFP dan kriteria penilaian")
    L = [["Ketentuan RFP", "Slide"],
         ["Ringkasan Eksekutif (1 halaman)", "2"],
         ["Teori Perubahan (Theory of Change) Tanoto Foundation", "5"],
         ["Deskripsi Proyek per tujuan", "9–12"],
         ["Pendekatan terdiferensiasi, penelaahan bukti akreditasi", "6, 8, 13"],
         ["Tindak lanjut hasil supervisi 2025", "4, 9, 14"],
         ["Pemetaan pelaksanaan PPL", "10"],
         ["Metode studi komparatif", "11, 15"],
         ["Pengelolaan data", "16"],
         ["Rute dan pengelolaan multilokasi", "17"]]
    R = [["Ketentuan RFP", "Slide"],
         ["Linimasa dan deliverables", "18–19"],
         ["Mitigasi risiko", "20"],
         ["Tata kelola (governance)", "21"],
         ["Tim dan personel kunci", "22"],
         ["Strategi keberlanjutan", "23"],
         ["Anggaran dan skema pembayaran", "24"],
         ["Pengalaman dan reputasi calon mitra", "25–28"],
         ["Komitmen kelembagaan dan kepatuhan", "29"]]
    table(s, ML, 1.85, [4.9, 1.05], L, size=11.5, row_h=0.47, head_h=0.42, aligns=["l", "c"], first_bold=False)
    table(s, ML + 6.1, 1.85, [4.9, 1.05], R, size=12.5, row_h=0.5, aligns=["l", "c"], first_bold=False)
    # bobot penilaian
    y = 6.45
    for i, (w, l, c) in enumerate([("50%", "Pendekatan teknis dan strategi implementasi", YELLOW),
                                   ("20%", "Pengalaman dan reputasi", TEAL), ("20%", "Kualitas dan kapabilitas tim", TEAL),
                                   ("10%", "Efisiensi anggaran dan jadwal", TEAL)]):
        x = ML + i * 3.03
        pill(s, x, y, 2.9, 0.42, fill=NAVY_DD, line=c, lw=1.5)
        text(s, x + 0.12, y, 0.75, 0.42, w, size=13, bold=True, color=YELLOW, anchor="m")
        text(s, x + 0.78, y, 2.1, 0.42, l, size=9, bold=True, color=WHITE, anchor="m", spacing=0.95, after=0)


# =============== 4. KONTEKS ===============
def s_context():
    s = new()
    header(s, "Pemahaman terhadap tujuan", "Konteks program: empat temuan Supervisi 2025 yang harus dijawab", tsize=25)
    items = [("169", "LPTK disupervisi\nmutu tahunan (2025)"),
             ("25", "LPTK disupervisi\nmendalam Direktorat PPG"),
             ("22", "dari 172 LPTK (2026) terpilih\npendampingan luring"),
             ("12", "LPTK didampingi Tanoto\nFoundation sejak 2025")]
    for i, (n, l) in enumerate(items):
        x = ML + i * 3.03
        card(s, x, 1.75, 2.85, 1.9, fill=TEAL, line=WHITE)
        text(s, x, 1.83, 2.85, 0.85, n, size=36, font=F_XB, align="c", anchor="m")
        text(s, x + 0.12, 2.66, 2.6, 0.9, l, size=11.5, bold=True, align="c")
    text(s, ML, 3.72, CW, 0.35, "Empat temuan utama Supervisi 2025 (Tanoto Foundation)", size=15, bold=True, color=YELLOW)
    findings = [("1", "Implementasi OBE masih terbatas", "Tujuan, desain perkuliahan, dan asesmen belum selaras secara konsisten."),
                ("2", "Pembelajaran inklusif dan berpusat pada mahasiswa belum konsisten", "Praktik pembelajaran belum mendorong partisipasi aktif, dukungan sesuai kebutuhan, dan keterlibatan mahasiswa."),
                ("3", "Mutu praktikum mengajar dan umpan balik belum konsisten", "Praktik mengajar, pendampingan, observasi kelas, dan umpan balik yang dapat ditindaklanjuti bervariasi antar-LPTK."),
                ("4", "Penjaminan mutu belum jadi dasar perbaikan sistematis", "Hasil evaluasi mutu belum ditindaklanjuti untuk memperbaiki penyelenggaraan PPG.")]
    for i, (n, t, d) in enumerate(findings):
        x = ML + (i % 2) * 6.08
        y = 4.05 + (i // 2) * 1.3
        card(s, x, y, 5.95, 1.22, fill=NAVY_DD, line=TEAL, lw=1.5, radius=0.18)
        text(s, x + 0.15, y + 0.1, 0.6, 0.5, n, size=24, font=F_XB, color=YELLOW)
        text(s, x + 0.75, y + 0.08, 5.05, 0.5, t, size=11.5, bold=True, spacing=0.95)
        text(s, x + 0.75, y + 0.63, 5.05, 0.55, d, size=9.5, color=GREY, spacing=0.98)
    text(s, ML, 6.68, CW, 0.32, "Sumber: TOR PPG Technical Assistance 2026 (Tanoto Foundation, dibagikan 25 Sep 2026). 12 LPTK sasaran adalah kohort yang sama dengan pembinaan luring Tanoto Foundation pada 2025, bukan LPTK baru.",
         size=8.8, color=GREY)


# =============== 5. TEORI PERUBAHAN ===============
def s_theory():
    s = new()
    header(s, "Kerangka logis program", "Teori Perubahan Tanoto Foundation, yang menjadi kerangka kerja kami", tsize=24)
    cols = [("Aktivitas", NAVY_DD, TEAL,
             ["Menyusun bersama pedoman dan modul pendampingan teknis", "Melaksanakan studi komparatif dengan dan tanpa dukungan Tanoto Foundation",
              "Memverifikasi kemajuan RTL pembinaan 2025", "Penguatan kapasitas teknis dan konsultasi langsung ke 12 LPTK"]),
            ("Luaran", NAVY_DD, TEAL,
             ["Pedoman dan modul pendampingan teknis", "Paket perbaikan akreditasi", "Analisis lintas LPTK dan rekomendasi berikutnya",
              "Laporan studi komparatif dan kemajuan RTL 2025"]),
            ("Outcome Antara", TEAL, WHITE,
             ["LPTK menunjukkan kemajuan melaksanakan tindakan prioritas hasil pembinaan 2025",
              "Direktorat PPG memiliki pendekatan pendampingan yang teruji dan bukti untuk mengarahkan dukungan berikutnya"]),
            ("Outcome", TEAL, WHITE,
             ["Kapasitas LPTK meningkat dalam memenuhi indikator utama akreditasi dan memperbaiki penyelenggaraan PPG"]),
            ("Dampak", YELLOW, NAVY_D,
             ["Kesiapan LPTK untuk reakreditasi meningkat dan menunjukkan kemajuan menuju Akreditasi Unggul"])]
    x = ML
    w = (CW - 0.4 * 4) / 5
    for i, (t, fc, tc, bl) in enumerate(cols):
        card(s, x, 1.85, w, 4.55, fill=fc, line=WHITE if fc != NAVY_DD else TEAL, lw=1.75, radius=0.2)
        text(s, x + 0.12, 2.0, w - 0.24, 0.65, t, size=13.5, bold=True, color=tc, spacing=0.95)
        text(s, x + 0.12, 2.65, w - 0.24, 3.6, bl, size=10, bullet="•", after=6, color=tc, spacing=1.0)
        if i < 4:
            text(s, x + w, 3.6, 0.4, 0.6, "→", size=20, color=YELLOW, bold=True, align="c")
        x += w + 0.4
    text(s, ML, 6.55, CW, 0.6,
         [[("Ini adalah kerangka logis milik Tanoto Foundation sendiri ", {"bold": True, "color": YELLOW}),
           ("(TOR PPG Technical Assistance 2026). Kami mengikat setiap kegiatan, luaran, dan indikator keberhasilan proposal ini langsung ke lima kolom di atas, bukan menciptakan kerangka baru.", {})]],
         size=11.5)


# =============== 6. FOKUS DAN PENDEKATAN ===============
def s_insight():
    s = new()
    header(s, "Pendekatan teknis · 50%", "Lima fokus 2026 Tanoto Foundation, dengan bahasa akreditasi sebagai penerjemah", tsize=23)
    rows = [["Fokus 2026 (bahasa Tanoto Foundation)", "Temuan Supervisi 2025 terkait", "Pendekatan kami", "Elemen akreditasi (LAMDIK)"],
            ["1  Peningkatan kualitas dosen", "Tidak ada temuan spesifik", "Memetakan kualifikasi dosen, identifikasi kebutuhan penguatan prioritas", "Kualitas dosen tetap"],
            ["2  Penguatan OBE", "Temuan 1: OBE masih terbatas", "Menelaah keselarasan tujuan, desain perkuliahan, asesmen capaian pembelajaran", "Asesmen capaian pembelajaran"],
            ["5  Optimalisasi PPL", "Temuan 2 dan 3: pembelajaran belum berpusat mahasiswa, praktikum dan umpan balik bervariasi", "Menelaah standar, peran, pendampingan, dan umpan balik PPL", "Pembelajaran mikro"],
            ["3  Penjaminan mutu berkelanjutan", "Temuan 4: penjaminan mutu belum jadi dasar perbaikan", "Menelaah pelaksanaan peningkatan mutu dan bukti pendukungnya", "Sistem Penjaminan Mutu Internal (SPMI)"],
            ["4  Karya inovatif dan publikasi", "Tidak ada temuan spesifik", "Mengidentifikasi potensi publikasi dan inovasi dosen, mahasiswa, guru pamong", "Publikasi dosen tetap"]]
    table(s, ML, 1.75, [2.55, 3.15, 3.85, 2.48], rows, size=10.8, head_size=10.8, row_h=0.78, head_h=0.56, first_bold=True)
    card(s, ML, 6.15, CW, 0.75, fill=NAVY_DD, line=YELLOW, lw=1.5, radius=0.2)
    text(s, ML + 0.25, 6.15, CW - 0.5, 0.75,
         [[("Kerangka utama proposal ini adalah bahasa 5 fokus dan 4 temuan Tanoto Foundation sendiri (TOR PPG Technical Assistance 2026). ", {"bold": True, "color": YELLOW}),
           ("Kolom kanan (elemen LAMDIK/BAN-PT, IAPSK 3.0) hanya penerjemah teknis ke bahasa akreditasi resmi. Detail ambang dan rubriknya ada di lampiran.", {})]],
         size=10.5, anchor="m")


# =============== 6. PERAN ===============
def s_roles():
    s = new()
    header(s, "Ruang lingkup mitra pelaksana", "Peran yang jelas: siapa melakukan apa")
    cols = [("Direktorat PPG", ["Menetapkan kebijakan dan Pedoman Pendampingan 2026", "Menetapkan Tim Pendamping Nasional (TPN, 2 orang per kunjungan)", "Pendampingan substantif ke LPTK", "Menerima rekomendasi penyempurnaan"]),
            ("Tanoto Foundation", ["Sponsor dan pengarah program", "Menyepakati kelompok pembanding dan akses data", "Mereviu laporan, memberi persetujuan hasil", "Menerima pembelajaran perbandingan LPTK"]),
            ("Sanggabiz (mitra pelaksana)", ["Menyiapkan perangkat: rencana kerja, instrumen, jadwal", "Mengoordinasikan kunjungan luring 12 LPTK dan kebutuhan operasional TPN", "Menghimpun bukti kemajuan RTL dan memantau", "Analisis, studi komparatif, pelaporan, repositori data"])]
    for i, (t, bl) in enumerate(cols):
        x = ML + i * 4.05
        card(s, x, 2.0, 3.9, 4.0, fill=TEAL if i == 2 else NAVY_DD, line=WHITE if i == 2 else TEAL, lw=2.25)
        tab(s, x + 0.35, 1.8, 3.2, t, size=12.5)
        text(s, x + 0.28, 2.5, 3.4, 3.5, bl, size=14.5, bullet="•", after=10)
    card(s, ML, 6.25, CW, 0.6, fill=NAVY_DD, line=YELLOW, lw=1.5, radius=0.3)
    text(s, ML + 0.3, 6.25, CW - 0.6, 0.6,
         [[("Batas peran: ", {"bold": True, "color": YELLOW}),
           ("kami tidak menggantikan TPN dan tidak menjanjikan pendampingan akreditasi substantif; kami memastikan pendampingan itu terkoordinasi, terdata, dan terbukti.", {})]],
         size=12, anchor="m")


# =============== 7. FRAMEWORK ===============
def s_framework():
    s = new()
    header(s, "Pendekatan pendampingan", "Kerangka kerja: siklus Plan-Do-Check-Act, selaras dengan PPEPP", tsize=25)
    steps = [("1", "Diagnosis", "Memahami baseline", ["Telaah supervisi 2025 dan RTL 12 LPTK", "Skor awal 1–4 pada 5 elemen", "Diagnosis kesiapan dokumen dan syarat awal konversi"], "Okt"),
             ("2", "Rencana", "Menyiapkan perangkat", ["Rencana kerja dan jadwal berklaster", "Instrumen: scorecard, observasi PPL, FGD, komparatif", "Kick-off dan penyelarasan"], "Okt"),
             ("3", "Pelaksanaan", "Menjalankan pendampingan", ["Koordinasi kunjungan TPN di 12 LPTK", "Himpun bukti, pantau RTL, konsultasi lanjutan", "Paket Perbaikan Akreditasi ≤14 hari"], "Okt–Nov"),
             ("4", "Evaluasi", "Menyimpulkan dan menyempurnakan", ["Verifikasi kemajuan, analisis lintas LPTK", "Studi komparatif dan rekomendasi", "Reviu bersama Tanoto Foundation dan Direktorat PPG"], "Nov–Des")]
    for i, (ch, nm, sub, bl, when) in enumerate(steps):
        x = ML + i * 3.03
        text(s, x, 1.7, 2.85, 1.0, ch, size=54, color=YELLOW, font=F_XB, anchor="m")
        text(s, x + 0.8, 1.9, 2.05, 0.4, nm, size=19, color=WHITE, font=F_XB)
        text(s, x + 0.8, 2.32, 2.05, 0.4, when, size=11, bold=True, color=YELLOW)
        if i < 3:
            text(s, x + 2.55, 1.95, 0.5, 0.6, "→", size=22, color=YELLOW, bold=True, align="c")
        card(s, x, 2.85, 2.85, 3.0, fill=TEAL, line=WHITE)
        text(s, x + 0.2, 3.0, 2.5, 0.7, sub, size=13.5, bold=True, color=WHITE, spacing=1.0)
        text(s, x + 0.2, 3.75, 2.5, 1.95, bl, size=12.5, bullet="•", after=7)
    card(s, ML, 6.05, CW, 0.78, fill=NAVY_DD, line=YELLOW, lw=1.5, radius=0.2)
    text(s, ML + 0.25, 6.05, CW - 0.5, 0.78,
         [[("Dasar kerangka: ", {"bold": True, "color": YELLOW}),
           ("siklus Plan-Do-Check-Act (Deming), sejalan dengan cara kerja SPMI (Sistem Penjaminan Mutu Internal) yang wajib dijalankan setiap perguruan tinggi menurut Permendikbudristek 53/2023. SPMI ini punya lima tahap yang disingkat PPEPP: menetapkan standar, melaksanakannya, mengevaluasi hasilnya, mengendalikan penyimpangan, lalu terus meningkatkan mutunya. Sistem yang sama yang kami pakai untuk menilai LPTK, kami terapkan juga pada cara kerja kami sendiri.", {})]],
         size=10.5, anchor="m")


# =============== 8-11. TUJUAN ===============
def s_goal(n, title, kegiatan, indikator, luaran, extra=None):
    s = new()
    header(s, f"Deskripsi proyek · Tujuan {n}", title, tsize=24)
    card(s, ML, 2.05, 2.3, 4.7, fill=TEAL, line=WHITE)
    text(s, ML, 2.2, 2.3, 1.6, str(n), size=90, font=F_XB, align="c", anchor="m", color=YELLOW)
    text(s, ML + 0.15, 3.85, 2.0, 0.5, "TUJUAN", size=12, bold=True, align="c", tracking=200)
    if extra:
        text(s, ML + 0.15, 4.4, 2.0, 2.3, extra, size=12, bold=True, align="c", color=WHITE, after=4)
    cols = [("Kegiatan", kegiatan, 4.55), ("Ukuran keberhasilan", indikator, 3.05), ("Luaran dan bukti", luaran, 1.9)]
    x = ML + 2.5
    widths = [4.35, 3.0, 2.2]
    for (t, bl, _), w in zip(cols, widths):
        tab(s, x, 1.9, min(w - 0.2, 2.3), t, fill=YELLOW, size=12)
        card(s, x, 2.22, w - 0.1, 4.53, fill=NAVY_DD, line=TEAL, lw=1.5, radius=0.22)
        text(s, x + 0.18, 2.5, w - 0.4, 4.2, bl, size=13.5, bullet="•", after=8)
        x += w


def s_goals():
    s_goal(1, "Mendata dan menganalisis kemajuan RTL hasil supervisi 2025",
           ["Telaah laporan supervisi 2025 dan RTL 12 LPTK", "Skor baseline 1–4 pada lima elemen, dengan status bukti tiap elemen",
            "Bangun RTL Tracker: rencana, berjalan, selesai, terbukti", "Konsultasi berkala dengan TPN dan LPTK; himpun bukti kemajuan",
            "Verifikasi kemajuan di akhir periode"],
           ["12 dari 12 LPTK punya skor baseline sebelum kunjungan", "Persentase RTL berstatus terbukti: awal vs akhir", "Perubahan skor per elemen dan per LPTK"],
           ["Laporan Awal", "RTL Tracker", "Evidence Register", "Skor baseline dan akhir"])
    s_goal(2, "Memetakan pelaksanaan PPL dan menjawab temuan pembelajaran berpusat mahasiswa",
           ["Telaah pedoman PPL, SK penugasan, instrumen observasi, rubrik penilaian", "FGD: DPL, guru pamong, mahasiswa PPG, pimpinan prodi (daring atau saat kunjungan)",
            "Skoring enam dimensi PPL dengan empat level, termasuk partisipasi aktif dan keterlibatan mahasiswa (Temuan 2 Supervisi 2025)",
            "Kolaborasi DPL dengan guru pamong dan siklus supervisi klinis (pra-observasi, observasi, umpan balik, refleksi), menjawab Temuan 3",
            "Dokumentasi praktik baik dan hambatan"],
           ["12 dari 12 LPTK terpetakan pada enam dimensi", "Minimal tiga kelompok informan per LPTK", "Praktik baik dan hambatan tercatat dengan bukti"],
           ["Laporan analisis implementasi PPL tiap LPTK", "Instrumen FGD", "Pemetaan lintas LPTK"],
           extra=["Standar pelaksanaan", "Pembagian peran", "Pendampingan", "Observasi", "Umpan balik", "Penilaian"])
    s_goal(3, "Membandingkan LPTK yang menerima dan tidak menerima dukungan Tanoto Foundation",
           ["Basis 2025: 12 LPTK Tanoto Foundation vs 13 LPTK Kemendikdasmen (TOR Annex 1)",
            "Basis 2026: 12 LPTK Tanoto Foundation vs 10 LPTK Kemendikdasmen, kedaluwarsa akreditasi 2027 (TOR Annex 2)",
            "Pencocokan berdasarkan peringkat, jenis PT, wilayah, skor baseline",
            "Analisis kemajuan disesuaikan baseline; faktor pembeda lewat analisis konfigurasional",
            "Keterbatasan dinyatakan terbuka; kelompok pembanding dikonfirmasi ulang bersama Tanoto Foundation dan Direktorat PPG"],
           ["Dua basis pembanding dikonfirmasi pada Laporan Awal", "Selisih kemajuan dilaporkan dengan rentang ketidakpastian", "Faktor pembeda kemajuan teridentifikasi dan berbukti"],
           ["Laporan studi komparatif dan kemajuan RTL Supervisi 2025", "Tabel keseimbangan kondisi awal"])
    s_goal(4, "Menghasilkan pembelajaran dan rekomendasi untuk siklus pendampingan berikutnya",
           ["Sintesis lintas LPTK: peta mutu, pola celah, praktik baik", "Rekomendasi perbaikan pedoman, modul, dan instrumen pendampingan",
            "Bahan presentasi untuk Tanoto Foundation, Direktorat PPG, pemangku kepentingan", "Serah terima repositori data dan bukti", "Log komentar: semua masukan Tanoto Foundation dan Direktorat diakomodasi"],
           ["Setiap rekomendasi tertaut ke temuan berbukti", "100% komentar Tanoto Foundation dan Direktorat tercatat dan dijawab", "Repositori diserahterimakan lengkap"],
           ["Laporan sintesis lintas LPTK", "Rekomendasi pedoman dan instrumen", "Bahan presentasi", "Repositori data dan bukti"])


# =============== 12. SCORECARD ===============
def s_scorecard():
    s = new()
    header(s, "Pendekatan teknis · penerjemah ke bahasa akreditasi", "Readiness Scorecard: rubrik LAMDIK sebagai lapisan kedua dari 5 fokus Tanoto Foundation", tsize=21)
    text(s, ML, 1.55, CW, 0.28,
         [[("Istilah: ", {"bold": True, "color": YELLOW}), ("DTPS adalah Dosen Tetap Program Studi. CPL adalah Capaian Pembelajaran Lulusan.", {})]],
         size=10, color=GREY)
    rows = [["Elemen", "Syarat perlu (PPG)", "Skor 4 (melampaui)", "Bukti kunci"],
            ["1  Kualitas DTPS", "≥20% doktor; ≥3 DTPS lektor kepala", "≥40% doktor dan ≥70% berjabatan lektor ke atas; dianalisis dampaknya", "Ijazah, SK jabatan"],
            ["2  Kurikulum (CPL)", "≥15% MK penciri dinilai CPL, dievaluasi, ditindaklanjuti", "≥20% MK penciri; evaluasi rinci dan komprehensif", "Laporan asesmen CPL, target vs capaian"],
            ["3  Pembelajaran mikro", "Lab lengkap, ≥2 praktik, 8 keterampilan, refleksi", "Lab terawat, ≥3 praktik per mahasiswa", "RPS, lembar observasi, rubrik, inventaris"],
            ["4  SPMI (PPEPP)", "Lima tahap siklus dilaksanakan", "Lima tahap berbasis TI, SN-Dikti dan standar PT", "Notulen, berita acara, laporan SPMI"],
            ["5  Publikasi DTPS", "≥20% DTPS Sinta 2/Scopus/WoS, 3 tahun", "≥20% dan analisis tren, faktor, dampak", "DOI, tautan, rekapitulasi"]]
    table(s, ML, 2.0, [2.2, 3.35, 3.5, 2.98], rows, size=11, head_size=11, row_h=0.63, head_h=0.44)
    text(s, ML, 5.85, 5.7, 1.1,
         [[("Skala 1–4 ", {"bold": True, "color": YELLOW}), ("dinilai asesor (expert judgement). Kami memakai skala yang sama pada baseline, pemantauan, dan akhir.", {})]],
         size=11.5)
    text(s, ML + 6.0, 5.85, 6.0, 1.1,
         [[("Catatan konversi: ", {"bold": True, "color": YELLOW}),
           ("hanya prodi PPG berperingkat Baik Sekali dengan nilai 321–360 yang dapat mengajukan konversi ke Unggul Tiga Tahun. Kesiapan tiap LPTK dipilah di Laporan Awal, agar target realistis.", {})]],
         size=11.5)


# =============== 13. EVIDENCE + GAP ===============
def s_evidence():
    s = new()
    header(s, "Pendekatan teknis · pengelolaan bukti dan tindak lanjut", "Evidence Register dan RTL Tracker")
    card(s, ML, 1.95, 5.9, 3.7, fill=TEAL, line=WHITE)
    icon(s, "database", ML + 0.3, 2.15, 0.55)
    text(s, ML + 1.0, 2.15, 4.7, 0.55, "Evidence Register", size=20, font=F_XB, anchor="m")
    text(s, ML + 0.3, 2.85, 5.4, 0.7, "Semua bukti tiap elemen tiap LPTK tercatat dan dapat ditelusuri.", size=12.5)
    fields = ["ID bukti", "Elemen", "Jenis dokumen", "Sumber", "Tanggal", "Tautan berkas", "Status", "Verifikator"]
    for i, f in enumerate(fields):
        pill(s, ML + 0.3 + (i % 2) * 2.7, 3.55 + (i // 2) * 0.48, 2.55, 0.38, fill=NAVY_DD, line=None)
        text(s, ML + 0.3 + (i % 2) * 2.7, 3.55 + (i // 2) * 0.48, 2.55, 0.38, f, size=11, bold=True, align="c", anchor="m")
    card(s, ML + 6.13, 1.95, 5.9, 3.7, fill=NAVY_DD, line=TEAL, lw=2.25)
    icon(s, "flag", ML + 6.43, 2.15, 0.55, color=YELLOW)
    text(s, ML + 7.13, 2.15, 4.7, 0.55, "RTL Tracker", size=20, font=F_XB, anchor="m")
    text(s, ML + 6.43, 2.85, 5.4, 0.7, "Setiap celah bukti menjadi tindakan yang bertanggung jawab dan bertenggat.", size=12.5)
    flow = ["Celah", "Tindakan", "PIC", "Tenggat", "Status", "Bukti penutup"]
    for i, f in enumerate(flow):
        x = ML + 6.43 + (i % 3) * 1.85
        y = 3.6 + (i // 3) * 0.75
        pill(s, x, y, 1.65, 0.45, fill=YELLOW)
        text(s, x, y, 1.65, 0.45, f, size=11.5, bold=True, color=NAVY_D, align="c", anchor="m")
    text(s, ML + 6.43, 5.1, 5.4, 0.5, "Status: rencana  →  berjalan  →  selesai  →  terbukti", size=11.5, bold=True, color=YELLOW)
    text(s, ML, 5.85, CW, 1.05,
         [[("Indikator status. ", {"bold": True, "color": YELLOW}),
           ("Rencana: belum ada tindakan. Berjalan: ada bukti awal, belum tuntas. Selesai: LPTK melapor tuntas, buktinya baru diklaim. Terbukti: mitra sudah memverifikasi bukti itu sendiri, sehingga statusnya naik jadi terverifikasi. Prinsip kami tetap sama: tidak ada bukti yang dibuat atau direkayasa, klaim tanpa dokumen dicatat sebagai diklaim, bukan terverifikasi.", {})]],
         size=11.5, spacing=1.05)


# =============== 15. STUDI KOMPARATIF ===============
def s_compare():
    s = new()
    header(s, "Pendekatan teknis · metode studi komparatif", "Dua basis pembanding sudah teridentifikasi dari TOR Tanoto Foundation", tsize=23)
    bases = [("Basis Supervisi 2025", "Mengukur kemajuan RTL sejak Supervisi 2025", "12 LPTK", "didanai dan dipantau Tanoto Foundation (RFP C.3)", "13 LPTK", "didanai Kemendikdasmen, disupervisi bersamaan 2025 (TOR Annex 1)"),
             ("Basis Pendampingan 2026", "Mengukur pendampingan teknis tahun berjalan", "12 LPTK", "sama, didanai Tanoto Foundation, kontinuitas dari 2025", "10 LPTK", "didanai Kemendikdasmen, dipilih dari 60 LPTK karena akreditasi kedaluwarsa 2027 (TOR Annex 2)")]
    for i, (t, sub, na, da, nb, db) in enumerate(bases):
        x = ML + i * 6.13
        card(s, x, 1.75, 5.95, 1.85, fill=NAVY_DD, line=TEAL, lw=1.5, radius=0.2)
        text(s, x + 0.22, 1.85, 5.5, 0.35, t, size=14, font=F_XB, color=YELLOW)
        text(s, x + 0.22, 2.2, 5.5, 0.3, sub, size=10.5, color=GREY)
        card(s, x + 0.22, 2.55, 2.7, 0.9, fill=TEAL, line=WHITE, lw=1.5, radius=0.12)
        text(s, x + 0.32, 2.6, 2.5, 0.4, na, size=17, font=F_XB)
        text(s, x + 0.32, 2.98, 2.5, 0.45, da, size=8.8, spacing=0.9)
        text(s, x + 3.05, 2.9, 0.4, 0.3, "vs", size=11, bold=True, color=YELLOW, align="c")
        card(s, x + 3.45, 2.55, 2.28, 0.9, fill=NAVY_P, line=YELLOW, lw=1.5, radius=0.12)
        text(s, x + 3.55, 2.6, 2.1, 0.4, nb, size=17, font=F_XB, color=YELLOW)
        text(s, x + 3.55, 2.98, 2.1, 0.45, db, size=8.8, spacing=0.9)
    steps = [("1", "Konfirmasi", "kelompok pembanding di atas dan akses data bersama Tanoto Foundation dan Direktorat"), ("2", "Cocokkan", "berdasar peringkat, jenis PT, wilayah, skor baseline"),
             ("3", "Hitung", "kemajuan disesuaikan baseline (selisih skor akhir − awal)"), ("4", "Uji", "sensitivitas non-parametrik untuk jumlah kecil"),
             ("5", "Jelaskan", "faktor pembeda: analisis konfigurasional (QCA)"), ("6", "Simpulkan", "kontribusi Tanoto Foundation: contribution analysis, terbuka pada penjelasan lain")]
    for i, (n, t, d) in enumerate(steps):
        x = ML + (i % 6) * 2.03
        y = 3.85
        card(s, x, y, 1.93, 1.1, fill=TEAL, line=WHITE, radius=0.1)
        text(s, x + 0.1, y + 0.06, 0.5, 0.4, n, size=18, font=F_XB, color=YELLOW)
        text(s, x + 0.1, y + 0.44, 1.75, 0.3, t, size=10.5, font=F_XB)
        text(s, x + 0.1, y + 0.72, 1.75, 0.36, d, size=7.8, spacing=0.9)
    card(s, ML, 5.15, 5.9, 1.65, fill=NAVY_DD, line=YELLOW, lw=1.5, radius=0.22)
    text(s, ML + 0.22, 5.22, 5.5, 0.35, "Keterbatasan yang kami nyatakan sejak awal", size=12.5, bold=True, color=YELLOW)
    text(s, ML + 0.22, 5.58, 5.5, 1.2, ["Jumlah kelompok kecil; tidak mengklaim kausalitas kuat", "Kriteria seleksi dua kelompok pembanding berbeda (basis 2025 vs 2026) → potensi confounding, dinyatakan terbuka"], size=11.5, bullet="•", after=4)
    card(s, ML + 6.13, 5.15, 5.9, 1.65, fill=NAVY_DD, line=TEAL, lw=1.5, radius=0.22)
    text(s, ML + 6.35, 5.22, 5.5, 0.35, "Sesuai batas RFP", size=12.5, bold=True, color=YELLOW)
    text(s, ML + 6.35, 5.58, 5.5, 1.2, ["Hanya data tersedia dan pengumpulan informasi daring, tanpa kunjungan lapangan tambahan", "Nama LPTK pembanding dari TOR 25 Sep 2026; dikonfirmasi ulang saat kick-off"], size=11.5, bullet="•", after=4)


# =============== 15. DATA ===============
def s_data():
    s = new()
    header(s, "Pendekatan teknis · pengelolaan data", "Data dikelola aman, terlacak, dan tidak dipublikasikan tanpa izin")
    cards = [("database", "Repositori terpusat", ["Akses berlapis menurut peran", "Penamaan dan versi baku", "Log akses dan pencadangan"]),
             ("lock", "Data pribadi", ["Hanya data yang benar-benar diperlukan", "Persetujuan (consent) untuk FGD dan rekaman", "Anonimisasi di laporan lintas LPTK"]),
             ("shield", "Kerahasiaan dan publikasi", ["NDA seluruh tim", "Tidak mempublikasikan data atau hasil tanpa persetujuan tertulis Tanoto Foundation dan Direktorat PPG"]),
             ("check", "Integritas bukti", ["Status verifikasi tiap bukti", "Jejak audit perubahan", "Tidak ada bukti direkayasa"])]
    for i, (ic, t, bl) in enumerate(cards):
        x = ML + i * 3.03
        card(s, x, 1.9, 2.85, 3.5, fill=TEAL, line=WHITE)
        icon(s, ic, x + 0.25, 2.1, 0.6)
        text(s, x + 0.25, 2.8, 2.4, 0.6, t, size=15, font=F_XB, spacing=1.0)
        text(s, x + 0.25, 3.5, 2.45, 1.9, bl, size=12.5, bullet="•", after=6)
    card(s, ML, 5.65, CW, 1.2, fill=NAVY_DD, line=TEAL, lw=1.5, radius=0.22)
    text(s, ML + 0.25, 5.72, 3.0, 0.35, "Struktur repositori", size=13, bold=True, color=YELLOW)
    folders = ["00 Panduan", "01 Baseline 2025", "02 LPTK (E1–E5)", "03 Instrumen", "04 Laporan", "05 Studi komparatif"]
    for i, f in enumerate(folders):
        pill(s, ML + 0.25 + i * 1.95, 6.2, 1.85, 0.42, fill=NAVY_P, line=TEAL, lw=1)
        text(s, ML + 0.25 + i * 1.95, 6.2, 1.85, 0.42, f, size=10.5, bold=True, align="c", anchor="m")


# =============== 16. RUTE ===============
def s_route():
    s = new()
    header(s, "Manajemen multilokasi · efisiensi anggaran", "12 LPTK dalam 5 klaster rute, dua rute berjalan paralel")
    cl = [("Medan", "Sumatera Utara", ["Universitas Islam Sumatera Utara", "Universitas Muhammadiyah Sumatera Utara", "Universitas HKBP Nommensen", "Universitas Negeri Medan"], "Rute Sumatera"),
          ("Pekanbaru", "Riau", ["Universitas Riau", "Universitas Islam Riau"], "Rute Sumatera"),
          ("Jambi", "Jambi", ["Universitas Jambi"], "Rute Sumatera"),
          ("Semarang – Salatiga – Kudus", "Jawa Tengah", ["Universitas PGRI Semarang", "Universitas Islam Sultan Agung", "Universitas Kristen Satya Wacana", "Universitas Muria Kudus"], "Rute Jawa–Kaltim"),
          ("Samarinda", "Kalimantan Timur", ["Universitas Mulawarman"], "Rute Jawa–Kaltim")]
    pos = [(0, 0), (1, 0), (2, 0), (0, 1), (1, 1)]
    w, h = 3.95, 2.35
    for (t, prov, lp, rt), (cx, cy) in zip(cl, pos):
        x = ML + cx * (w + 0.1)
        y = 1.85 + cy * (h + 0.15)
        card(s, x, y, w, h, fill=TEAL if rt == "Rute Sumatera" else NAVY_DD, line=WHITE if rt == "Rute Sumatera" else TEAL, lw=2)
        text(s, x + 0.2, y + 0.12, w - 0.4, 0.4, t, size=14, font=F_XB)
        text(s, x + 0.2, y + 0.5, w - 0.4, 0.3, prov, size=10.5, bold=True, color=YELLOW)
        text(s, x + 0.2, y + 0.85, w - 0.4, 1.5, lp, size=11.5, bullet="•", after=3, spacing=1.0)
    # efisiensi
    x = ML + 2 * (w + 0.1)
    y = 1.85 + (h + 0.15)
    card(s, x, y, w, h, fill=NAVY_DD, line=YELLOW, lw=2)
    text(s, x + 0.2, y + 0.12, w - 0.4, 0.4, "Dampak klaster", size=14, font=F_XB, color=YELLOW)
    text(s, x + 0.2, y + 0.55, w - 0.4, 1.8,
         [[("29 vs 36", {"bold": True, "color": YELLOW, "size": 16}), (" malam menginap", {})],
          [("5 vs 12", {"bold": True, "color": YELLOW, "size": 16}), (" perjalanan pulang-pergi", {})],
          [("Penghematan membiayai 2 LPTK tambahan", {"size": 11})]], size=12, after=4)
    text(s, ML, 6.72, CW, 0.3, "Lokasi sesuai RFP bagian C.3. Kunjungan berurutan dalam satu klaster; jadwal mengikuti penugasan TPN dan ketersediaan LPTK.", size=9.5, color=GREY)


# =============== 17. GANTT ===============
def s_gantt():
    s = new()
    header(s, "Linimasa proyek", "Jadwal indikatif 6 Oktober – 7 Desember 2026")
    weeks = ["5 Okt", "12", "19", "26", "2 Nov", "9", "16", "23", "30", "7 Des"]
    x0, lw_, cw = ML + 3.35, 3.35, 0.868
    y0 = 1.85
    for i, wk in enumerate(weeks):
        pill(s, x0 + i * cw + 0.02, y0, cw - 0.04, 0.38, fill=YELLOW if i in (2, 7, 9) else TEAL)
        text(s, x0 + i * cw, y0, cw, 0.38, wk, size=10, bold=True, color=NAVY_D if i in (2, 7, 9) else WHITE, align="c", anchor="m")
    rows = [("Persiapan, kick-off, baseline, instrumen", 0, 2, TEAL),
            ("Pendampingan luring: Rute Sumatera (Medan, Pekanbaru, Jambi)", 3, 5, YELLOW),
            ("Pendampingan luring: Rute Jawa–Kaltim (Semarang, Salatiga, Kudus, Samarinda)", 3, 5, YELLOW),
            ("Laporan Pendampingan tiap LPTK (≤14 hari setelah kunjungan)", 4, 7, TEAL),
            ("Pengumpulan data studi komparatif dan pemantauan RTL", 3, 7, TEAL),
            ("Analisis lintas LPTK dan studi komparatif", 6, 8, TEAL),
            ("Reviu, finalisasi, penyerahan", 8, 9, YELLOW)]
    for i, (lab, a, b, c) in enumerate(rows):
        y = y0 + 0.55 + i * 0.5
        rect(s, ML, y - 0.02, CW, 0.46, fill=NAVY_DD if i % 2 == 0 else NAVY_P)
        text(s, ML + 0.1, y - 0.02, 3.2, 0.46, lab, size=10, bold=True, anchor="m", spacing=0.92, after=0)
        pill(s, x0 + a * cw + 0.05, y + 0.06, (b - a + 1) * cw - 0.1, 0.3, fill=c)
    # milestone
    ms = [(2, "Laporan Awal\n(target 23 Okt)"), (7, "Draf Laporan Akhir\n(target 27 Nov)"), (9, "Final\n7 Des")]
    for i, lab in ms:
        cx = x0 + i * cw + cw / 2
        d = s.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(cx - 0.17), Inches(5.95), Inches(0.34), Inches(0.34))
        d.fill.solid(); d.fill.fore_color.rgb = rgb(YELLOW); d.line.fill.background()
        text(s, cx - 1.0, 6.3, 2.0, 0.6, lab, size=10, bold=True, align="c", color=WHITE, spacing=1.0, after=0)
    text(s, ML, 5.9, 3.2, 0.5, "Tonggak", size=13, bold=True, color=YELLOW)
    text(s, ML, 6.35, 3.2, 0.6, "Jadwal kunjungan menyesuaikan penugasan TPN dan ketersediaan LPTK.", size=9.5, color=GREY)


# =============== 18. LUARAN ===============
def s_deliverables():
    s = new()
    header(s, "Linimasa dan deliverables", "Tiga luaran utama dan waktunya")
    rows = [["Luaran", "Isi", "Waktu"],
            [[("1  Laporan Awal", {"bold": True})], "Telaah hasil supervisi 2025, rencana kerja, pedoman dan modul pendampingan, pemetaan jadwal TPN pada 12 LPTK, desain dan instrumen evaluasi dan studi komparatif, instrumen FGD PPL", "Sebelum kunjungan pertama (target 23 Okt 2026)"],
            [[("2  Laporan Pendampingan dan Paket Perbaikan Akreditasi (12 LPTK)", {"bold": True})], "Pemetaan dosen dan guru pamong, telaah kurikulum OBE, pemetaan awal microteaching dan PPL, publikasi/karya inovatif, SPMI/PPEPP, kesenjangan bukti akreditasi, rekomendasi, RTL", "Maksimal 14 hari setelah tiap kunjungan (target internal 10 hari)"],
            [[("3  Paket Laporan Akhir", {"bold": True})], "(1) Studi komparatif dan kemajuan RTL Supervisi 2025; (2) analisis implementasi PPL tiap LPTK; (3) sintesis lintas LPTK dan rekomendasi; (4) rekomendasi pedoman dan instrumen, bahan presentasi, repositori data dan bukti", "Draf November (target 27 Nov); final 7 Desember 2026"]]
    table(s, ML, 1.85, [3.0, 6.0, 3.03], rows, size=12, row_h=1.05, head_h=0.42, first_bold=False)
    text(s, ML, 5.7, 9, 0.3, "Prinsip mutu yang kami terapkan pada setiap luaran", size=12.5, bold=True, color=YELLOW)
    chips = ["Format disepakati", "Kondisi awal ≠ terkini ≠ rekomendasi", "Bukti dapat ditelusuri", "PIC dan tenggat tiap RTL", "Kuantitatif dan kualitatif", "PDF final + versi editable", "Semua komentar diakomodasi"]
    x, y = ML, 6.1
    for c in chips:
        w = 0.28 + 0.088 * len(c)
        if x + w > SW - ML:
            x, y = ML, y + 0.42
        pill(s, x, y, w, 0.34, fill=NAVY_DD, line=TEAL, lw=1)
        text(s, x, y, w, 0.34, c, size=9.5, bold=True, align="c", anchor="m")
        x += w + 0.1


# =============== 19. RISIKO ===============
def s_risk():
    s = new()
    header(s, "Pendekatan teknis · mitigasi risiko", "Risiko utama dan cara kami mengelolanya")
    rows = [["Risiko", "Dampak", "Mitigasi"],
            ["Jadwal TPN, Direktorat, atau LPTK berubah", "Kunjungan bergeser, laporan terlambat", "Kalender bersama, tanggal cadangan, dua rute paralel dan berklaster"],
            ["Dokumen atau data LPTK terlambat", "Analisis lemah", "Paket kesiapan dokumen H-7, pengingat H-3, status bukti diklaim vs terverifikasi"],
            ["Kondisi awal LPTK beragam", "Perbandingan tidak adil", "Skor baseline, matching, keterbatasan dinyatakan"],
            ["Data pembanding terbatas", "Studi komparatif dangkal", "Minta akses data sejak kick-off; analisis konfigurasional dan contribution analysis"],
            ["Beban puncak November", "Laporan tertunda", "Rencana mundur dari 7 Desember, tim laporan paralel, template siap sebelum kunjungan"],
            ["Data pribadi dan kerahasiaan", "Pelanggaran ketentuan", "Minimasi data, akses berlapis, NDA, izin tertulis sebelum publikasi"],
            ["Kendala perjalanan", "Biaya dan jadwal", "Tiket fleksibel, kontingensi, klaster rute"]]
    table(s, ML, 1.85, [3.55, 2.95, 5.53], rows, size=11.5, row_h=0.6, head_h=0.42)


# =============== 20. TATA KELOLA ===============
def s_gov():
    s = new()
    header(s, "Tata kelola proyek", "Struktur pengawasan dan pengambilan keputusan")
    tiers = [("Steering Committee", "Kemendikdasmen (Direktorat PPG) · Tanoto Foundation · Direktur Sanggabiz", "Arah strategis; keputusan lingkup, jadwal, anggaran. Rapat: kick-off, reviu tengah (awal Nov), penutupan.", YELLOW, NAVY_D),
             ("Technical Committee", "Perwakilan Direktorat PPG/TPN · unsur Kemdiktisaintek · Project Manager · Advisor · Analis", "Memastikan mutu luaran, instrumen, jadwal kunjungan; keselarasan dengan revitalisasi LPTK dan persyaratan akreditasi. Rapat mingguan 30 menit.", TEAL, WHITE),
             ("Advisory Board", "Ahli PPG, akreditasi/penjaminan mutu, evaluasi", "Reviu metodologi dan draf laporan; dua sesi (baseline dan draf akhir).", NAVY_DD, WHITE)]
    for i, (t, who, tor, fc, tc) in enumerate(tiers):
        y = 1.85 + i * 1.5
        card(s, ML, y, 7.3, 1.35, fill=fc, line=WHITE if fc != NAVY_DD else TEAL, lw=2)
        text(s, ML + 0.25, y + 0.1, 6.8, 0.4, t, size=16, font=F_XB, color=tc)
        text(s, ML + 0.25, y + 0.5, 6.8, 0.35, who, size=10.5, bold=True, color=tc, spacing=0.95)
        text(s, ML + 0.25, y + 0.85, 6.8, 0.5, tor, size=10, color=tc, spacing=0.98)
    card(s, ML, 6.35, 7.3, 0.5, fill=NAVY_P, line=TEAL, lw=1.5, radius=0.2)
    text(s, ML + 0.25, 6.35, 6.8, 0.5, "Tim pelaksana: Project Manager, koordinator lapangan (2), analis data, administrator dokumentasi", size=11, bold=True, anchor="m")
    x = ML + 7.6
    card(s, x, 1.85, 4.4, 5.0, fill=NAVY_DD, line=TEAL, lw=1.5)
    text(s, x + 0.25, 2.0, 3.9, 0.4, "Mekanisme kerja", size=15, font=F_XB, color=YELLOW)
    text(s, x + 0.25, 2.5, 3.95, 4.3,
         ["Satu pintu komunikasi: Project Manager ↔ narahubung Tanoto Foundation dan Direktorat", "Laporan kemajuan mingguan singkat (status, risiko, keputusan yang diminta)",
          "Koordinasi dengan Kemdiktisaintek dilakukan melalui Direktorat PPG, bukan langsung, sesuai struktur tim pendampingan bersama",
          "Setiap perubahan lingkup dan alokasi anggaran memerlukan persetujuan tertulis Tanoto Foundation", "Reviu dua lapis sebelum laporan dikirim", "ToR lengkap tiap badan pada lampiran"],
         size=11, bullet="•", after=6)


# =============== 21. TIM ===============
def s_team():
    s = new()
    header(s, "Kualitas dan kapabilitas tim · 20%", "Struktur tim dan pembagian peran")
    roles = [("Direktur Proyek", "Steering, hubungan dengan Tanoto Foundation dan Direktorat PPG, jaminan mutu"),
             ("Project Manager dan Koordinator Nasional", "Jadwal, koordinasi TPN, kontrol mutu, laporan mingguan"),
             ("Ahli Akreditasi dan Penjaminan Mutu PPG", "Rubrik, Paket Perbaikan Akreditasi, reviu laporan"),
             ("Ahli OBE, Kurikulum, dan PPL", "Instrumen CPL, pemetaan PPL dan microteaching, FGD"),
             ("Analis Data dan Studi Komparatif", "Matching, analisis kemajuan, dashboard, tabel keseimbangan"),
             ("Koordinator Lapangan (2 orang)", "Rute Sumatera; Rute Jawa–Kaltim; logistik dan kunjungan"),
             ("Administrator Data dan Dokumentasi", "Evidence Register, repositori, notulen, log komentar")]
    for i, (r, d) in enumerate(roles):
        col, row = i % 2, i // 2
        x = ML + col * 6.08
        y = 1.85 + row * 1.25
        card(s, x, y, 5.95, 1.12, fill=TEAL if i in (0, 1) else NAVY_DD, line=WHITE if i in (0, 1) else TEAL, lw=1.75, radius=0.2)
        text(s, x + 0.2, y + 0.08, 5.55, 0.35, r, size=13, font=F_XB)
        text(s, x + 0.2, y + 0.43, 5.55, 0.35, d, size=11.5, color=WHITE)
        text(s, x + 0.2, y + 0.78, 5.55, 0.3, "Nama dan CV: [diisi, lihat Lampiran CV Personel Kunci]", size=9.5, italic=True, color=YELLOW)
    # sisa ruang kanan bawah
    card(s, ML + 6.08, 1.85 + 3 * 1.25, 5.95, 1.12, fill=NAVY_P, line=YELLOW, lw=1.5, radius=0.2)
    text(s, ML + 6.3, 1.85 + 3 * 1.25 + 0.05, 5.55, 1.05,
         [[("Kepakaran PPG dan akreditasi ", {"bold": True, "color": YELLOW}),
           ("dipastikan lewat Advisor dan Advisory Board; CV personel kunci disajikan di bagian terpisah sesuai RFP.", {})]],
         size=11, anchor="m")


# =============== 22. KEBERLANJUTAN ===============
def s_sustain():
    s = new()
    header(s, "Strategi keberlanjutan", "Hasil program tetap bekerja setelah Desember 2026")
    cards = [("database", "Repositori dan dashboard", "Data, bukti, skor, dan tracker diserahterimakan ke Tanoto Foundation dan Direktorat PPG dalam format terbuka."),
             ("book", "Pedoman dan instrumen", "Rubrik, instrumen observasi PPL, FGD, dan templat laporan yang disempurnakan siap dipakai ulang."),
             ("users", "Komunitas praktik antar-LPTK", "Berbagi praktik baik 12 LPTK lewat sesi daring; memanfaatkan pertemuan yang sudah ada."),
             ("layers", "Replikasi ke LPTK lain", "Playbook ringkas Akreditasi-Ready dan rekomendasi tahapan untuk 172 LPTK.")]
    for i, (ic, t, d) in enumerate(cards):
        x = ML + (i % 2) * 6.08
        y = 1.9 + (i // 2) * 2.55
        card(s, x, y, 5.95, 2.4, fill=TEAL if (i + i // 2) % 2 == 0 else NAVY_DD, line=WHITE if (i + i // 2) % 2 == 0 else TEAL, lw=2)
        icon(s, ic, x + 0.3, y + 0.3, 0.6, color=YELLOW if (i + i // 2) % 2 else WHITE)
        text(s, x + 1.1, y + 0.3, 4.6, 0.6, t, size=16, font=F_XB, anchor="m")
        text(s, x + 0.3, y + 1.1, 5.4, 1.25, d, size=14.5)


# =============== 23. ANGGARAN ===============
def s_budget():
    s = new()
    a, b = B.A, B.B
    header(s, "Efisiensi anggaran dan kelayakan jadwal · 10%", "Anggaran ringkas dan skema pembayaran")
    L = a["lines"]
    rows = [["Komponen (14 LPTK: 12 inti + 2 tambahan)", "Asumsi", "Jumlah"],
            ["Honorarium TPN", f"2 orang × 14 kunjungan × {B.rp(B.HONOR_TPN)}", B.rp(L["honor"])],
            ["Akomodasi", f"{a['nights']} malam × 3 orang × {B.rp(B.RATE_NIGHT)}", B.rp(L["akom"])],
            ["Tiket pulang-pergi", f"{a['trips']} perjalanan × 3 orang × {B.rp(B.TICKET_RT)}", B.rp(L["tiket"])],
            ["Transportasi lokal", f"14 kunjungan × 3 hari × {B.rp(B.LOCAL_DAY)}", B.rp(L["lokal"])],
            ["Uang harian dan logistik", f"14 × 3 orang × 3 hari × {B.rp(B.DAILY)}", B.rp(L["harian"])],
            ["Platform, repositori, alat data", "Paket", B.rp(L["tools"])],
            ["Kontingensi perjalanan", "±2% dari subtotal", B.rp(L["kontingensi"])],
            [[("Subtotal perjalanan dan logistik", {"bold": True, "color": YELLOW})], f"{a['perj_pct']:.1f}%", [(B.rp(a["perj"]), {"bold": True, "color": YELLOW})]],
            ["Management fee", "termasuk honorarium mitra", B.rp(a["fee"])],
            ["PPN 11% atas fee", "", B.rp(a["ppn"])],
            [[("TOTAL", {"bold": True, "color": NAVY_D})], f"fee + pajak {a['fee_pct']:.1f}%", [(B.rp(a["total"]), {"bold": True, "color": NAVY_D})]]]
    gs = table(s, ML, 1.75, [3.7, 3.4, 1.7], rows, size=10.5, head_size=11, row_h=0.36, head_h=0.4, aligns=["l", "l", "r"], first_bold=False)
    tbl = gs.table
    for c in range(3):
        cell = tbl.cell(len(rows) - 1, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(YELLOW)
        for para in cell.text_frame.paragraphs:
            for r_ in para.runs:
                r_.font.color.rgb = rgb(NAVY_D)
    x = ML + 9.0
    card(s, x, 1.75, 3.03, 2.05, fill=TEAL, line=WHITE)
    text(s, x + 0.15, 1.85, 2.75, 0.3, "Usulan utama", size=11, bold=True, color=YELLOW)
    text(s, x + 0.15, 2.15, 2.75, 0.5, B.rp(a["total"]), size=17, font=F_XB)
    text(s, x + 0.15, 2.7, 2.75, 1.1, "14 LPTK (12 inti + 2 tambahan; lokasi tambahan diberitahukan Tanoto Foundation)", size=10.5)
    card(s, x, 3.95, 3.03, 1.6, fill=NAVY_DD, line=TEAL, lw=1.5)
    text(s, x + 0.15, 4.03, 2.75, 0.3, "Alternatif 12 LPTK", size=11, bold=True, color=YELLOW)
    text(s, x + 0.15, 4.33, 2.75, 0.45, B.rp(b["total"]), size=15, font=F_XB)
    text(s, x + 0.15, 4.85, 2.75, 0.7, "Sisa pagu tidak digunakan", size=10.5)
    card(s, x, 5.7, 3.03, 1.15, fill=NAVY_P, line=YELLOW, lw=1.5)
    text(s, x + 0.15, 5.75, 2.75, 1.1,
         [[("Termin (usulan)", {"bold": True, "color": YELLOW, "size": 11})],
          [("30% Laporan Awal disetujui · 40% laporan 12 LPTK · 30% paket akhir final", {"size": 10})]], size=10, after=2)
    text(s, ML, 6.42, 8.9, 0.5,
         "Harga satuan tiket, honorarium TPN (mengacu SBM Kementerian), dan uang harian adalah asumsi awal; disesuaikan dengan harga aktual dan kesepakatan TF. Pembayaran diproses maks 14 hari kerja setelah dokumen lengkap. Rincian di Proposal Anggaran.",
         size=9, color=GREY)


# =============== 24. PROFIL ===============
def s_profile():
    s = new()
    header(s, "Pengalaman dan reputasi calon mitra · 20%", "Sanggabiz dalam angka")
    text(s, ML, 1.62, CW, 0.6, "Solusi bisnis virtual berbasis department-as-a-service: keuangan, SDM, layanan pelanggan, pemasaran, manajemen proyek, administrasi, dan operasional. Berkantor di Yogyakarta.",
         size=12, color=GREY)
    nums = [("6.000+", "Mitra engagement", "sektor swasta dan pemerintah", "handshake"), ("400+", "Proyek dan layanan", "terlaksana", "layers"),
            ("500+", "Tim", "ahli dan trainer", "people"), ("100+", "Komunitas mitra", "", "users"), ("400+", "Channel media partner", "", "link"),
            ("11.000+", "Total pengguna", "B2C, B2B, B2G", "target")]
    icons = ["users", "layers", "people", "users", "link", "target"]
    for i, (n, l, sub, _) in enumerate(nums):
        x = ML + (i % 3) * 4.05
        y = 2.5 + (i // 3) * 2.15
        card(s, x, y, 3.9, 1.95, fill=TEAL if i in (0, 1, 2) else NAVY_DD, line=WHITE if i in (0, 1, 2) else TEAL, lw=2)
        icon(s, icons[i], x + 0.25, y + 0.35, 0.65)
        text(s, x + 1.05, y + 0.25, 2.7, 0.8, n, size=34, font=F_XB, anchor="m")
        text(s, x + 0.25, y + 1.25, 3.5, 0.35, l, size=13, bold=True)
        if sub:
            text(s, x + 0.25, y + 1.55, 3.5, 0.3, sub, size=11, color=WHITE)
    text(s, ML, 6.8, CW, 0.25, "Sumber: Company Profile Sanggabiz 2026.", size=9, color=GREY)


# =============== 25-27. PENGALAMAN ===============
def s_exp(photo_file, title, subtitle, blocks, rel, focus=(0.5, 0.5)):
    s = new(foot=False)
    photo(s, P(photo_file), 0, 0, 6.6, SH, rounded=False, focus=focus)
    rect(s, 0, 0, 6.6, SH, fill=NAVY_D, alpha=72)
    rect(s, 6.6, 0, SW - 6.6, SH, fill=NAVY_P)
    text(s, 0.5, 0.38, 5.7, 0.3, "Pengalaman relevan", size=11.5, bold=True, color=YELLOW, caps=True, tracking=120)
    text(s, 0.5, 0.68, 5.7, 1.0, title, size=24, font=F_XB, spacing=1.0)
    text(s, 0.5, 1.6, 5.7, 0.35, subtitle, size=11.5, color=WHITE)
    y = 2.1
    for (nm, l1, l2, l3) in blocks:
        text(s, 0.5, y, 5.8, 0.35, nm, size=13.5, bold=True, color=YELLOW)
        rows = [("calendar", l1, 0.36), ("people", l2, 0.36), ("target", l3, 0.5)]
        yy = y + 0.4
        for ic, t, hh in rows:
            if not t:
                continue
            icon(s, ic, 0.55, yy + 0.03, 0.28)
            text(s, 1.0, yy, 5.3, hh, t, size=11, anchor="t", spacing=0.95, after=0)
            yy += hh
        y += 1.6
    text(s, 7.1, 0.5, 5.6, 0.4, "Relevansi terhadap RFP", size=15, bold=True, color=YELLOW, caps=True, tracking=80)
    yy = 1.1
    for h, d in rel:
        card(s, 7.1, yy, 5.6, 1.6, fill=NAVY_DD, line=TEAL, lw=1.5, radius=0.2)
        text(s, 7.35, yy + 0.12, 5.1, 0.35, h, size=14, font=F_XB)
        text(s, 7.35, yy + 0.55, 5.15, 1.0, d, size=12.5, spacing=1.05)
        yy += 1.8
    logo(s, 7.1, 6.95, 0.26)
    text(s, SW - ML - 1.0, 7.06, 1.0, 0.25, str(N[0]), size=10, bold=True, color=YELLOW, align="r")


def s_exps():
    s_exp("image112.jpeg", "Manajemen program multilokasi", "Pelatihan dan pengembangan lintas kota dan lembaga",
          [("MSME Training Rumah BUMN", "Webinar dan pelatihan UMKM, 2025–2026", "Jejaring Rumah BUMN di banyak daerah", "Surabaya, Purbalingga, Labuan Bajo, Yogyakarta, Jakarta, Bojonegoro, Kabanjahe, Samarinda, Pekalongan, dst."),
           ("TechPath Bojonegoro (Plan International)", "Desember 2025 – Maret 2026", "69 peserta (individu)", "Digital marketing, content creation, soft skill dan job readiness"),
           ("FYEP: Retail Skills for Youth Empowerment", "Program berjalan (Plan International)", "Model P-L-A-N: Prepare, Learn, Apply, Navigate", "")],
          [("Koordinasi banyak lokasi", "Menjangkau jejaring Rumah BUMN di 9+ kota, termasuk Samarinda (Kalimantan Timur), dengan koordinasi jarak jauh."),
           ("Kemitraan lembaga pembangunan", "Bekerja dengan yayasan dan organisasi pembangunan (Plan International), lengkap dengan pelaporan program."),
           ("Program end-to-end", "Dari desain, pelaksanaan, hingga pendampingan peserta; sejalan dengan siklus kunjungan dan tindak lanjut.")],
          focus=(0.5, 0.4))
    s_exp("image117.jpeg", "Class management skala besar", "Administrasi kelas, peserta, dan pelaporan berkelanjutan",
          [("Class Management LingoTalk", "Juli 2023 – sekarang", "1000+ peserta · 350+ kelas Business English", "28+ institusi pemerintah dan non-pemerintah"),
           ("Program Pra Kerja", "Juli–November 2023 dan 2024", "±4.000 peserta", "Customer service, akuntansi UMKM, kewirausahaan"),
           ("Amartha Cendekia", "Program berjalan", "End-to-end class management", "")],
          [("Koordinasi jadwal dan peserta", "Pendaftaran, pengingat, administrasi kelas, dan dokumen pasca-kelas."),
           ("Institusi pemerintah", "Melayani institusi pemerintah dan non-pemerintah dalam skala ratusan kelas."),
           ("Disiplin pelaporan", "Data peserta, kehadiran, dan hasil dikelola konsisten; dasar untuk Evidence Register.")],
          focus=(0.5, 0.5))
    s_exp("image137.jpeg", "Kerja sama instansi dan pengembangan modul", "Regulator, BUMD, dan bahan ajar",
          [("Otoritas Jasa Keuangan (OJK) Yogyakarta", "April 2025", "50 peserta (korporat)", "Digital marketing bagi lembaga keuangan yang patuh regulasi OJK"),
           ("PT Riau Petroleum Rokan", "November 2025 dan April 2026", "8, 30, dan 19 peserta", "Bisnis migas terintegrasi; pelatihan BNSP dan ISO"),
           ("Pengembangan modul dan video", "Modul STEM, Leadership for First-Time Manager, FYEP", "16 video pembelajaran (FYEP)", "")],
          [("Kerja dengan regulator dan BUMD", "Pengalaman menyiapkan materi dan sesi bagi instansi pemerintah dan lembaga berstandar."),
           ("Standar dan mutu", "Pelatihan BNSP dan ISO (manajemen risiko, audit internal, sertifikasi) selaras dengan disiplin standar dan bukti."),
           ("Pengembangan pedoman dan modul", "Modul, materi presentasi, dan video: relevan untuk penyempurnaan pedoman dan instrumen.")],
          focus=(0.5, 0.5))


# =============== 28. KOMITMEN ===============
def s_commit():
    s = new()
    header(s, "Persyaratan kelembagaan (RFP bagian D)", "Komitmen kepatuhan dan etika")
    items = ["Memiliki NPWP dan rekening atas nama lembaga", "Tidak ada catatan pelanggaran hak anak, kekerasan, pelecehan, diskriminasi, atau eksploitasi anak",
             "Tidak terlibat, berafiliasi, atau menerima pendanaan dari partai politik dan organisasi terlarang", "Tidak menerima pendanaan dari perusahaan rokok, perjudian, atau susu formula selama kemitraan",
             "Mematuhi kebijakan perlindungan anak, antikorupsi, konflik kepentingan, dan kode etik Tanoto Foundation", "Menjaga kerahasiaan data dan membatasi data pribadi pada yang benar-benar diperlukan",
             "Tidak mempublikasikan data atau hasil program tanpa persetujuan tertulis TF dan Direktorat PPG", "Bersedia mematuhi prosedur dan kebijakan Tanoto Foundation"]
    for i, t in enumerate(items):
        x = ML + (i % 2) * 6.08
        y = 1.85 + (i // 2) * 1.2
        card(s, x, y, 5.95, 1.08, fill=NAVY_DD, line=TEAL, lw=1.5, radius=0.2)
        icon(s, "check", x + 0.22, y + 0.28, 0.5, color=YELLOW)
        text(s, x + 0.95, y + 0.05, 4.85, 0.98, t, size=11.5, anchor="m", spacing=1.02)
    text(s, ML, 6.75, CW, 0.3, "Bukti pendukung disampaikan setelah nominasi oleh Tanoto Foundation (RFP bagian D, catatan kaki 2).", size=9.5, color=GREY)


# =============== 29. PENUTUP ===============
def s_close():
    s = new(foot=False)
    text(s, ML, 1.6, 9.5, 1.6, "Siap berjalan bersama Tanoto Foundation dan Direktorat PPG", size=38, font=F_XB, spacing=1.0)
    text(s, ML, 3.35, 9.5, 0.8, "Mengubah rencana tindak lanjut menjadi bukti, dan bukti menjadi Akreditasi Unggul.", size=16, font=F_M, color=GREY)
    logo(s, ML, 4.35, 0.75)
    for i, (ic, t) in enumerate((("link", "www.sanggabiz.com"), ("target", "@sanggabiz"), ("doc", "partnership@sanggabiz.com"))):
        x = ML + i * 3.85
        pill(s, x, 5.55, 3.7, 0.52, fill=NAVY_DD, line=WHITE, lw=1.5)
        icon(s, ic, x + 0.2, 5.65, 0.32)
        text(s, x + 0.65, 5.55, 3.0, 0.52, t, size=12, bold=True, anchor="m")
    text(s, ML, 6.3, 11, 0.6, [[("Yogyakarta  |  ", {"bold": True}), ("Jl. Bambu Gading No.111, Sendon, Tirtoadi, Kec. Mlati, Kab. Sleman, DI Yogyakarta 55287", {})]], size=10.5, color=GREY)
    text(s, ML, 6.7, 11, 0.3, "Kontak proyek: [nama, jabatan, telepon: diisi]", size=10.5, italic=True, color=YELLOW)


s_cover(); s_exec(); s_compliance(); s_context(); s_theory(); s_insight(); s_roles(); s_framework(); s_goals()
s_scorecard(); s_evidence(); s_compare(); s_data(); s_route(); s_gantt(); s_deliverables(); s_risk(); s_gov(); s_team()
s_sustain(); s_budget(); s_profile(); s_exps(); s_commit(); s_close()
assert len(prs.slides) == 30, len(prs.slides)
prs.save(OUT)
print("saved", OUT, len(prs.slides))
