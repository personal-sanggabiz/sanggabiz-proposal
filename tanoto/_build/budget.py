"""Model anggaran (satu sumber angka untuk deck dan XLSX). Semua harga satuan = ASUMSI yang harus diverifikasi."""

PAGU = 510_000_000
PERSONS = 3                 # 2 TPN + 1 personel mitra (RFP E)
RATE_NIGHT = 850_000        # RFP E
NIGHTS_STD = 3              # RFP E
CLUSTERS = [("Medan", 4), ("Pekanbaru", 2), ("Jambi", 1), ("Semarang–Salatiga–Kudus", 4), ("Samarinda", 1)]
CLUSTER_EXTRA_NIGHTS = 2    # tiap LPTK tambahan pada klaster yang sama
TICKET_RT = 3_000_000       # ASUMSI rata-rata tiket PP per orang
LOCAL_DAY = 900_000         # ASUMSI sewa mobil+driver+BBM per hari
LOCAL_DAYS = 3
DAILY = 350_000             # ASUMSI uang harian/konsumsi/logistik per orang per hari
HONOR_TPN = 7_500_000       # ASUMSI honorarium TPN per orang per kunjungan (acu SBM)
N_TPN = 2
TOOLS = 9_900_000           # platform daring, repositori, alat pengumpulan data
PPN = 0.11


def scenario(extras, fee_net=None, balance=False):
    n_core = sum(k for _, k in CLUSTERS)
    n = n_core + extras
    nights_core = sum(NIGHTS_STD + CLUSTER_EXTRA_NIGHTS * (k - 1) for _, k in CLUSTERS)
    nights = nights_core + extras * NIGHTS_STD
    trips = len(CLUSTERS) + extras
    L = {}
    L["honor"] = N_TPN * n * HONOR_TPN
    L["akom"] = nights * PERSONS * RATE_NIGHT
    L["tiket"] = trips * PERSONS * TICKET_RT
    L["lokal"] = n * LOCAL_DAYS * LOCAL_DAY
    L["harian"] = n * PERSONS * LOCAL_DAYS * DAILY
    L["tools"] = TOOLS
    base = sum(L.values())
    if balance:
        fee = fee_net
        ppn = round(fee * PPN)
        L["kontingensi"] = PAGU - base - fee - ppn
    else:
        L["kontingensi"] = 9_330_000
        fee = fee_net
        ppn = round(fee * PPN)
    perj = sum(L.values())
    total = perj + fee + ppn
    return dict(n=n, nights=nights, trips=trips, lines=L, perj=perj, fee=fee, ppn=ppn, total=total,
                fee_gross=fee + ppn, perj_pct=perj / total * 100, fee_pct=(fee + ppn) / total * 100)


A = scenario(extras=2, fee_net=42_000_000, balance=True)   # usulan utama = pagu Rp510 jt
B = scenario(extras=0, fee_net=34_000_000, balance=False)  # alternatif 12 LPTK saja


def rp(v, dots=True):
    s = f"{int(round(v)):,}".replace(",", ".")
    return "Rp" + s


if __name__ == "__main__":
    for nm, sc in (("A 14 LPTK", A), ("B 12 LPTK", B)):
        print(nm, sc["n"], "LPTK", sc["nights"], "malam", sc["trips"], "perjalanan")
        for k, v in sc["lines"].items():
            print("  ", k, rp(v))
        print("  perjalanan", rp(sc["perj"]), f"{sc['perj_pct']:.2f}%", "fee", rp(sc["fee"]), "ppn", rp(sc["ppn"]),
              f"fee gross {sc['fee_pct']:.2f}%", "TOTAL", rp(sc["total"]))
