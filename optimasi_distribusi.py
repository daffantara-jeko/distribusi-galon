# =====================================================================
#  OPTIMASI DISTRIBUSI GALON - VRP DENGAN TIME WINDOW (TANPA CLUSTER)
#  Depo tunggal | Kapasitas | Time window per slot | Multi-trip | FCFS
# ---------------------------------------------------------------------
#  Cara pakai singkat:
#     hasil = jalankan("Input_Distribusi.xlsx")
#  Input  : file Excel berisi sheet Parameter, Lokasi, Order
#           (opsional: Matriks Jarak & Matriks Waktu)
#  Output : Hasil_Optimasi_Distribusi.xlsx (+ Peta_Rute.html, Gantt_Armada.png)
#
#  Jumlah outlet, jumlah order, jumlah armada, kapasitas, slot, dll.
#  TIDAK di-hardcode: semuanya dibaca dari file input.
# =====================================================================
import math, time, itertools
from datetime import time as dtime, datetime

import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import requests
except ImportError:
    requests = None
try:
    from ortools.constraint_solver import pywrapcp, routing_enums_pb2
except ImportError:
    pywrapcp = None

INF = float("inf")


# ---------------------------------------------------------------------
# 1. UTILITAS WAKTU
# ---------------------------------------------------------------------
def ke_menit(v):
    """Ubah jam (07:50 / time / angka Excel) menjadi menit sejak 00:00."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.hour * 60 + v.minute + v.second / 60
    if isinstance(v, dtime):
        return v.hour * 60 + v.minute + v.second / 60
    if isinstance(v, (int, float)):
        return float(v) * 1440 if v < 1 else float(v)  # pecahan hari Excel / menit
    s = str(v).strip().replace(".", ":")
    j, m = s.split(":")[:2]
    return int(j) * 60 + int(m)


def ke_jam(menit):
    """Menit -> teks HH:MM."""
    if menit is None:
        return "-"
    menit = int(round(menit))
    return f"{menit // 60:02d}:{menit % 60:02d}"


# ---------------------------------------------------------------------
# 2. MEMBACA INPUT
# ---------------------------------------------------------------------
PARAM_DEFAULT = {
    "jumlah armada": 4,
    "kapasitas armada": 100,
    "mulai pengiriman": "08:00",
    "interval slot": 120,
    "waktu muat": 5,               # detik / galon
    "waktu bongkar": 0.1,          # menit / galon
    "prioritas fcfs": "Ya",
    "konsumsi bbm": 8,             # km / liter
    "harga bbm": 8600,             # Rp / liter
    "hari operasional per bulan": 30,
    "sumber jarak": "OSRM",        # OSRM atau Haversine
    "faktor koreksi jarak": 1.3,   # dipakai jika Haversine
    "kecepatan rata-rata": 25,     # km/jam, dipakai jika Haversine
    "batas waktu solver": 2,       # detik per sub-masalah OR-Tools
}


def _norm(s):
    return str(s).strip().lower() if s is not None else ""


def baca_parameter(ws):
    p = dict(PARAM_DEFAULT)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        nama = _norm(row[0])
        for k in PARAM_DEFAULT:
            if nama.startswith(k):
                p[k] = row[1] if row[1] is not None else p[k]
    p["mulai pengiriman"] = ke_menit(p["mulai pengiriman"])
    for k in ["jumlah armada", "kapasitas armada"]:
        p[k] = int(p[k])
    for k in ["interval slot", "waktu muat", "waktu bongkar", "konsumsi bbm", "harga bbm",
              "hari operasional per bulan", "faktor koreksi jarak", "kecepatan rata-rata",
              "batas waktu solver"]:
        p[k] = float(p[k])
    p["prioritas fcfs"] = _norm(p["prioritas fcfs"]) in ("ya", "yes", "true", "1")
    return p


def _header_index(ws, kandidat):
    """Cari baris header & posisi kolom berdasarkan nama kolom."""
    for r in range(1, min(ws.max_row, 15) + 1):
        vals = [_norm(c.value) for c in ws[r]]
        pos = {}
        for key, names in kandidat.items():
            for i, v in enumerate(vals):
                if any(v.startswith(n) for n in names):
                    pos[key] = i
                    break
        if len(pos) == len(kandidat):
            return r, pos
    raise ValueError(f"Header sheet '{ws.title}' tidak ditemukan. Kolom wajib: {list(kandidat)}")


def baca_lokasi(ws):
    r0, pos = _header_index(ws, {"kode": ["kode"], "lat": ["lat"], "lon": ["lon", "lng"]})
    _, posn = r0, {}
    hdr = [_norm(c.value) for c in ws[r0]]
    for i, v in enumerate(hdr):
        if v.startswith("nama"):
            posn["nama"] = i
        if v.startswith("segmen"):
            posn["segmen"] = i
    lok = []
    for row in ws.iter_rows(min_row=r0 + 1, values_only=True):
        if row[pos["kode"]] in (None, ""):
            continue
        lok.append(dict(kode=str(row[pos["kode"]]).strip(),
                        nama=row[posn["nama"]] if "nama" in posn else "",
                        segmen=row[posn["segmen"]] if "segmen" in posn else "",
                        lat=float(row[pos["lat"]]), lon=float(row[pos["lon"]])))
    if lok[0]["kode"].upper() != "DEPO":
        raise ValueError("Baris pertama sheet Lokasi harus DEPO.")
    return lok


def baca_order(ws, kode_valid):
    r0, pos = _header_index(ws, {"kode": ["kode"], "qty": ["jumlah", "qty", "galon"], "jam": ["jam", "order time", "waktu"]})
    orders = []
    for row in ws.iter_rows(min_row=r0 + 1, values_only=True):
        if row[pos["kode"]] in (None, ""):
            continue
        k = str(row[pos["kode"]]).strip()
        if k not in kode_valid:
            raise ValueError(f"Kode order '{k}' tidak ada di sheet Lokasi.")
        orders.append(dict(kode=k, qty=int(row[pos["qty"]]), t=ke_menit(row[pos["jam"]])))
    return orders


def baca_matriks(wb, nama, kode):
    """Pakai matriks dari Excel jika ada & lengkap untuk semua kode lokasi."""
    if nama not in wb.sheetnames:
        return None
    ws = wb[nama]
    hdr = [str(c.value).strip() if c.value is not None else None for c in ws[1]][1:]
    baris = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            baris[str(row[0]).strip()] = row[1:]
    if not all(k in hdr and k in baris for k in kode):
        return None
    ci = {k: hdr.index(k) for k in kode}
    M = np.array([[float(baris[a][ci[b]]) for b in kode] for a in kode])
    return M


# ---------------------------------------------------------------------
# 3. MATRIKS JARAK & WAKTU (OSRM otomatis, cadangan Haversine)
# ---------------------------------------------------------------------
def haversine(a, b):
    R = 6371.0
    la1, lo1, la2, lo2 = map(math.radians, [a["lat"], a["lon"], b["lat"], b["lon"]])
    h = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    return 2 * R * math.asin(math.sqrt(h))


def matriks_haversine(lok, faktor, kecepatan):
    n = len(lok)
    D = np.array([[haversine(lok[i], lok[j]) * faktor for j in range(n)] for i in range(n)])
    T = D / kecepatan * 60
    return D, T


def matriks_osrm(lok, blok=40, server="https://router.project-osrm.org"):
    """Ambil jarak (km) & waktu (menit) jalan raya dari OSRM, dipecah per blok."""
    n = len(lok)
    D = np.zeros((n, n)); T = np.zeros((n, n))
    idx = list(range(n))
    for s0 in range(0, n, blok):
        for d0 in range(0, n, blok):
            src = idx[s0:s0 + blok]; dst = idx[d0:d0 + blok]
            semua = src + [j for j in dst if j not in src]
            coord = ";".join(f"{lok[i]['lon']},{lok[i]['lat']}" for i in semua)
            url = (f"{server}/table/v1/driving/{coord}?annotations=distance,duration"
                   f"&sources={';'.join(str(semua.index(i)) for i in src)}"
                   f"&destinations={';'.join(str(semua.index(j)) for j in dst)}")
            for coba in range(4):
                try:
                    r = requests.get(url, timeout=60).json()
                    if r.get("code") == "Ok":
                        break
                except Exception:
                    pass
                time.sleep(2 + coba * 3)
            else:
                raise RuntimeError("OSRM tidak merespons.")
            for a, i in enumerate(src):
                for b, j in enumerate(dst):
                    D[i, j] = (r["distances"][a][b] or 0) / 1000
                    T[i, j] = (r["durations"][a][b] or 0) / 60
            time.sleep(1)
    return np.round(D, 2), np.round(T, 2)


def siapkan_matriks(wb, lok, p, log):
    kode = [l["kode"] for l in lok]
    D = baca_matriks(wb, "Matriks Jarak", kode)
    T = baca_matriks(wb, "Matriks Waktu", kode)
    if D is not None and T is not None:
        log("Matriks jarak & waktu diambil dari file input.")
        return D, T, "Matriks dari file input"
    if _norm(p["sumber jarak"]).startswith("osrm") and requests is not None:
        try:
            log("Mengambil matriks jarak & waktu dari OSRM (jalan raya) ...")
            D, T = matriks_osrm(lok)
            return D, T, "OSRM (jarak jalan raya)"
        except Exception as e:
            log(f"  OSRM gagal ({e}). Beralih ke Haversine x faktor koreksi.")
    D, T = matriks_haversine(lok, p["faktor koreksi jarak"], p["kecepatan rata-rata"])
    return D, T, f"Haversine x {p['faktor koreksi jarak']} (kecepatan {p['kecepatan rata-rata']} km/jam)"


# ---------------------------------------------------------------------
# 4. MODEL: RUTE SATU TRIP (TSP) & PENGELOMPOKAN (CVRP)
# ---------------------------------------------------------------------
class Model:
    def __init__(self, D, T, orders, p):
        self.D, self.T, self.o, self.p = D, T, orders, p
        self.cap = p["kapasitas armada"]
        self._tsp = {}

    def jarak_rute(self, rute):
        """rute = list indeks order; Depo -> ... -> Depo."""
        loc = [0] + [self.o[i]["loc"] for i in rute] + [0]
        return sum(self.D[a, b] for a, b in zip(loc, loc[1:]))

    def tsp(self, grup):
        """Urutan kunjungan terpendek untuk satu trip."""
        key = tuple(sorted(grup))
        if key in self._tsp:
            return self._tsp[key]
        g = list(key)
        if len(g) <= 7:                                   # eksak (semua permutasi)
            best = min(itertools.permutations(g), key=self.jarak_rute)
            best = list(best)
        else:                                             # nearest neighbor + 2-opt
            sisa = set(g); best = []; cur = 0
            while sisa:
                nx = min(sisa, key=lambda i: self.D[cur, self.o[i]["loc"]])
                best.append(nx); sisa.remove(nx); cur = self.o[nx]["loc"]
            ok = True
            while ok:
                ok = False
                for i in range(len(best) - 1):
                    for j in range(i + 1, len(best)):
                        c = best[:i] + best[i:j + 1][::-1] + best[j + 1:]
                        if self.jarak_rute(c) < self.jarak_rute(best) - 1e-9:
                            best, ok = c, True
        res = (self.jarak_rute(best), best)
        self._tsp[key] = res
        return res

    def muatan(self, grup):
        return sum(self.o[i]["qty"] for i in grup)

    # ---- CVRP eksak untuk jumlah order kecil (partisi himpunan) ----
    def cvrp_eksak(self, items, maks_trip):
        n = len(items); full = (1 << n) - 1
        cost = {}
        for mask in range(1, full + 1):
            g = [items[i] for i in range(n) if mask >> i & 1]
            if self.muatan(g) <= self.cap:
                cost[mask] = self.tsp(g)[0]
        dp = {0: (0.0, [])}
        best = [{0: (0.0, [])}]
        for _ in range(maks_trip):
            prev = best[-1]; cur = dict(prev)
            for mask in range(1, full + 1):
                low = mask & -mask; sub = mask
                while sub:
                    if sub & low and sub in cost and (mask ^ sub) in prev:
                        c = cost[sub] + prev[mask ^ sub][0]
                        if c < cur.get(mask, (INF,))[0] - 1e-9:
                            cur[mask] = (c, prev[mask ^ sub][1] + [sub])
                    sub = (sub - 1) & mask
            best.append(cur)
        res = best[-1]
        out = {}
        for p in range(n + 1):
            m = (1 << p) - 1
            if m in res:
                out[p] = (res[m][0], [[items[i] for i in range(n) if s >> i & 1] for s in res[m][1]])
            else:
                out[p] = (INF, None)
        return out  # out[p] = solusi optimal untuk p order pertama

    # ---- CVRP heuristik OR-Tools untuk jumlah order besar ----
    def cvrp_ortools(self, items, maks_trip):
        if not items:
            return 0.0, []
        if pywrapcp is None:
            raise ImportError("Install OR-Tools: pip install ortools")
        nodes = [0] + [self.o[i]["loc"] for i in items]
        dem = [0] + [self.o[i]["qty"] for i in items]
        man = pywrapcp.RoutingIndexManager(len(nodes), maks_trip, 0)
        rt = pywrapcp.RoutingModel(man)
        cb = rt.RegisterTransitCallback(lambda a, b: int(round(self.D[nodes[man.IndexToNode(a)], nodes[man.IndexToNode(b)]] * 1000)))
        rt.SetArcCostEvaluatorOfAllVehicles(cb)
        dcb = rt.RegisterUnaryTransitCallback(lambda a: dem[man.IndexToNode(a)])
        rt.AddDimensionWithVehicleCapacity(dcb, 0, [self.cap] * maks_trip, True, "Muatan")
        sp = pywrapcp.DefaultRoutingSearchParameters()
        sp.first_solution_strategy = routing_enums_pb2.FirstSolutionStrategy.PARALLEL_CHEAPEST_INSERTION
        sp.local_search_metaheuristic = routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH
        sp.time_limit.FromMilliseconds(int(self.p["batas waktu solver"] * 1000))
        sol = rt.SolveWithParameters(sp)
        if sol is None:
            return INF, None
        grup = []
        for v in range(maks_trip):
            i = rt.Start(v); g = []
            while not rt.IsEnd(i):
                nd = man.IndexToNode(i)
                if nd:
                    g.append(items[nd - 1])
                i = sol.Value(rt.NextVar(i))
            if g:
                grup.append(g)
        return sum(self.tsp(g)[0] for g in grup), grup


# ---------------------------------------------------------------------
# 5. BATCHING PER SLOT
# ---------------------------------------------------------------------
def batching_slot(M, idx, m_bebas, fcfs):
    """
    idx      : order dalam slot, sudah urut jam order (FCFS)
    m_bebas  : armada yang sudah siap di depo saat jam slot
    Hasil    : list grup (trip) dalam urutan keberangkatan.

    Jika FCFS aktif:
      - Gelombang pertama: p order paling awal dibagi ke <= m_bebas armada
        (berangkat bersamaan pada jam slot, digabung bebas -> CVRP).
      - Sisanya diberangkatkan berurutan: tiap trip berisi order yang
        berurutan waktu masuknya (order lebih awal tidak pernah dimuat
        lebih lambat). Dipilih kombinasi dengan total jarak minimum
        (dynamic programming).
    Jika FCFS tidak aktif: CVRP bebas seluruh order slot.
    """
    n, cap = len(idx), M.cap
    if not fcfs:
        _, grup = M.cvrp_ortools(idx, n)
        return sorted(grup, key=lambda g: min(M.o[i]["t"] for i in g))

    # DP segmen berurutan untuk sisa order
    seg = [INF] * (n + 1); pil = [None] * (n + 1); seg[n] = 0.0
    for i in range(n - 1, -1, -1):
        q = 0
        for j in range(i + 1, n + 1):
            q += M.o[idx[j - 1]]["qty"]
            if q > cap:
                break
            c = M.tsp(idx[i:j])[0] + seg[j]
            if c < seg[i] - 1e-9:
                seg[i], pil[i] = c, j

    # gelombang pertama (CVRP untuk tiap kemungkinan p)
    pmax, q = 0, 0
    for i in idx:
        if q + M.o[i]["qty"] > m_bebas * cap:
            break
        q += M.o[i]["qty"]; pmax += 1
    if m_bebas > 0 and pmax <= 12:
        gel = M.cvrp_eksak(idx[:pmax], m_bebas)
    else:
        gel = {0: (0.0, [])}
        for p in range(1, pmax + 1):
            gel[p] = M.cvrp_ortools(idx[:p], m_bebas)

    best = (INF, None)
    for p, (c, g) in gel.items():
        if g is not None and c + seg[p] < best[0] - 1e-9:
            best = (c + seg[p], p)
    p = best[1]
    hasil = list(gel[p][1])
    i = p
    while i < n:
        j = pil[i]; hasil.append(idx[i:j]); i = j
    return hasil


# ---------------------------------------------------------------------
# 6. SIMULASI JADWAL (berlaku untuk VRP maupun FCFS)
# ---------------------------------------------------------------------
def jalankan_trip(M, rute, mulai_muat, p):
    muat = M.muatan(rute)
    berangkat = mulai_muat + muat * p["waktu muat"] / 60
    t, pos, tiba = berangkat, 0, []
    for i in rute:
        loc = M.o[i]["loc"]
        t += M.T[pos, loc]; tiba.append(t)
        t += M.o[i]["qty"] * p["waktu bongkar"]; pos = loc
    kembali = t + M.T[pos, 0]
    return dict(muatan=muat, berangkat=berangkat, tiba=tiba, kembali=kembali, km=M.jarak_rute(rute))


def jam_slot(t, p):
    s0, iv = p["mulai pengiriman"], p["interval slot"]
    return s0 if t <= s0 else s0 + math.ceil((t - s0) / iv - 1e-9) * iv


def simulasi_vrp(M, p, log):
    NV = p["jumlah armada"]
    bebas = [p["mulai pengiriman"]] * NV
    rit = [0] * NV
    km_armada = [0.0] * NV
    trips = []
    slots = sorted({o["slot"] for o in M.o})
    for s in slots:
        idx = sorted([i for i, o in enumerate(M.o) if o["slot"] == s], key=lambda i: (M.o[i]["t"], i))
        m_bebas = sum(1 for b in bebas if b <= s + 1e-9)
        grup = batching_slot(M, idx, m_bebas, p["prioritas fcfs"])
        for g in grup:
            rute = M.tsp(g)[1]
            v = min(range(NV), key=lambda v: (max(bebas[v], s), km_armada[v], v))  # seimbangkan beban
            mulai = max(bebas[v], s)
            tr = jalankan_trip(M, rute, mulai, p)
            rit[v] += 1
            tr.update(id=f"T{len(trips) + 1:02d}", slot=s, armada=f"V{v + 1}", rit=rit[v], rute=rute, mulai_muat=mulai)
            trips.append(tr); bebas[v] = tr["kembali"]; km_armada[v] += tr["km"]
        log(f"  Slot {ke_jam(s)}: {len(idx)} order -> {len(grup)} trip")
    return trips


def simulasi_fcfs(M, p):
    """Kondisi eksisting: 1 order = 1 trip, langsung dikirim, armada paling cepat kosong."""
    NV = p["jumlah armada"]
    bebas = [p["mulai pengiriman"]] * NV; rit = [0] * NV; trips = []
    for i in sorted(range(len(M.o)), key=lambda i: (M.o[i]["t"], i)):
        v = min(range(NV), key=lambda v: (bebas[v], v))
        mulai = max(bebas[v], M.o[i]["t"], p["mulai pengiriman"])
        tr = jalankan_trip(M, [i], mulai, p); rit[v] += 1
        tr.update(id=f"F{len(trips) + 1:02d}", slot=None, armada=f"V{v + 1}", rit=rit[v], rute=[i], mulai_muat=mulai)
        trips.append(tr); bebas[v] = tr["kembali"]
    return trips


# ---------------------------------------------------------------------
# 7. KPI & VALIDASI
# ---------------------------------------------------------------------
def detail_order(M, trips):
    det = {}
    for tr in trips:
        for u, (i, a) in enumerate(zip(tr["rute"], tr["tiba"]), 1):
            det[i] = dict(trip=tr["id"], armada=tr["armada"], rit=tr["rit"], urutan=u, mulai_muat=tr["mulai_muat"],
                          berangkat=tr["berangkat"], tiba=a, tunggu=a - M.o[i]["t"])
    return det


def cek_fcfs(M, det):
    """Order yang masuk lebih awal (slot sama) tidak boleh mulai dimuat lebih lambat."""
    hasil = {}
    for i, o in enumerate(M.o):
        ok = all(not (M.o[j]["slot"] == o["slot"] and M.o[j]["t"] > o["t"] and det[j]["mulai_muat"] < det[i]["mulai_muat"] - 1e-6)
                 for j in range(len(M.o)))
        hasil[i] = "OK" if ok else "LANGGAR"
    return hasil


def hitung_kpi(M, trips, p):
    det = detail_order(M, trips)
    total_q = sum(o["qty"] for o in M.o)
    km = sum(t["km"] for t in trips)
    bbm = km / p["konsumsi bbm"] * p["harga bbm"]
    return {
        "Jumlah order": len(M.o),
        "Total galon": total_q,
        "Jumlah trip": len(trips),
        "Total jarak (km)": round(km, 2),
        "Load factor": total_q / (len(trips) * M.cap),
        "Rata-rata muatan per trip (galon)": round(total_q / len(trips), 1),
        "Rata-rata waktu tunggu (menit)": round(float(np.mean([d["tunggu"] for d in det.values()])), 2),
        "Waktu tunggu terlama (menit)": round(max(d["tunggu"] for d in det.values()), 1),
        "Customer terakhir tiba": ke_jam(max(d["tiba"] for d in det.values())),
        "Armada terakhir kembali": ke_jam(max(t["kembali"] for t in trips)),
        "Total waktu armada di luar depo (jam)": round(sum(t["kembali"] - t["berangkat"] for t in trips) / 60, 2),
        "Biaya BBM per hari (Rp)": round(bbm),
        "Biaya BBM per bulan (Rp)": round(bbm * p["hari operasional per bulan"]),
        "Muatan maksimum per trip (galon)": max(t["muatan"] for t in trips),
    }


# ---------------------------------------------------------------------
# 8. OUTPUT EXCEL, PETA & GANTT
# ---------------------------------------------------------------------
NAVY = "1F4E78"
HDR = dict(font=Font(bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor=NAVY),
           alignment=Alignment(horizontal="center", vertical="center", wrap_text=True))
TIPIS = Side(style="thin", color="A6A6A6")
BORDER = Border(left=TIPIS, right=TIPIS, top=TIPIS, bottom=TIPIS)


def tulis_tabel(ws, r0, judul, header, rows, lebar=None, fmt=None):
    ws.cell(r0, 1, judul).font = Font(bold=True, size=13, color=NAVY)
    for c, h in enumerate(header, 1):
        cell = ws.cell(r0 + 1, c, h)
        for k, v in HDR.items():
            setattr(cell, k, v)
        cell.border = BORDER
    for r, row in enumerate(rows, r0 + 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v); cell.border = BORDER
            cell.alignment = Alignment(horizontal="center" if not isinstance(v, str) or len(v) < 12 else "left", vertical="center")
            if fmt and header[c - 1] in fmt:
                cell.number_format = fmt[header[c - 1]]
            if v == "LANGGAR":
                cell.font = Font(bold=True, color="C00000")
    if lebar:
        for c, w in enumerate(lebar, 1):
            ws.column_dimensions[get_column_letter(c)].width = w
    return r0 + len(rows) + 3


def simpan_excel(path, M, lok, p, vrp, fcfs, kpi_v, kpi_f, sumber, D, T):
    wb = openpyxl.Workbook()
    nama = {l["kode"]: l["nama"] for l in lok}
    det_v, det_f = detail_order(M, vrp), detail_order(M, fcfs)
    cek = cek_fcfs(M, det_v)

    # --- Ringkasan
    ws = wb.active; ws.title = "Ringkasan"
    rows = []
    for k in kpi_v:
        a, b = kpi_f[k], kpi_v[k]
        d = (b - a) / a if isinstance(a, (int, float)) and a else "-"
        rows.append([k, a, b, d])
    r = tulis_tabel(ws, 1, "PERBANDINGAN KINERJA: EKSISTING (FCFS) vs VRP TIME WINDOW", ["Indikator", "Eksisting (FCFS)", "VRP Time Window", "Perubahan"],
                    rows, [42, 20, 20, 14])
    for rr in range(3, 3 + len(rows)):
        if ws.cell(rr, 1).value == "Load factor":
            ws.cell(rr, 2).number_format = ws.cell(rr, 3).number_format = "0.0%"
        if isinstance(ws.cell(rr, 4).value, float):
            ws.cell(rr, 4).number_format = "+0.0%;-0.0%;0%"
        if "Rp" in ws.cell(rr, 1).value:
            ws.cell(rr, 2).number_format = ws.cell(rr, 3).number_format = "#,##0"
    validasi = [
        ["Semua order terkirim", "OK" if len(det_v) == len(M.o) else "LANGGAR"],
        ["Muatan per trip <= kapasitas", "OK" if kpi_v["Muatan maksimum per trip (galon)"] <= M.cap else "LANGGAR"],
        ["Tidak ada pengiriman sebelum slot (time window)", "OK" if all(det_v[i]["mulai_muat"] >= o["slot"] - 1e-6 for i, o in enumerate(M.o)) else "LANGGAR"],
        ["Urutan FCFS dalam slot", f"{sum(v == 'OK' for v in cek.values())}/{len(cek)} OK"],
        ["Selesai di hari yang sama (kembali < 24:00)", "OK" if max(t["kembali"] for t in vrp) < 1440 else "LANGGAR"],
    ]
    r = tulis_tabel(ws, r, "VALIDASI CONSTRAINT", ["Pemeriksaan", "Hasil"], validasi)
    prm = [[k.title(), ke_jam(v) if k == "mulai pengiriman" else ("Ya" if v is True else "Tidak" if v is False else v)] for k, v in p.items()]
    prm.append(["Sumber matriks jarak/waktu", sumber])
    tulis_tabel(ws, r, "PARAMETER YANG DIPAKAI", ["Parameter", "Nilai"], prm)

    # --- Jadwal per order
    ws = wb.create_sheet("Jadwal per Order")
    rows = []
    for i in sorted(range(len(M.o)), key=lambda i: (M.o[i]["slot"], det_v[i]["trip"], det_v[i]["urutan"])):
        o, d = M.o[i], det_v[i]
        rows.append([o["id"], nama[o["kode"]], o["qty"], ke_jam(o["t"]), ke_jam(o["slot"]), d["trip"], d["armada"], d["rit"], d["urutan"],
                     ke_jam(d["mulai_muat"]), ke_jam(d["berangkat"]), ke_jam(d["tiba"]), round(d["tunggu"], 1), cek[i]])
    tulis_tabel(ws, 1, "JADWAL PENGIRIMAN PER ORDER (VRP TIME WINDOW)",
                ["Kode", "Nama Outlet", "Galon", "Jam Order", "Slot Kirim", "Trip", "Armada", "Rit ke-", "Urutan", "Mulai Muat", "Berangkat", "Tiba", "Waktu Tunggu (menit)", "Cek FCFS"],
                rows, [9, 32, 8, 10, 10, 8, 9, 8, 8, 11, 11, 9, 13, 10])
    ws.freeze_panes = "A3"

    # --- Rekap trip
    def rekap_trip(trips):
        return [[t["id"], ke_jam(t["slot"]) if t["slot"] is not None else "-", t["armada"], t["rit"],
                 "DEPO → " + " → ".join(M.o[i]["id"] for i in t["rute"]) + " → DEPO", len(t["rute"]), t["muatan"], t["muatan"] / M.cap,
                 ke_jam(t["mulai_muat"]), ke_jam(t["berangkat"]), ke_jam(t["kembali"]), round(t["km"], 2)] for t in trips]
    hdr_trip = ["Trip", "Slot", "Armada", "Rit ke-", "Rute", "Jml Outlet", "Muatan (galon)", "Load Factor", "Mulai Muat", "Berangkat", "Kembali", "Jarak (km)"]
    ws = wb.create_sheet("Rekap Trip")
    rows = rekap_trip(vrp)
    rows.append(["TOTAL", "", "", "", "", sum(len(t["rute"]) for t in vrp), sum(t["muatan"] for t in vrp), kpi_v["Load factor"], "", "", "", round(kpi_v["Total jarak (km)"], 2)])
    tulis_tabel(ws, 1, "REKAP PER TRIP (VRP TIME WINDOW)", hdr_trip, rows, [8, 8, 9, 8, 46, 10, 12, 11, 11, 11, 10, 11], fmt={"Load Factor": "0%"})
    ws.freeze_panes = "A3"

    # --- Rekap armada & slot
    ws = wb.create_sheet("Rekap Armada & Slot")
    rows = []
    for v in range(p["jumlah armada"]):
        tt = [t for t in vrp if t["armada"] == f"V{v + 1}"]
        rows.append([f"V{v + 1}", len(tt), sum(t["muatan"] for t in tt), round(sum(t["km"] for t in tt), 2),
                     ke_jam(max(t["kembali"] for t in tt)) if tt else "-"])
    rows.append(["Total", len(vrp), sum(t["muatan"] for t in vrp), round(kpi_v["Total jarak (km)"], 2), kpi_v["Armada terakhir kembali"]])
    r = tulis_tabel(ws, 1, "REKAP PER ARMADA", ["Armada", "Jumlah Trip", "Total Galon", "Jarak (km)", "Selesai (kembali depo)"], rows, [12, 14, 14, 14, 20, 40])
    rows = []
    for s in sorted({o["slot"] for o in M.o}):
        oo = [o for o in M.o if o["slot"] == s]; tt = [t for t in vrp if t["slot"] == s]
        rows.append([ke_jam(s), len(oo), ", ".join(o["id"] for o in oo), sum(o["qty"] for o in oo), len(tt), round(sum(t["km"] for t in tt), 2)])
    rows.append(["Total", len(M.o), "", sum(o["qty"] for o in M.o), len(vrp), round(kpi_v["Total jarak (km)"], 2)])
    tulis_tabel(ws, r, "REKAP PER SLOT TIME WINDOW", ["Slot", "Jumlah Order", "Kode Customer", "Total Galon", "Jumlah Trip", "Jarak (km)"], rows)
    ws.column_dimensions["C"].width = 50

    # --- FCFS eksisting
    ws = wb.create_sheet("FCFS Eksisting")
    rows = rekap_trip(fcfs)
    rows.append(["TOTAL", "", "", "", "", len(M.o), sum(t["muatan"] for t in fcfs), kpi_f["Load factor"], "", "", "", round(kpi_f["Total jarak (km)"], 2)])
    tulis_tabel(ws, 1, "KONDISI EKSISTING (FCFS: 1 ORDER = 1 TRIP)", hdr_trip, rows, [8, 8, 9, 8, 30, 10, 12, 11, 11, 11, 10, 11], fmt={"Load Factor": "0%"})

    # --- Matriks
    kode = [l["kode"] for l in lok]
    for nm, X in [("Matriks Jarak", D), ("Matriks Waktu", T)]:
        ws = wb.create_sheet(nm)
        ws.cell(1, 1, "Dari / Ke").font = Font(bold=True)
        for j, k in enumerate(kode, 2):
            ws.cell(1, j, k).font = Font(bold=True); ws.cell(j, 1, k).font = Font(bold=True)
        for a in range(len(kode)):
            for b in range(len(kode)):
                ws.cell(a + 2, b + 2, round(float(X[a, b]), 2))
    wb.save(path)


def simpan_peta(path, M, lok, vrp):
    try:
        import folium
    except ImportError:
        return False
    warna = ["red", "blue", "green", "purple", "orange", "darkred", "cadetblue", "darkgreen", "pink", "black"]
    dep = lok[0]
    fm = folium.Map(location=[dep["lat"], dep["lon"]], zoom_start=13)
    folium.Marker([dep["lat"], dep["lon"]], tooltip="DEPO", icon=folium.Icon(color="black", icon="home")).add_to(fm)
    for l in lok[1:]:
        folium.CircleMarker([l["lat"], l["lon"]], radius=5, tooltip=f"{l['kode']} - {l['nama']}", fill=True).add_to(fm)
    for k, t in enumerate(vrp):
        pts = [(dep["lat"], dep["lon"])] + [(lok[M.o[i]["loc"]]["lat"], lok[M.o[i]["loc"]]["lon"]) for i in t["rute"]] + [(dep["lat"], dep["lon"])]
        fg = folium.FeatureGroup(name=f"{t['id']} | {t['armada']} | slot {ke_jam(t['slot'])}")
        folium.PolyLine(pts, color=warna[int(t["armada"][1:]) % len(warna)], weight=3, opacity=0.8,
                        tooltip=f"{t['id']} {t['armada']}: " + " → ".join(M.o[i]["id"] for i in t["rute"])).add_to(fg)
        fg.add_to(fm)
    folium.LayerControl(collapsed=False).add_to(fm)
    fm.save(path)
    return True


def simpan_gantt(path, M, vrp, p):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    NV = p["jumlah armada"]
    fig, ax = plt.subplots(figsize=(14, 1 + 0.7 * NV))
    for t in vrp:
        v = int(t["armada"][1:])
        ax.broken_barh([(t["mulai_muat"] / 60, (t["berangkat"] - t["mulai_muat"]) / 60)], (v - 0.35, 0.7), color="#9DC3E6")
        ax.broken_barh([(t["berangkat"] / 60, (t["kembali"] - t["berangkat"]) / 60)], (v - 0.35, 0.7), color="#2E75B6", edgecolor="white")
        ax.text((t["berangkat"] + t["kembali"]) / 120, v, t["id"], ha="center", va="center", color="white", fontsize=8)
    ax.set_yticks(range(1, NV + 1)); ax.set_yticklabels([f"V{v}" for v in range(1, NV + 1)])
    ax.invert_yaxis(); ax.set_xlabel("Jam"); ax.set_title("Jadwal Armada (biru muda = muat, biru = perjalanan)")
    for s in sorted({o["slot"] for o in M.o}):
        ax.axvline(s / 60, color="grey", ls=":", lw=0.8)
    ax.grid(axis="x", alpha=0.3); fig.tight_layout(); fig.savefig(path, dpi=150); plt.close(fig)


# ---------------------------------------------------------------------
# 9. PROGRAM UTAMA
# ---------------------------------------------------------------------
def jalankan(file_input, file_output="Hasil_Optimasi_Distribusi.xlsx", peta="Peta_Rute.html", gantt="Gantt_Armada.png", tampilkan=True, override=None):
    log = print if tampilkan else (lambda *a: None)
    t0 = time.time()
    wb = openpyxl.load_workbook(file_input, data_only=True)
    p = baca_parameter(wb["Parameter"])
    if override:
        p.update(override)
    lok = baca_lokasi(wb["Lokasi"])
    kode2loc = {l["kode"]: i for i, l in enumerate(lok)}
    orders_raw = baca_order(wb["Order"], kode2loc)
    log(f"Data terbaca: {len(lok) - 1} outlet, {len(orders_raw)} order, {sum(o['qty'] for o in orders_raw)} galon, "
        f"{p['jumlah armada']} armada x {p['kapasitas armada']} galon")

    # order > kapasitas dipecah menjadi beberapa bagian
    orders = []
    for o in orders_raw:
        sisa, k = o["qty"], 1
        bagian = math.ceil(sisa / p["kapasitas armada"])
        while sisa > 0:
            q = min(sisa, p["kapasitas armada"])
            orders.append(dict(id=o["kode"] if bagian == 1 else f"{o['kode']}#{k}", kode=o["kode"], loc=kode2loc[o["kode"]],
                               qty=q, t=o["t"], slot=jam_slot(o["t"], p)))
            sisa -= q; k += 1

    D, T, sumber = siapkan_matriks(wb, lok, p, log)
    M = Model(D, T, orders, p)

    log("Menyusun batching, rute & jadwal ...")
    vrp = simulasi_vrp(M, p, log)
    fcfs = simulasi_fcfs(M, p)
    kpi_v, kpi_f = hitung_kpi(M, vrp, p), hitung_kpi(M, fcfs, p)

    simpan_excel(file_output, M, lok, p, vrp, fcfs, kpi_v, kpi_f, sumber, D, T)
    ada_peta = simpan_peta(peta, M, lok, vrp) if peta else False
    if gantt:
        simpan_gantt(gantt, M, vrp, p)

    log("\n================ HASIL ================")
    log(f"{'Indikator':40s}{'FCFS':>14s}{'VRP TW':>14s}")
    for k in ["Jumlah trip", "Total jarak (km)", "Load factor", "Rata-rata waktu tunggu (menit)", "Armada terakhir kembali", "Biaya BBM per hari (Rp)"]:
        a, b = kpi_f[k], kpi_v[k]
        if k == "Load factor":
            a, b = f"{a:.0%}", f"{b:.0%}"
        log(f"{k:40s}{str(a):>14s}{str(b):>14s}")
    cek = cek_fcfs(M, detail_order(M, vrp))
    log(f"Cek FCFS: {sum(v == 'OK' for v in cek.values())}/{len(cek)} OK")
    log(f"\nFile tersimpan: {file_output}" + (f", {peta}" if ada_peta else "") + (f", {gantt}" if gantt else ""))
    log(f"Waktu proses: {time.time() - t0:.1f} detik")
    return dict(vrp=vrp, fcfs=fcfs, kpi_vrp=kpi_v, kpi_fcfs=kpi_f, model=M, lokasi=lok, param=p, sumber=sumber,
                cek_fcfs=cek, detail=detail_order(M, vrp))


if __name__ == "__main__":
    import sys
    jalankan(sys.argv[1] if len(sys.argv) > 1 else "Input_Distribusi.xlsx")
