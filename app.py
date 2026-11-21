# =====================================================================
#  DSS DISTRIBUSI GALON - Aplikasi Web (Streamlit)
#  Jalankan lokal :  streamlit run app.py
#  Mesin optimasi :  optimasi_distribusi.py (file yang sama dengan Colab)
# =====================================================================
import io, os, tempfile, contextlib
from datetime import time as dtime

import openpyxl
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import optimasi_distribusi as od

st.set_page_config(page_title="DSS Distribusi Galon", page_icon="🚚", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1.6rem;}
div[data-testid="stMetric"] {background:#EEF5FD; border:1px solid #CFE0F5; border-radius:10px; padding:10px 14px;}
div[data-testid="stMetricLabel"] p {font-size:0.85rem;}
div[data-testid="stMetricValue"] {font-size:1.55rem;}
</style>""", unsafe_allow_html=True)

DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(DIR, "Input_Distribusi.xlsx")


# ---------------------------------------------------------------------
# Menjalankan optimasi (hasil di-cache: file & parameter sama -> tidak dihitung ulang)
# ---------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def proses(file_bytes, override_items):
    tmp = tempfile.mkdtemp()
    f_in = os.path.join(tmp, "input.xlsx")
    open(f_in, "wb").write(file_bytes)
    out = {k: os.path.join(tmp, v) for k, v in
           dict(xlsx="Hasil_Optimasi_Distribusi.xlsx", peta="Peta_Rute.html", gantt="Gantt_Armada.png").items()}
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        r = od.jalankan(f_in, out["xlsx"], out["peta"], out["gantt"], tampilkan=True, override=dict(override_items))
    M, lok = r["model"], r["lokasi"]
    nama = {l["kode"]: l["nama"] for l in lok}
    det, cek = r["detail"], r["cek_fcfs"]

    jadwal = pd.DataFrame([{
        "Kode": o["id"], "Nama Outlet": nama[o["kode"]], "Galon": o["qty"], "Jam Order": od.ke_jam(o["t"]),
        "Slot Kirim": od.ke_jam(o["slot"]), "Trip": det[i]["trip"], "Armada": det[i]["armada"], "Rit ke-": det[i]["rit"],
        "Urutan": det[i]["urutan"], "Mulai Muat": od.ke_jam(det[i]["mulai_muat"]), "Berangkat": od.ke_jam(det[i]["berangkat"]),
        "Tiba": od.ke_jam(det[i]["tiba"]), "Waktu Tunggu (menit)": round(det[i]["tunggu"], 1), "Cek FCFS": cek[i]}
        for i, o in enumerate(M.o)]).sort_values(["Slot Kirim", "Trip", "Urutan"]).reset_index(drop=True)

    def tabel_trip(trips):
        return pd.DataFrame([{
            "Trip": t["id"], "Slot": od.ke_jam(t["slot"]) if t["slot"] is not None else "-", "Armada": t["armada"], "Rit ke-": t["rit"],
            "Rute": "DEPO → " + " → ".join(M.o[i]["id"] for i in t["rute"]) + " → DEPO", "Muatan (galon)": t["muatan"],
            "Load Factor": f"{t['muatan'] / M.cap:.0%}", "Mulai Muat": od.ke_jam(t["mulai_muat"]), "Berangkat": od.ke_jam(t["berangkat"]),
            "Kembali": od.ke_jam(t["kembali"]), "Jarak (km)": round(t["km"], 2)} for t in trips])

    trip, fcfs = tabel_trip(r["vrp"]), tabel_trip(r["fcfs"])
    armada = trip.groupby("Armada").agg(**{"Jumlah Trip": ("Trip", "count"), "Total Galon": ("Muatan (galon)", "sum"),
                                           "Jarak (km)": ("Jarak (km)", "sum"), "Selesai": ("Kembali", "max")}).reset_index()
    armada["Jarak (km)"] = armada["Jarak (km)"].round(2)
    slot = jadwal.groupby("Slot Kirim").agg(**{"Jumlah Order": ("Kode", "count"), "Kode Customer": ("Kode", ", ".join),
                                               "Total Galon": ("Galon", "sum")}).reset_index()
    slot["Jumlah Trip"] = slot["Slot Kirim"].map(trip.groupby("Slot")["Trip"].count())
    slot["Jarak (km)"] = slot["Slot Kirim"].map(trip.groupby("Slot")["Jarak (km)"].sum().round(2))

    files = {k: open(v, "rb").read() for k, v in out.items() if os.path.exists(v)}
    return dict(kpi_v=r["kpi_vrp"], kpi_f=r["kpi_fcfs"], jadwal=jadwal, trip=trip, fcfs=fcfs, armada=armada, slot=slot,
                files=files, log=buf.getvalue(), sumber=r["sumber"], p=r["param"],
                lok=pd.DataFrame(lok), n_out=len(lok) - 1)


def baca_param_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    return od.baca_parameter(wb["Parameter"])


# ---------------------------------------------------------------------
# SIDEBAR: input
# ---------------------------------------------------------------------
with st.sidebar:
    st.header("📥 Input Data")
    if os.path.exists(TEMPLATE):
        st.download_button("⬇️ Download template Excel", open(TEMPLATE, "rb").read(), "Input_Distribusi.xlsx",
                           use_container_width=True)
    up = st.file_uploader("Upload file input (.xlsx)", type=["xlsx"])
    pakai_contoh = st.checkbox("Pakai data contoh (template)", value=up is None)
    file_bytes = up.getvalue() if up is not None else (open(TEMPLATE, "rb").read() if pakai_contoh and os.path.exists(TEMPLATE) else None)

    override = {}
    if file_bytes:
        try:
            p0 = baca_param_excel(file_bytes)
        except Exception as e:
            st.error(f"Sheet Parameter tidak terbaca: {e}")
            st.stop()
        st.header("⚙️ Parameter")
        st.caption("Nilai awal dibaca dari sheet Parameter. Ubah di sini untuk mencoba skenario lain.")
        override["jumlah armada"] = st.number_input("Jumlah armada", 1, 50, int(p0["jumlah armada"]))
        override["kapasitas armada"] = st.number_input("Kapasitas armada (galon/rit)", 1, 1000, int(p0["kapasitas armada"]))
        slot_pilih = {60: "1 jam", 120: "2 jam", 180: "3 jam", 240: "4 jam"}
        iv = int(p0["interval slot"])
        override["interval slot"] = st.selectbox("Lebar time window", list(slot_pilih), index=list(slot_pilih).index(iv) if iv in slot_pilih else 1,
                                                 format_func=lambda v: slot_pilih[v])
        mp = st.time_input("Mulai pengiriman", dtime(int(p0["mulai pengiriman"] // 60), int(p0["mulai pengiriman"] % 60)))
        override["mulai pengiriman"] = mp.hour * 60 + mp.minute
        override["prioritas fcfs"] = st.toggle("Prioritas FCFS dalam slot", value=bool(p0["prioritas fcfs"]))
        with st.expander("Parameter lain"):
            override["waktu muat"] = st.number_input("Waktu muat (detik/galon)", 0.0, 120.0, float(p0["waktu muat"]))
            override["waktu bongkar"] = st.number_input("Waktu bongkar (menit/galon)", 0.0, 10.0, float(p0["waktu bongkar"]), step=0.05)
            override["konsumsi bbm"] = st.number_input("Konsumsi BBM (km/liter)", 1.0, 50.0, float(p0["konsumsi bbm"]))
            override["harga bbm"] = st.number_input("Harga BBM (Rp/liter)", 0.0, 50000.0, float(p0["harga bbm"]), step=100.0)
        jalan = st.button("🚀 Jalankan Optimasi", type="primary", use_container_width=True)

# ---------------------------------------------------------------------
# HALAMAN UTAMA
# ---------------------------------------------------------------------
st.title("🚚 Decision Support System – Distribusi Galon")
st.caption("VRP dengan Time Window, Batching, Multi-trip & Prioritas FCFS · Depo tunggal · dibandingkan dengan kondisi eksisting (FCFS)")

if not file_bytes:
    st.info("Upload file input Excel di panel kiri (atau centang **Pakai data contoh**), lalu klik **Jalankan Optimasi**.")
    st.markdown("""
**Format file input** (download template di panel kiri):
- **Parameter** – jumlah armada, kapasitas, jam mulai, lebar slot, waktu muat/bongkar, FCFS, BBM
- **Lokasi** – baris pertama DEPO, lalu semua outlet (kode, nama, latitude, longitude)
- **Order** – satu baris per order (kode, jumlah galon, jam order)
- **Matriks Jarak / Waktu** *(opsional)* – jika tidak ada, jarak jalan raya diambil otomatis dari OSRM
""")
    st.stop()

kunci = (file_bytes, tuple(sorted(override.items())))
if jalan:
    st.session_state["kunci"] = kunci
if st.session_state.get("kunci") != kunci:
    st.info("Klik **🚀 Jalankan Optimasi** di panel kiri untuk menghitung.")
    st.stop()

with st.spinner("Menghitung batching, rute & jadwal ... (data baru tanpa matriks butuh waktu mengambil jarak OSRM)"):
    try:
        R = proses(*kunci)
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
        st.stop()

kv, kf = R["kpi_v"], R["kpi_f"]

# ---------- KPI
st.subheader("📊 Ringkasan Kinerja")
c = st.columns(6)
def pct(a, b): return f"{(b - a) / a:+.1%}" if a else None
c[0].metric("Jumlah trip", kv["Jumlah trip"], pct(kf["Jumlah trip"], kv["Jumlah trip"]), delta_color="inverse")
c[1].metric("Total jarak", f"{kv['Total jarak (km)']:.1f} km", pct(kf["Total jarak (km)"], kv["Total jarak (km)"]), delta_color="inverse")
c[2].metric("Load factor", f"{kv['Load factor']:.0%}", f"{(kv['Load factor'] - kf['Load factor']) * 100:+.0f} poin")
c[3].metric("Biaya BBM/hari", f"Rp {kv['Biaya BBM per hari (Rp)'] / 1000:,.0f} rb".replace(",", "."),
            pct(kf["Biaya BBM per hari (Rp)"], kv["Biaya BBM per hari (Rp)"]), delta_color="inverse")
c[4].metric("Rata-rata tunggu", f"{kv['Rata-rata waktu tunggu (menit)']:.0f} mnt",
            f"{kv['Rata-rata waktu tunggu (menit)'] - kf['Rata-rata waktu tunggu (menit)']:+.0f} menit", delta_color="inverse")
c[5].metric("Terakhir kembali", kv["Armada terakhir kembali"])
st.caption(f"Perubahan dibandingkan kondisi eksisting FCFS (1 order = 1 trip). Sumber jarak: {R['sumber']}.")

ok_all = (R["jadwal"]["Cek FCFS"] == "OK").all()
v1, v2, v3, v4 = st.columns(4)
v1.success(f"✔ Semua {len(R['jadwal'])} order terjadwal")
if kv["Muatan maksimum per trip (galon)"] <= R["p"]["kapasitas armada"]:
    v2.success("✔ Muatan ≤ kapasitas")
else:
    v2.error("✘ Muatan melebihi kapasitas")
n_ok = (R["jadwal"]["Cek FCFS"] == "OK").sum()
if ok_all:
    v3.success(f"✔ FCFS {n_ok}/{len(R['jadwal'])}")
else:
    v3.warning(f"FCFS {n_ok}/{len(R['jadwal'])} (prioritas FCFS nonaktif)")
if kv["Armada terakhir kembali"] < "24:00":
    v4.success("✔ Selesai di hari yang sama")
else:
    v4.error("✘ Melewati 24:00 – tambah armada")

# ---------- TAB
tab = st.tabs(["🔎 Cari Order", "🚚 Jadwal per Armada", "📋 Jadwal per Order", "🧾 Rekap Trip", "📈 Rekap & Perbandingan",
               "🗓️ Gantt", "🗺️ Peta Rute", "⬇️ Download"])

with tab[0]:
    st.markdown("Pilih kode outlet untuk melihat **dikirim kapan, oleh armada mana, trip ke berapa, dan urutan ke berapa**.")
    kode = st.selectbox("Kode outlet", R["jadwal"]["Kode"].tolist(),
                        format_func=lambda k: f"{k} – {R['jadwal'].set_index('Kode').loc[k, 'Nama Outlet']}")
    row = R["jadwal"].set_index("Kode").loc[kode]
    a, b, c_, d = st.columns(4)
    a.metric(f"Slot kirim (order masuk {row['Jam Order']})", row["Slot Kirim"])
    b.metric("Armada", f"{row['Armada']} · rit ke-{row['Rit ke-']}")
    c_.metric("Trip · urutan kunjungan", f"{row['Trip']} · ke-{row['Urutan']}")
    d.metric(f"Tiba (tunggu {row['Waktu Tunggu (menit)']:.0f} menit)", row["Tiba"])
    t = R["trip"].set_index("Trip").loc[row["Trip"]]
    st.info(f"**Rute {row['Trip']} ({t['Armada']})**: {t['Rute']}  ·  muatan {t['Muatan (galon)']} galon ({t['Load Factor']})  ·  "
            f"berangkat {t['Berangkat']}, kembali {t['Kembali']}  ·  {t['Jarak (km)']} km")

with tab[1]:
    arm = st.radio("Pilih armada", sorted(R["trip"]["Armada"].unique(), key=lambda s: int(s[1:])), horizontal=True)
    tt = R["trip"][R["trip"]["Armada"] == arm]
    st.markdown(f"**{arm}**: {len(tt)} trip · {tt['Muatan (galon)'].sum()} galon · {tt['Jarak (km)'].sum():.2f} km")
    for _, t in tt.iterrows():
        with st.expander(f"Rit {t['Rit ke-']} · {t['Trip']} · slot {t['Slot']} · berangkat {t['Berangkat']} → kembali {t['Kembali']}", expanded=True):
            st.write(t["Rute"])
            st.dataframe(R["jadwal"][R["jadwal"]["Trip"] == t["Trip"]][["Urutan", "Kode", "Nama Outlet", "Galon", "Jam Order", "Tiba"]],
                         hide_index=True, use_container_width=True)

with tab[2]:
    f1, f2 = st.columns(2)
    fs = f1.multiselect("Filter slot", sorted(R["jadwal"]["Slot Kirim"].unique()))
    fa = f2.multiselect("Filter armada", sorted(R["jadwal"]["Armada"].unique()))
    df = R["jadwal"]
    if fs: df = df[df["Slot Kirim"].isin(fs)]
    if fa: df = df[df["Armada"].isin(fa)]
    st.dataframe(df, hide_index=True, use_container_width=True, height=560)

with tab[3]:
    st.dataframe(R["trip"], hide_index=True, use_container_width=True, height=560)

with tab[4]:
    k1, k2 = st.columns([3, 2])
    with k1:
        st.markdown("**Eksisting (FCFS) vs VRP Time Window**")
        rows = []
        for k in kv:
            a, b = kf[k], kv[k]
            rows.append({"Indikator": k, "Eksisting (FCFS)": f"{a:.0%}" if k == "Load factor" else a,
                         "VRP Time Window": f"{b:.0%}" if k == "Load factor" else b,
                         "Perubahan": pct(a, b) if isinstance(a, (int, float)) else "-"})
        st.dataframe(pd.DataFrame(rows).astype(str), hide_index=True, use_container_width=True, height=560)
    with k2:
        st.markdown("**Per armada**"); st.dataframe(R["armada"], hide_index=True, use_container_width=True)
        st.markdown("**Per slot**"); st.dataframe(R["slot"], hide_index=True, use_container_width=True)
    with st.expander("Kondisi eksisting FCFS (1 order = 1 trip)"):
        st.dataframe(R["fcfs"], hide_index=True, use_container_width=True)

with tab[5]:
    if "gantt" in R["files"]:
        st.image(R["files"]["gantt"], use_container_width=True)

with tab[6]:
    if "peta" in R["files"]:
        st.caption("Garis = urutan kunjungan tiap trip (garis lurus antar titik). Centang/hapus centang trip di kanan atas peta.")
        components.html(R["files"]["peta"].decode("utf-8"), height=620)
    else:
        st.warning("Library folium belum terpasang, peta tidak dibuat.")

with tab[7]:
    st.download_button("⬇️ Hasil Optimasi (Excel)", R["files"]["xlsx"], "Hasil_Optimasi_Distribusi.xlsx", use_container_width=True)
    if "gantt" in R["files"]:
        st.download_button("⬇️ Gantt Armada (PNG)", R["files"]["gantt"], "Gantt_Armada.png", use_container_width=True)
    if "peta" in R["files"]:
        st.download_button("⬇️ Peta Rute (HTML)", R["files"]["peta"], "Peta_Rute.html", use_container_width=True)
    with st.expander("Log proses"):
        st.code(R["log"])
