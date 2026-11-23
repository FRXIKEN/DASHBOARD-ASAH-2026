import html
import io
import os
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import data_loader as dl

st.set_page_config(page_title="by.ficlo Dashboard", page_icon="🌸", layout="wide")

HERE = Path(__file__).resolve().parent
DATA_PATH = os.environ.get("FICLO_DATA", str(HERE / "Dataset_Dasboard_DS_byficlo.xlsx"))

# ------------------------------------------------------------------ tema
GREEN, GREEN_L, PINK, PINK_L, PEACH = "#6B7F4B", "#A9BA8B", "#E8788F", "#F4B6C2", "#F2A883"
PALETTE = [GREEN, PINK, PEACH, GREEN_L, PINK_L, "#8E6C8A", "#E9C46A"]
STATUS_COLOR = {
    "Aman": GREEN, "Menipis": PEACH, "Perlu Restock": PEACH, "Habis": "#C0392B",
    "On Track": GREEN, "Tepat Waktu": GREEN_L, "Terlambat": PEACH, "Overdue": PINK, "Batal": "#9E9E9E",
}

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&family=Dancing+Script:wght@600;700&display=swap');
html, body, .stApp, [class*="st-"] { font-family: 'Poppins', sans-serif; }
.stApp { background: #FBF4F1; }
header[data-testid="stHeader"] { background: transparent; }
.block-container { padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1400px; }

/* ---------- sidebar ---------- */
section[data-testid="stSidebar"] { background: #66794A; }
section[data-testid="stSidebar"] .block-container { padding-top: 1.2rem; }
section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label { color: #FFFFFF; }
.logo { font-family: 'Dancing Script', cursive; font-size: 2.3rem; font-weight: 700; color: #fff; line-height: 1; margin-bottom: 2px; }
.logo-sub { font-size: .55rem; letter-spacing: .16em; color: #E4EBD6; margin-bottom: 14px; }
.admin-pill { display: inline-block; background: rgba(255,255,255,.18); color: #fff; padding: 3px 18px; border-radius: 999px;
              font-size: .62rem; letter-spacing: .12em; margin-bottom: 22px; }
section[data-testid="stSidebar"] div[role="radiogroup"] { gap: 6px; }
section[data-testid="stSidebar"] div[role="radiogroup"] > label { padding: 11px 16px; border-radius: 14px; width: 100%; cursor: pointer; }
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child { display: none; }
section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover { background: rgba(255,255,255,.14); }
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) { background: #FBF4F1; }
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) p { color: #3F4F2B; font-weight: 600; }
.side-foot { margin-top: 40px; color: #E4EBD6; font-size: .7rem; line-height: 1.5; }
.side-flower { font-size: 3.4rem; opacity: .45; margin-top: 8px; }

/* ---------- header ---------- */
.topbar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 18px; }
.topbar .ttl { font-size: 1.9rem; font-weight: 600; color: #2F3A22; line-height: 1.15; }
.topbar .sub { font-size: .85rem; color: #6f6f6f; }
.topbar .right { display: flex; align-items: center; gap: 18px; }
.bell { position: relative; font-size: 1.3rem; }
.badge { position: absolute; top: -6px; right: -9px; background: #E8788F; color: #fff; font-size: .6rem; font-weight: 600;
         border-radius: 999px; padding: 1px 6px; }
.user { display: flex; align-items: center; gap: 8px; font-size: .8rem; font-weight: 600; color: #2F3A22; }
.avatar { width: 38px; height: 38px; border-radius: 50%; background: #F9DCE4; display: flex; align-items: center; justify-content: center; font-size: 1.2rem; }

/* ---------- kartu KPI ---------- */
.kpi { border-radius: 20px; padding: 16px 18px; position: relative; overflow: hidden; min-height: 132px; box-shadow: 0 2px 10px rgba(90,70,60,.06); }
.kpi .top { display: flex; align-items: center; gap: 10px; }
.kpi .ico { width: 36px; height: 36px; border-radius: 12px; background: rgba(255,255,255,.75); display: flex; align-items: center; justify-content: center; font-size: 1.05rem; }
.kpi .lbl { font-size: .86rem; color: #4b4b4b; font-weight: 500; }
.kpi .val { font-size: 1.65rem; font-weight: 700; color: #2F3A22; margin: 8px 0 2px; }
.kpi .sub { font-size: .72rem; color: #7a7a7a; }
.kpi .deco { position: absolute; right: 10px; bottom: -8px; font-size: 3.4rem; opacity: .30; }
.kpi-green { background: #E4EBD6; } .kpi-pink { background: #F9DCE4; } .kpi-peach { background: #FCE5D6; }

/* ---------- panel putih ---------- */
.panel { background: #fff; border-radius: 20px; padding: 16px 18px; box-shadow: 0 2px 10px rgba(90,70,60,.06); }
.panel .ph { font-weight: 600; color: #2F3A22; font-size: .95rem; margin-bottom: 8px; }
.row { display: flex; align-items: center; gap: 12px; padding: 9px 0; border-top: 1px solid #F1E8E4; }
.row:first-of-type { border-top: none; }
.thumb { width: 44px; height: 44px; border-radius: 12px; background: #F9DCE4; display: flex; align-items: center; justify-content: center; font-size: 1.4rem; flex-shrink: 0; }
.row .meta { flex: 1; min-width: 0; }
.row .nm { font-size: .78rem; font-weight: 600; color: #2F3A22; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.row .amt { font-size: .85rem; font-weight: 700; color: #2F3A22; }
.row .dt { font-size: .68rem; color: #8a8a8a; }
.pill { font-size: .66rem; padding: 3px 10px; border-radius: 999px; font-weight: 600; white-space: nowrap; }
.pill-green { background: #E4EBD6; color: #4D5E33; } .pill-peach { background: #FCE5D6; color: #A5602F; }
.pill-pink { background: #F9DCE4; color: #A8405A; } .pill-grey { background: #EEE; color: #666; }
.banner { border-radius: 20px; min-height: 250px; padding: 28px 30px; position: relative; overflow: hidden;
          background: linear-gradient(120deg, #F4B6C2 0%, #F8D3DB 50%, #E4EBD6 100%); box-shadow: 0 2px 10px rgba(90,70,60,.06); }
.banner .bt { font-size: 1.9rem; font-weight: 600; color: #6B2F42; line-height: 1.2; margin-top: 70px; }
.banner .bs { font-size: .85rem; color: #7a4a58; margin-top: 8px; }
.banner .bf { position: absolute; right: 24px; top: 14px; font-size: 6rem; opacity: .55; }

/* ---------- komponen bawaan ---------- */
div[data-testid="stPlotlyChart"] { background: #fff; border-radius: 20px; padding: 8px 10px; box-shadow: 0 2px 10px rgba(90,70,60,.06); }
div[data-testid="stDataFrame"] { border-radius: 16px; overflow: hidden; box-shadow: 0 2px 10px rgba(90,70,60,.06); }
button[data-baseweb="tab"][aria-selected="true"] { color: #66794A; font-weight: 600; }
div[data-baseweb="tab-highlight"] { background-color: #66794A; }
div[data-testid="stExpander"] { background: #fff; border-radius: 16px; border: none; box-shadow: 0 2px 10px rgba(90,70,60,.06); }
.stDownloadButton button, .stButton button { border-radius: 999px; border: 1px solid #66794A; color: #66794A; background: #fff; }
h2, h3 { color: #2F3A22; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# ------------------------------------------------------------------ helper tampilan
def rp(x):
    try:
        return "Rp " + f"{float(x):,.0f}".replace(",", ".")
    except Exception:
        return "-"


def pct(x):
    try:
        return f"{float(x) * 100:.1f}%"
    except Exception:
        return "-"


def num(x):
    return f"{int(x):,}".replace(",", ".")


DECO = {"green": "🌿", "pink": "🌸", "peach": "🌼"}


def cards(items):
    """items: (label, nilai, keterangan, ikon, tema[green|pink|peach])"""
    cols = st.columns(len(items))
    for col, (label, value, sub, icon, theme) in zip(cols, items):
        col.markdown(
            f'<div class="kpi kpi-{theme}"><div class="top"><div class="ico">{icon}</div><div class="lbl">{html.escape(str(label))}</div></div>'
            f'<div class="val">{html.escape(str(value))}</div><div class="sub">{html.escape(str(sub))}</div>'
            f'<div class="deco">{DECO[theme]}</div></div>',
            unsafe_allow_html=True,
        )


def style(fig, height=340):
    fig.update_layout(
        template="plotly_white", height=height, margin=dict(l=10, r=10, t=48, b=10),
        legend=dict(orientation="h", y=-0.22), colorway=PALETTE,
        paper_bgcolor="#fff", plot_bgcolor="#fff",
        font=dict(family="Poppins, sans-serif", color="#4b4b4b"),
        title=dict(font=dict(size=15, color="#2F3A22")),
    )
    fig.update_xaxes(gridcolor="#F1E8E4")
    fig.update_yaxes(gridcolor="#F1E8E4")
    return fig


def plot(fig, height=340):
    st.plotly_chart(style(fig, height), theme=None)


def table(df, money=(), pct_cols=(), dates=()):
    cfg = {}
    for c in money:
        if c in df.columns:
            cfg[c] = st.column_config.NumberColumn(c.replace("_", " "), format="Rp %d")
    for c in pct_cols:
        if c in df.columns:
            cfg[c] = st.column_config.NumberColumn(c.replace("_", " "), format="%.1f%%")
    for c in dates:
        if c in df.columns:
            cfg[c] = st.column_config.DateColumn(c.replace("_", " "), format="DD/MM/YYYY")
    st.dataframe(df, hide_index=True, column_config=cfg)


def empty(df, msg="Belum ada data untuk bagian ini."):
    if df is None or len(df) == 0:
        st.info(msg)
        return True
    return False


def cat_icon(cat):
    c = str(cat).lower()
    if "snack" in c:
        return "🍫"
    if "money" in c:
        return "💵"
    if "artificial" in c:
        return "🌸"
    if "kemasan" in c:
        return "🎁"
    if "pita" in c:
        return "🎀"
    if "alat" in c:
        return "✂️"
    if "bunga" in c:
        return "🌷"
    return "💐"


def pill_class(status):
    s = str(status).lower()
    if s in ("selesai", "lunas", "aman", "tepat waktu", "on track"):
        return "pill-green"
    if s in ("diproses", "menunggu", "dp", "belum lunas", "menipis", "perlu restock"):
        return "pill-peach"
    if s in ("dibatalkan", "batal", "habis", "overdue", "terlambat"):
        return "pill-pink"
    return "pill-grey"


# ------------------------------------------------------------------ sidebar (navigasi)
PAGES = ["🏠 Overview", "📈 Sales Analytics", "📦 Inventory Monitoring", "🧾 Order Monitoring", "💰 Financial Analytics"]

st.sidebar.markdown(
    '<div class="logo">by.ficlo</div><div class="logo-sub">FLORIST &amp; BOUQUET DECOR</div>'
    '<div class="admin-pill">ADMIN PANEL</div>',
    unsafe_allow_html=True,
)
page = st.sidebar.radio("Navigasi", PAGES, label_visibility="collapsed")

# ------------------------------------------------------------------ slot header, lalu kontrol data
header_slot = st.container()
ctrl = st.expander("🔎 Filter & sumber data")


@st.cache_data(show_spinner="Membaca & membersihkan data...")
def load(file_bytes, path, mtime):
    src = io.BytesIO(file_bytes) if file_bytes else path
    return dl.load_all(src)


with ctrl:
    u1, u2 = st.columns([3, 1])
    with u1:
        upload = st.file_uploader("Ganti dataset (.xlsx)", type=["xlsx"])
    with u2:
        st.write("")
        st.write("")
        if st.button("🔄 Muat ulang data"):
            st.cache_data.clear()
            st.rerun()

if upload is not None:
    d = load(upload.getvalue(), None, 0)
elif os.path.exists(DATA_PATH):
    d = load(None, DATA_PATH, os.path.getmtime(DATA_PATH))
else:
    st.error(f"File '{DATA_PATH}' tidak ditemukan. Taruh file Excel di folder yang sama dengan app.py, atau unggah lewat panel Filter.")
    st.stop()

produk, bahan, penjualan = d["produk"], d["bahan"], d["penjualan"]
mutasi, pesanan, produksi = d["mutasi"], d["pesanan"], d["produksi"]
pengeluaran, kas, today = d["pengeluaran"], d["kas"], d["today"]

if len(penjualan):
    dmin, dmax = penjualan["Tanggal"].min().date(), penjualan["Tanggal"].max().date()
else:
    dmin = dmax = today.date()
opt_kat = sorted(penjualan["Kategori"].dropna().unique())
opt_kanal = sorted(penjualan["Kanal_Penjualan"].dropna().unique())
opt_status = sorted(penjualan["Status_Transaksi"].dropna().unique())

with ctrl:
    st.markdown("**Filter data penjualan**")
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        rentang = st.date_input("Periode", value=(dmin, dmax), min_value=dmin, max_value=dmax)
    with f2:
        f_kat = st.multiselect("Kategori produk", opt_kat, default=opt_kat)
    with f3:
        f_kanal = st.multiselect("Kanal penjualan", opt_kanal, default=opt_kanal)
    with f4:
        f_status = st.multiselect("Status transaksi", opt_status, default=opt_status)

start, end = (rentang if isinstance(rentang, (list, tuple)) and len(rentang) == 2 else (dmin, dmax))
pj = penjualan[
    (penjualan["Tanggal"] >= pd.Timestamp(start))
    & (penjualan["Tanggal"] <= pd.Timestamp(end))
    & penjualan["Kategori"].isin(f_kat)
    & penjualan["Kanal_Penjualan"].isin(f_kanal)
    & penjualan["Status_Transaksi"].isin(f_status)
]
pg = pengeluaran[(pengeluaran["Tanggal"] >= pd.Timestamp(start)) & (pengeluaran["Tanggal"] <= pd.Timestamp(end))]

# ---- ringkasan peringatan (untuk lonceng di header)
stok_items = pd.concat([
    produk.assign(Nama=produk["Nama_Produk"], Min=produk["Stok_Minimum"], Jenis="Produk", Kat=produk["Kategori"],
                  Satuan=produk["Satuan"] if "Satuan" in produk.columns else "pcs")[["Nama", "Stok_Saat_Ini", "Min", "Satuan", "Status_Stok", "Jenis", "Kat"]],
    bahan.assign(Nama=bahan["Nama_Bahan"], Min=bahan["Stok_Minimum"], Jenis="Bahan", Kat=bahan["Kategori_Bahan"])[["Nama", "Stok_Saat_Ini", "Min", "Satuan", "Status_Stok", "Jenis", "Kat"]],
], ignore_index=True)
n_stok = int((stok_items["Status_Stok"] != "Aman").sum())
n_overdue = int((pesanan["Status_SLA"] == "Overdue").sum())
n_notif = n_stok + n_overdue + len(d["dq"]["peringatan"])

SUBTITLE = {
    PAGES[0]: "Selamat datang kembali, Admin! Ringkasan kondisi bisnis by.ficlo.",
    PAGES[1]: "Tren penjualan, produk terlaris, dan pendapatan bulanan.",
    PAGES[2]: "Monitoring stok bunga, wrapping, pita, aksesoris, dan produk jadi.",
    PAGES[3]: "Pesanan masuk, diproses, selesai, dan dibatalkan.",
    PAGES[4]: "Modal, pengeluaran, pendapatan, dan keuntungan.",
}
TITLE = {PAGES[0]: "Overview", PAGES[1]: "Sales Analytics", PAGES[2]: "Inventory Monitoring",
         PAGES[3]: "Order Monitoring", PAGES[4]: "Financial Analytics"}
with header_slot:
    st.markdown(
        f'<div class="topbar"><div><div class="ttl">☰ {TITLE[page]}</div><div class="sub">{SUBTITLE[page]}</div></div>'
        f'<div class="right"><div class="bell">🔔<span class="badge">{n_notif}</span></div>'
        f'<div class="user"><div class="avatar">🌸</div>Admin Ficlo</div></div></div>',
        unsafe_allow_html=True,
    )

st.sidebar.markdown(
    f'<div class="side-foot">Data per {today.strftime("%d %b %Y")}<br>{len(pj)} baris penjualan terpilih</div>'
    '<div class="side-flower">🌸🌿</div>',
    unsafe_allow_html=True,
)


# ------------------------------------------------------------------ 1. OVERVIEW
def page_overview():
    omzet, laba = pj["Total_Penjualan"].sum(), pj["Laba_Kotor"].sum()
    n_trx = pj["ID_Transaksi"].nunique()
    cards([
        ("Total Pendapatan", rp(omzet), f"Margin kotor {pct(laba / omzet)}" if omzet else "-", "💰", "peach"),
        ("Total Pesanan", num(n_trx), f"+ {len(pesanan)} custom order (PO)", "🛍️", "green"),
        ("Keuntungan (Laba Kotor)", rp(laba), "Pendapatan − HPP", "📈", "pink"),
        ("Produk Terjual", num(pj["Jumlah_Terjual"].sum()), "Total item pada periode", "📦", "green"),
    ])
    st.write("")

    if len(pj) == 0:
        st.info("Tidak ada penjualan pada filter yang dipilih.")
    else:
        c1, c2 = st.columns([2, 1])
        with c1:
            gran = st.radio("Tampilkan per", ["Harian", "Mingguan", "Bulanan"], horizontal=True, key="gran_ov")
            freq = {"Harian": "D", "Mingguan": "W", "Bulanan": "MS"}[gran]
            tren = pj.groupby(pd.Grouper(key="Tanggal", freq=freq))["Total_Penjualan"].sum().reset_index()
            fig = go.Figure(go.Scatter(
                x=tren["Tanggal"], y=tren["Total_Penjualan"], mode="lines+markers",
                line=dict(color=GREEN, width=3, shape="spline"), marker=dict(size=8, color=GREEN),
                fill="tozeroy", fillcolor="rgba(107,127,75,0.15)",
                hovertemplate="%{x|%d %b %Y}<br>Rp %{y:,.0f}<extra></extra>"))
            fig.update_layout(title="Grafik Pendapatan", yaxis_title=None, xaxis_title=None)
            plot(fig, 340)
        with c2:
            terbaru = pj.sort_values(["Tanggal", "ID_Transaksi"], ascending=False).head(4)
            rows = "".join(
                f'<div class="row"><div class="thumb">{cat_icon(r.Kategori)}</div><div class="meta">'
                f'<div class="nm">{html.escape(r.ID_Transaksi)} · {html.escape(str(r.Nama_Produk))}</div>'
                f'<div class="amt">{rp(r.Total_Penjualan)}</div><div class="dt">{r.Tanggal.strftime("%d %b %Y")}</div></div>'
                f'<span class="pill {pill_class(r.Status_Pembayaran)}">{html.escape(str(r.Status_Pembayaran))}</span></div>'
                for r in terbaru.itertuples()
            )
            st.markdown(f'<div class="panel" style="min-height:392px"><div class="ph">Pesanan Terbaru</div>{rows}</div>', unsafe_allow_html=True)

    st.write("")
    c1, c2 = st.columns([1, 2])
    with c1:
        urut = stok_items.assign(_s=stok_items["Status_Stok"].map({"Habis": 0, "Menipis": 1, "Perlu Restock": 1, "Aman": 2}),
                                 _r=stok_items["Stok_Saat_Ini"] / stok_items["Min"].where(stok_items["Min"] > 0))
        urut = urut.sort_values(["_s", "_r"]).head(3)
        judul = "Stok Menipis" if n_stok else "Stok Terendah (semua masih aman)"
        rows = "".join(
            f'<div class="row"><div class="thumb">{cat_icon(r.Kat)}</div><div class="meta"><div class="nm">{html.escape(str(r.Nama))}</div>'
            f'<div class="dt">Stok tersisa: {r.Stok_Saat_Ini:g} {html.escape(str(r.Satuan))}</div></div>'
            f'<span class="pill {pill_class(r.Status_Stok)}">&lt; {r.Min:g}</span></div>'
            for r in urut.itertuples()
        )
        st.markdown(f'<div class="panel" style="min-height:250px"><div class="ph">⚠️ {judul}</div>{rows}</div>', unsafe_allow_html=True)
    with c2:
        st.markdown(
            '<div class="banner"><div class="bf">💐</div><div class="bt">Make It Bloom<br>With by.ficlo ♡</div>'
            '<div class="bs">Buket cantik, untuk setiap momen spesialmu.</div></div>',
            unsafe_allow_html=True,
        )

    if len(pj):
        st.write("")
        c1, c2 = st.columns(2)
        with c1:
            kat = pj.groupby("Kategori")["Total_Penjualan"].sum().reset_index()
            plot(px.pie(kat, names="Kategori", values="Total_Penjualan", hole=0.55, title="Pendapatan per Kategori"), 320)
        with c2:
            top = (pj.groupby("Nama_Produk").agg(Omzet=("Total_Penjualan", "sum"), Terjual=("Jumlah_Terjual", "sum"))
                   .reset_index().sort_values("Omzet", ascending=False).head(5).sort_values("Omzet"))
            fig = px.bar(top, x="Omzet", y="Nama_Produk", orientation="h", text="Terjual", title="Top 5 Produk Terjual (label = jumlah)")
            fig.update_traces(marker_color=PINK)
            fig.update_layout(yaxis_title=None)
            plot(fig, 320)

    st.write("")
    with st.expander(f"🧹 Catatan kualitas data ({len(d['dq']['peringatan'])})"):
        st.caption("Catatan otomatis dari proses pembersihan. Rapikan di Excel supaya angka makin akurat.")
        if not d["dq"]["peringatan"]:
            st.success("Tidak ada masalah yang terdeteksi ✅")
        for w in d["dq"]["peringatan"]:
            st.warning(w)
        if len(d["dq"]["selisih_penjualan"]):
            st.markdown("**Selisih nilai Excel vs hitung ulang (PENJUALAN)**")
            table(d["dq"]["selisih_penjualan"], money=["Nilai_di_Excel", "Nilai_Hitung_Ulang"], dates=["Tanggal"])
        if len(d["dq"]["selisih_stok"]):
            st.markdown("**Selisih stok Excel vs hitung ulang**")
            table(d["dq"]["selisih_stok"])


# ------------------------------------------------------------------ 2. SALES
def page_sales():
    if len(pj) == 0:
        st.info("Tidak ada penjualan pada filter yang dipilih.")
        return
    omzet, laba = pj["Total_Penjualan"].sum(), pj["Laba_Kotor"].sum()
    n_trx = pj["ID_Transaksi"].nunique()
    cards([
        ("Pendapatan", rp(omzet), "Periode terpilih", "💰", "peach"),
        ("Laba Kotor", rp(laba), f"Margin {pct(laba / omzet)}" if omzet else "-", "📈", "green"),
        ("Transaksi", num(n_trx), f"Item terjual: {num(pj['Jumlah_Terjual'].sum())}", "🛍️", "pink"),
        ("Rata-rata / Transaksi", rp(omzet / n_trx) if n_trx else "-", "Nilai keranjang", "🧺", "green"),
    ])
    st.write("")
    t1, t2, t3, t4, t5 = st.tabs(["Tren Penjualan", "Produk Terlaris", "Pendapatan Bulanan", "Kanal & Pelanggan", "Data Transaksi"])

    with t1:
        gran = st.radio("Granularitas", ["Harian", "Mingguan", "Bulanan"], horizontal=True, key="gran_sales")
        freq = {"Harian": "D", "Mingguan": "W", "Bulanan": "MS"}[gran]
        tren = pj.groupby([pd.Grouper(key="Tanggal", freq=freq), "Kategori"])["Total_Penjualan"].sum().reset_index()
        fig = px.bar(tren, x="Tanggal", y="Total_Penjualan", color="Kategori", title=f"Pendapatan {gran} per Kategori")
        fig.update_layout(yaxis_title="Rupiah", xaxis_title=None, legend_title=None)
        plot(fig, 380)
        kum = pj.groupby("Tanggal")["Total_Penjualan"].sum().cumsum().reset_index(name="Omzet_Kumulatif")
        fig = px.area(kum, x="Tanggal", y="Omzet_Kumulatif", title="Pendapatan Kumulatif")
        fig.update_traces(line_color=GREEN, fillcolor="rgba(107,127,75,0.15)")
        plot(fig, 300)

    with t2:
        ringkas = (pj.groupby("Nama_Produk").agg(Omzet=("Total_Penjualan", "sum"), Terjual=("Jumlah_Terjual", "sum"),
                                                 Laba_Kotor=("Laba_Kotor", "sum")).reset_index())
        ringkas["Margin_Persen"] = (ringkas["Laba_Kotor"] / ringkas["Omzet"]).fillna(0) * 100
        metrik = st.radio("Urutkan berdasarkan", ["Omzet", "Terjual", "Laba_Kotor"], horizontal=True)
        top = ringkas.sort_values(metrik, ascending=False).head(10).sort_values(metrik)
        fig = px.bar(top, x=metrik, y="Nama_Produk", orientation="h", title=f"Produk Terlaris berdasarkan {metrik.replace('_', ' ')}")
        fig.update_traces(marker_color=PINK)
        fig.update_layout(yaxis_title=None)
        plot(fig, 380)
        fig = px.bar(ringkas.sort_values("Margin_Persen"), x="Margin_Persen", y="Nama_Produk", orientation="h", title="Margin Kotor per Produk (%)")
        fig.update_traces(marker_color=GREEN)
        fig.update_layout(yaxis_title=None)
        plot(fig, 320)
        table(ringkas.sort_values("Omzet", ascending=False), money=["Omzet", "Laba_Kotor"], pct_cols=["Margin_Persen"])

    with t3:
        bln = pj.groupby("Bulan").agg(Pendapatan=("Total_Penjualan", "sum"), Laba_Kotor=("Laba_Kotor", "sum")).reset_index()
        fig = px.bar(bln, x="Bulan", y=["Pendapatan", "Laba_Kotor"], barmode="group", title="Pendapatan & Laba Kotor per Bulan",
                     color_discrete_map={"Pendapatan": GREEN, "Laba_Kotor": PINK})
        fig.update_layout(yaxis_title="Rupiah", xaxis_title=None, legend_title=None)
        plot(fig, 380)
        bk = pj.groupby(["Bulan", "Kategori"])["Total_Penjualan"].sum().reset_index()
        plot(px.bar(bk, x="Bulan", y="Total_Penjualan", color="Kategori", title="Pendapatan Bulanan per Kategori"), 320)
        table(bln, money=["Pendapatan", "Laba_Kotor"])

    with t4:
        c1, c2 = st.columns(2)
        for col, kolom, judul in [(c1, "Kanal_Penjualan", "Pendapatan per Kanal"), (c2, "Metode_Pembayaran", "Per Metode Pembayaran"),
                                  (c1, "Tipe_Pelanggan", "Per Tipe Pelanggan"), (c2, "Status_Pembayaran", "Status Pembayaran")]:
            with col:
                g = pj.groupby(kolom)["Total_Penjualan"].sum().reset_index()
                plot(px.pie(g, names=kolom, values="Total_Penjualan", hole=0.55, title=judul), 300)
        belum = pj[pj["Status_Pembayaran"] != "Lunas"]["Total_Penjualan"].sum()
        st.info(f"Penjualan belum lunas (piutang): **{rp(belum)}**")

    with t5:
        kolom = ["ID_Transaksi", "Tanggal", "Nama_Produk", "Kategori", "Jumlah_Terjual", "Harga_Jual_Satuan", "Diskon",
                 "Total_Penjualan", "Total_HPP", "Laba_Kotor", "Metode_Pembayaran", "Status_Pembayaran", "Kanal_Penjualan",
                 "Tipe_Pelanggan", "Status_Transaksi"]
        table(pj[kolom], money=["Harga_Jual_Satuan", "Diskon", "Total_Penjualan", "Total_HPP", "Laba_Kotor"], dates=["Tanggal"])
        st.download_button("⬇️ Unduh CSV", pj[kolom].to_csv(index=False).encode("utf-8"), "penjualan_terfilter.csv", "text/csv")


# ------------------------------------------------------------------ 3. INVENTORY
def stok_chart(df, nama_col, judul, tinggi=380):
    long = df.melt(id_vars=[nama_col], value_vars=["Stok_Saat_Ini", "Stok_Minimum"], var_name="Jenis", value_name="Jumlah")
    fig = px.bar(long, x=nama_col, y="Jumlah", color="Jenis", barmode="group", title=judul,
                 color_discrete_map={"Stok_Saat_Ini": GREEN, "Stok_Minimum": PINK})
    fig.update_layout(xaxis_title=None, legend_title=None)
    return fig


def page_inventory():
    cards([
        ("Produk Aktif", num((produk["Status"] == "Aktif").sum()), f"{len(produk)} produk di katalog", "💐", "green"),
        ("Produk Menipis / Habis", num((produk["Status_Stok"] != "Aman").sum()), "Perlu produksi ulang", "🎁", "pink"),
        ("Bahan Perlu Restock", num((bahan["Status_Stok"] != "Aman").sum()), f"dari {len(bahan)} jenis bahan", "⚠️", "peach"),
        ("Nilai Persediaan Bahan", rp(bahan["Nilai_Persediaan"].sum()), "Stok × harga beli", "🧮", "green"),
    ])
    st.write("")
    t1, t2, t3 = st.tabs(["Produk Jadi", "Bunga, Wrapping, Pita & Aksesoris", "Mutasi & Produksi"])

    with t1:
        plot(stok_chart(produk, "Nama_Produk", "Stok Produk vs Stok Minimum"), 380)
        table(produk[["ID_Produk", "Nama_Produk", "Kategori", "Harga_Jual", "HPP_Standar", "Persen_Margin", "Stok_Awal",
                      "Stok_Saat_Ini", "Stok_Minimum", "Status_Stok", "Status"]].assign(Persen_Margin=lambda x: x["Persen_Margin"] * 100),
              money=["Harga_Jual", "HPP_Standar"], pct_cols=["Persen_Margin"])

    with t2:
        ringkas = bahan.groupby("Kategori_Bahan").agg(Jenis=("ID_Bahan", "count"), Nilai=("Nilai_Persediaan", "sum"),
                                                      Restock=("Status_Stok", lambda s: int((s != "Aman").sum()))).reset_index()
        temas = ["green", "pink", "peach"]
        cards([(r.Kategori_Bahan, f"{r.Jenis} jenis", f"Nilai {rp(r.Nilai)} · {r.Restock} perlu restock", cat_icon(r.Kategori_Bahan), temas[i % 3])
               for i, r in enumerate(ringkas.itertuples())][:4])
        st.write("")
        pilih = st.radio("Kategori bahan", ["Semua"] + sorted(bahan["Kategori_Bahan"].unique()), horizontal=True)
        b = bahan if pilih == "Semua" else bahan[bahan["Kategori_Bahan"] == pilih]
        c1, c2 = st.columns(2)
        with c1:
            g = b.groupby("Status_Stok").size().reset_index(name="Jumlah")
            plot(px.pie(g, names="Status_Stok", values="Jumlah", hole=0.55, title="Status Stok Bahan",
                        color="Status_Stok", color_discrete_map=STATUS_COLOR), 300)
        with c2:
            g = b.groupby("Kategori_Bahan")["Nilai_Persediaan"].sum().reset_index()
            fig = px.bar(g, x="Kategori_Bahan", y="Nilai_Persediaan", title="Nilai Persediaan per Kategori")
            fig.update_traces(marker_color=PINK)
            fig.update_layout(xaxis_title=None)
            plot(fig, 300)
        rendah = b.dropna(subset=["Rasio_Stok"]).sort_values("Rasio_Stok").head(10)
        if len(rendah):
            plot(stok_chart(rendah, "Nama_Bahan", "10 Bahan dengan Rasio Stok/Minimum Terendah"), 380)
        table(b[["ID_Bahan", "Nama_Bahan", "Kategori_Bahan", "Supplier", "Satuan", "Stok_Awal", "Mutasi_Masuk", "Mutasi_Keluar",
                 "Stok_Saat_Ini", "Stok_Minimum", "Harga_Beli_Satuan", "Nilai_Persediaan", "Status_Stok"]],
              money=["Harga_Beli_Satuan", "Nilai_Persediaan"])

    with t3:
        st.markdown("**Mutasi stok**")
        if not empty(mutasi, "Belum ada mutasi stok."):
            g = mutasi.groupby(["Alasan_Mutasi", "Jenis_Mutasi"])["Jumlah"].sum().reset_index()
            plot(px.bar(g, x="Alasan_Mutasi", y="Jumlah", color="Jenis_Mutasi", barmode="group", title="Mutasi per Alasan"), 300)
            table(mutasi.drop(columns=["Nama_Item"], errors="ignore"), dates=["Tanggal"])
        st.markdown("**Produksi (pemakaian bahan)**")
        if not empty(produksi, "Belum ada catatan produksi."):
            g = produksi.groupby("Nama_Produk")[["Total_Biaya_Bahan", "Biaya_Tambahan"]].sum().reset_index()
            plot(px.bar(g, x="Nama_Produk", y=["Total_Biaya_Bahan", "Biaya_Tambahan"], title="Biaya Produksi per Produk"), 300)
            table(produksi, money=["Harga_Beli_Satuan", "Total_Biaya_Bahan", "Biaya_Tambahan", "Total_Biaya_Produksi"], dates=["Tanggal_Produksi"])


# ------------------------------------------------------------------ 4. ORDERS
def page_orders():
    masuk = len(pesanan)
    proses = int((pesanan["Status_Pesanan"].isin(["Diproses", "Menunggu"])).sum())
    selesai_n = int((pesanan["Status_Pesanan"] == "Selesai").sum())
    batal = int(pesanan["Status_Pesanan"].isin(["Dibatalkan", "Batal"]).sum())
    cards([
        ("Pesanan Masuk", num(masuk), "Total pesanan custom/PO", "📥", "green"),
        ("Diproses", num(proses), "Sedang dikerjakan", "🛠️", "peach"),
        ("Selesai", num(selesai_n), "Sudah diserahkan", "✅", "green"),
        ("Dibatalkan", num(batal), "Pesanan batal", "✖️", "pink"),
    ])
    if empty(pesanan, "Belum ada data pesanan."):
        return
    st.write("")
    selesai = pesanan[pesanan["Status_SLA"].isin(["Tepat Waktu", "Terlambat"])]
    ontime = (selesai["Status_SLA"] == "Tepat Waktu").mean() if len(selesai) else None
    cards([
        ("Overdue", num((pesanan["Status_SLA"] == "Overdue").sum()), "Lewat deadline, belum selesai", "⏰", "pink"),
        ("Tepat Waktu", pct(ontime) if ontime is not None else "-", "Dari pesanan yang sudah selesai", "🎯", "green"),
        ("Total Nilai Pesanan", rp(pesanan["Total_Harga"].sum()), f"DP diterima {rp(pesanan['DP_Dibayar'].sum())}", "💰", "peach"),
        ("Sisa Tagihan", rp(pesanan["Sisa_Pembayaran"].sum()), "Total − DP", "🧾", "peach"),
    ])
    st.write("")
    c1, c2 = st.columns(2)
    with c1:
        alur = pd.DataFrame({"Tahap": ["Masuk", "Diproses", "Selesai", "Dibatalkan"], "Jumlah": [masuk, proses, selesai_n, batal]})
        fig = px.bar(alur, x="Tahap", y="Jumlah", title="Alur Pesanan", color="Tahap",
                     color_discrete_map={"Masuk": GREEN_L, "Diproses": PEACH, "Selesai": GREEN, "Dibatalkan": PINK})
        fig.update_layout(showlegend=False, xaxis_title=None)
        plot(fig, 320)
    with c2:
        g = pesanan.groupby("Momen_Acara")["Total_Harga"].sum().reset_index()
        plot(px.pie(g, names="Momen_Acara", values="Total_Harga", hole=0.55, title="Nilai Pesanan per Momen Acara"), 320)

    tl = pesanan.dropna(subset=["Tanggal_Pesanan", "Tanggal_Deadline"]).copy()
    if len(tl):
        tl["Deadline_Akhir"] = tl["Tanggal_Deadline"].where(tl["Tanggal_Deadline"] > tl["Tanggal_Pesanan"], tl["Tanggal_Pesanan"] + pd.Timedelta(days=1))
        fig = px.timeline(tl, x_start="Tanggal_Pesanan", x_end="Deadline_Akhir", y="ID_Pesanan", color="Status_SLA",
                          color_discrete_map=STATUS_COLOR, title="Timeline Pesanan (tanggal pesan → deadline)")
        fig.update_yaxes(autorange="reversed")
        st.write("")
        plot(fig, 140 + 40 * len(tl))

    st.subheader("Daftar Pesanan")
    opsi = sorted(pesanan["Status_SLA"].unique())
    pilih = st.multiselect("Filter SLA", opsi, default=opsi)
    tabel = pesanan[pesanan["Status_SLA"].isin(pilih)].sort_values("Tanggal_Deadline")
    table(tabel[["ID_Pesanan", "Tanggal_Pesanan", "Kode_Pelanggan", "Nama_Produk", "Jumlah", "Jenis_Order", "Momen_Acara",
                 "Tanggal_Deadline", "Hari_ke_Deadline", "Status_SLA", "Status_Pesanan", "Total_Harga", "DP_Dibayar",
                 "Sisa_Pembayaran", "Catatan"]],
          money=["Total_Harga", "DP_Dibayar", "Sisa_Pembayaran"], dates=["Tanggal_Pesanan", "Tanggal_Deadline"])


# ------------------------------------------------------------------ 5. FINANCIAL
def page_financial():
    incl = st.checkbox("Sertakan pembelian bahan/stok sebagai biaya operasional", value=False,
                       help="Biasanya HPP sudah memuat biaya bahan, jadi pembelian stok tidak dihitung lagi agar tidak ganda.")
    mask_bahan = pg["Kategori_Pengeluaran"].astype(str).str.contains("bahan|stok|pembelian", case=False, na=False)
    biaya = pg if incl else pg[~mask_bahan]

    omzet, hpp = pj["Total_Penjualan"].sum(), pj["Total_HPP"].sum()
    laba_kotor, ops = omzet - hpp, biaya["Jumlah"].sum()
    laba_bersih = laba_kotor - ops
    modal_masuk = kas.loc[kas["Jenis_Arus"] == "Masuk", "Nominal"].sum()
    modal_keluar = kas.loc[kas["Jenis_Arus"] == "Keluar", "Nominal"].sum()
    diterima = penjualan.loc[penjualan["Status_Pembayaran"] == "Lunas", "Total_Penjualan"].sum()
    est_kas = modal_masuk - modal_keluar + diterima - pengeluaran["Jumlah"].sum()
    piutang = penjualan.loc[penjualan["Status_Pembayaran"] != "Lunas", "Total_Penjualan"].sum()

    cards([
        ("Modal", rp(modal_masuk), f"Prive/penarikan {rp(modal_keluar)}", "🏦", "green"),
        ("Pengeluaran", rp(ops), "Biaya operasional periode ini", "🧾", "pink"),
        ("Pendapatan", rp(omzet), f"HPP {rp(hpp)}", "💰", "peach"),
        ("Keuntungan Bersih", rp(laba_bersih), f"Margin {pct(laba_bersih / omzet)}" if omzet else "-", "📈", "green"),
    ])
    st.write("")
    cards([
        ("Laba Kotor", rp(laba_kotor), f"Margin {pct(laba_kotor / omzet)}" if omzet else "-", "🌿", "green"),
        ("Estimasi Saldo Kas", rp(est_kas), "Modal + penjualan lunas − pengeluaran", "👛", "peach"),
        ("ROI", pct(laba_bersih / modal_masuk) if modal_masuk else "-", "Laba bersih ÷ modal masuk", "🎯", "pink"),
        ("Piutang", rp(piutang), "Penjualan belum lunas", "⏳", "green"),
    ])
    st.write("")

    c1, c2 = st.columns(2)
    with c1:
        fig = go.Figure(go.Waterfall(
            x=["Pendapatan", "HPP", "Laba Kotor", "Pengeluaran", "Laba Bersih"],
            measure=["absolute", "relative", "total", "relative", "total"],
            y=[omzet, -hpp, 0, -ops, 0],
            connector={"line": {"color": "#BBB"}},
            increasing={"marker": {"color": GREEN}}, decreasing={"marker": {"color": PINK}},
            totals={"marker": {"color": GREEN_L}},
        ))
        fig.update_layout(title="Dari Pendapatan ke Laba Bersih")
        plot(fig, 360)
    with c2:
        if len(pg):
            g = pg.groupby("Kategori_Pengeluaran")["Jumlah"].sum().reset_index()
            plot(px.pie(g, names="Kategori_Pengeluaran", values="Jumlah", hole=0.55, title="Pengeluaran per Kategori"), 360)
        else:
            st.info("Belum ada pengeluaran pada periode ini.")

    if len(pg):
        g = pg.groupby(["Tanggal", "Jenis_Biaya"])["Jumlah"].sum().reset_index()
        plot(px.bar(g, x="Tanggal", y="Jumlah", color="Jenis_Biaya", title="Pengeluaran Harian (Tetap vs Variabel)"), 300)
        table(pg.drop(columns=["ID_Referensi"], errors="ignore"), money=["Jumlah"], dates=["Tanggal"])

    st.subheader("Arus Kas Modal")
    if not empty(kas, "Belum ada catatan modal/kas."):
        fig = px.line(kas, x="Tanggal", y="Saldo_Kumulatif", markers=True, line_shape="hv", title="Saldo Kumulatif Modal")
        fig.update_traces(line_color=GREEN)
        plot(fig, 300)
        table(kas[["ID_Kas", "Tanggal", "Jenis_Arus", "Kategori_Kas", "Keterangan", "Nominal", "Metode_Pembayaran", "Saldo_Kumulatif"]],
              money=["Nominal", "Saldo_Kumulatif"], dates=["Tanggal"])


{
    PAGES[0]: page_overview,
    PAGES[1]: page_sales,
    PAGES[2]: page_inventory,
    PAGES[3]: page_orders,
    PAGES[4]: page_financial,
}[page]()
