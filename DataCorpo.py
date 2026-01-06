import streamlit as st
import pandas as pd
import numpy as np
import os

# ===============================
# PAGE CONFIG
# ===============================
st.set_page_config(
    page_title="RO Dashboard Corporate",
    layout="wide"
)

st.title("RO vs Non-RO Dashboard (Corporate)")
st.caption("Repeat Order Analysis – Executive View")

FILE_NAME = "DataCorpo.xlsx"

# ===============================
# LOAD DATA
# ===============================
if not os.path.exists(FILE_NAME):
    st.error(f"File {FILE_NAME} tidak ditemukan")
    st.stop()

df = pd.read_excel(FILE_NAME)

# ===============================
# PREPROCESSING
# ===============================
df["realisasidate"] = pd.to_datetime(df["realisasidate"])
df["Tahun"] = df["realisasidate"].dt.year
df["Bulan"] = df["realisasidate"].dt.month
df["Bulan_Nama"] = df["realisasidate"].dt.strftime("%b")

# ===============================
# FILTER (TOP BAR)
# ===============================
with st.container():
    col1, col2, col3, col4 = st.columns([1,1,1,2])

    with col1:
        tahun_selected = st.multiselect(
            "Tahun",
            sorted(df["Tahun"].unique()),
            default=sorted(df["Tahun"].unique())
        )

    with col2:
        bulan_selected = st.multiselect(
            "Bulan",
            list(range(1, 13)),
            default=list(range(1, 13)),
            format_func=lambda x: pd.to_datetime(str(x), format="%m").strftime("%B")
        )

    with col3:
        segmen_selected = st.multiselect(
            "Produk / Segmen",
            sorted(df["Segmen"].unique()),
            default=sorted(df["Segmen"].unique())
        )

    with col4:
        search_pt = st.text_input(
            "Search Account Name",
            placeholder="Ketik nama PT..."
        )

df = df[
    (df["Tahun"].isin(tahun_selected)) &
    (df["Bulan"].isin(bulan_selected)) &
    (df["Segmen"].isin(segmen_selected))
]

if search_pt:
    df = df[df["accountname"].str.contains(search_pt, case=False, na=False)]

# ===============================
# RO FLAG
# ===============================
df = df.sort_values(["Customerid", "Tahun", "realisasidate"])

df["trx_ke"] = (
    df.groupby(["Customerid", "Tahun"])
    .cumcount() + 1
)

df["RO_Status"] = np.where(df["trx_ke"] > 1, "RO", "Non-RO")

# ===============================
# KEY METRICS
# ===============================
# Semua customer
total_customer = df["Customerid"].nunique()

# Customer dengan transaksi RO
customer_ro_set = set(df[df["RO_Status"] == "RO"]["Customerid"].unique())

# Customer dengan transaksi Non-RO
customer_non_ro_set = set(df[df["RO_Status"] == "Non-RO"]["Customerid"].unique())

# Customer Non-RO murni = Non-RO tapi bukan RO
customer_non_ro_murni_set = customer_non_ro_set - customer_ro_set

# Hitung jumlah customer masing-masing
customer_ro = len(customer_ro_set)
customer_non_ro = len(customer_non_ro_murni_set)

# Produk dan transaksi RO/Non-RO
total_produk = len(df)
ro_produk = (df["RO_Status"] == "RO").sum()
non_ro_produk = (df["RO_Status"] == "Non-RO").sum()
ro_rate = ro_produk / total_produk * 100 if total_produk > 0 else 0

c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
c1.metric("Total Customer", f"{total_customer:,}")
c2.metric("Total Produk", f"{total_produk:,}")
c3.metric("RO Produk", f"{ro_produk:,}")
c4.metric("Non-RO Produk", f"{non_ro_produk:,}")
c5.metric("RO Rate", f"{ro_rate:.1f}%")
c6.metric("Customer RO", f"{customer_ro:,}")
c7.metric("Customer Non-RO", f"{customer_non_ro:,}")

# ===============================
# TREND RO VS NON-RO
# ===============================
st.markdown("## Tren RO vs Non-RO")

trend_df = (
    df.groupby([df["realisasidate"].dt.to_period("M"), "RO_Status"])
    .size()
    .reset_index(name="Jumlah")
)

trend_df["Periode"] = trend_df["realisasidate"].astype(str)

pivot_trend = (
    trend_df
    .pivot(index="Periode", columns="RO_Status", values="Jumlah")
    .fillna(0)
    .sort_index()
)

st.line_chart(pivot_trend)

# ===============================
# RO vs NON-RO PER PRODUK
# ===============================
st.markdown("## RO vs Non-RO per Jenis Produk")

produk_summary = (
    df.groupby(["Segmen", "RO_Status"])
    .size()
    .reset_index(name="Jumlah")
)

produk_summary["Total"] = (
    produk_summary
    .groupby("Segmen")["Jumlah"]
    .transform("sum")
)

produk_summary["Persen"] = (
    produk_summary["Jumlah"] / produk_summary["Total"] * 100
).round(2)

produk_summary["Display"] = (
    produk_summary["Jumlah"].astype(str)
    + " (" + produk_summary["Persen"].astype(str) + "%)"
)

pivot_produk = (
    produk_summary
    .pivot(index="Segmen", columns="RO_Status", values="Display")
    .fillna("0 (0%)")
    .reset_index()
    .rename(columns={"Segmen": "Product"})
)

st.dataframe(pivot_produk, use_container_width=True)

# ===============================
# PT DENGAN RO TERBANYAK
# ===============================
st.markdown("## PT dengan RO Terbanyak")

ro_only = df[df["RO_Status"] == "RO"]

pt_produk = (
    ro_only
    .groupby(["accountname", "Segmen"])
    .agg(Jumlah_Transaksi=("NoContract", "count"))
    .reset_index()
)

pt_produk["Detail"] = (
    pt_produk["Segmen"] + " = " +
    pt_produk["Jumlah_Transaksi"].astype(str)
)

pt_summary = (
    pt_produk
    .groupby("accountname")
    .agg(
        Jumlah_Produk_RO=("Segmen", "nunique"),
        Total_Transaksi_RO=("Jumlah_Transaksi", "sum"),
        Produk_dan_Transaksi=("Detail", lambda x: " | ".join(x))
    )
    .reset_index()
    .sort_values("Jumlah_Produk_RO", ascending=False)
)

st.dataframe(
    pt_summary,
    use_container_width=True,
)

# ===============================
# MULTI UNIT REALISASI
# ===============================
st.markdown("## 📅 Realisasi Multi Unit Detail per PT")

# 1. Ambil tanggal yang ada lebih dari 1 PT (multi unit)
multi_unit_dates = (
    df.groupby("realisasidate")["accountname"]
    .nunique()
    .reset_index()
    .rename(columns={"accountname": "Jumlah_PT"})
)

multi_unit_dates = multi_unit_dates[multi_unit_dates["Jumlah_PT"] > 1]

# 2. Filter data yang tanggalnya termasuk multi unit
df_multi_detail = df[df["realisasidate"].isin(multi_unit_dates["realisasidate"])]

# 3. Group per tanggal dan per PT
multi_unit_detail = (
    df_multi_detail.groupby(["realisasidate", "accountname"])
    .agg(
        Jumlah_Produk=("Segmen", "nunique"),
        Jumlah_Realisasi=("NoContract", "count"),
        Produk=("Segmen", lambda x: ", ".join(sorted(x.unique())))
    )
    .reset_index()
    .sort_values(["realisasidate", "Jumlah_Realisasi"], ascending=[True, False])
)

# 4. Tampilkan
st.dataframe(multi_unit_detail, use_container_width=True)

st.markdown("## 📅 Realisasi Multi Unit per Bulan per Perusahaan")

# Tambahkan kolom Tahun-Bulan
df["TahunBulan"] = df["realisasidate"].dt.to_period("M").astype(str)

# Group per bulan dan PT
multi_unit_bulan = (
    df.groupby(["TahunBulan", "accountname"])
    .agg(
        Jumlah_Produk=("Segmen", "nunique"),
        Jumlah_Realisasi=("NoContract", "count"),
        Produk=("Segmen", lambda x: ", ".join(sorted(x.unique())))
    )
    .reset_index()
    .sort_values(["TahunBulan", "Jumlah_Realisasi"], ascending=[True, False])
)

# Tampilkan tabel
st.dataframe(multi_unit_bulan, use_container_width=True)


# ===============================
# DETAIL DATA (OPTIONAL)
# ===============================
with st.expander("Lihat Detail Data"):
    st.dataframe(df, use_container_width=True)
