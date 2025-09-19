# app.py — Ringkasan Laporan Affiliator (versi refactor)
# - Tahan banting: kolom fleksibel, aman dari KeyError/None
# - Tanggal format Indonesia, NA -> "-"
# - GMV & Est. commission -> Rp#, Items/Videos/Live -> int, tanggal -> "D MMMM YYYY"
# - UI filter SKU & Bulan, join database affiliator <-> penjualan

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Iterable, Optional, Dict, List

import numpy as np
import pandas as pd
import streamlit as st

# ===================== Konstanta & Config =====================
st.set_page_config(page_title="Ringkasan Laporan", page_icon="🗂️", layout="wide")

BULAN_ID = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]
BULAN_LOOKUP = {b.lower(): i + 1 for i, b in enumerate(BULAN_ID)}
RE_TGL_TEXT = re.compile(r"^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$")

# ===================== Util Format =====================
def fmt_tgl_id(x: pd.Timestamp | datetime | pd.NaT | None) -> str:
    """Format tanggal ke 'D NamaBulan YYYY' (ID)."""
    if pd.isna(x):
        return "-"
    if not isinstance(x, (pd.Timestamp, datetime)):
        return "-"
    return f"{x.day} {BULAN_ID[x.month - 1]} {x.year}"

def fmt_rp(x: float | int | None) -> str:
    """Format rupiah 'Rp1.234.567' atau '-' jika NA."""
    if x is None or pd.isna(x):
        return "-"
    try:
        return "Rp" + format(int(x), ",").replace(",", ".")
    except Exception:
        return "-"

def fmt_int_disp(x: float | int | None) -> str:
    """Tampilkan bilangan bulat tanpa desimal, '-' jika NA."""
    if x is None or pd.isna(x):
        return "-"
    try:
        return f"{int(x)}"
    except Exception:
        return "-"

# ===================== Util Parsers =====================
def parse_dt_any(s: object) -> pd.Timestamp | pd.NaT:
    """
    Parsing tanggal fleksibel:
    - ISO/Excel umum: 'YYYY-MM-DD', 'YYYY-MM-DD HH:MM:SS'
    - ID umum: 'DD/MM/YYYY', 'DD-MM-YYYY'
    - Teks ID: '12 Agustus 2025'
    """
    if s is None:
        return pd.NaT
    v = str(s).strip()
    if v == "":
        return pd.NaT

    # coba pandas
    dt = pd.to_datetime(v, errors="coerce", dayfirst=False)
    if not pd.isna(dt):
        return dt

    # pattern manual
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return pd.Timestamp(datetime.strptime(v, fmt))
        except Exception:
            pass

    m = RE_TGL_TEXT.match(v)
    if m:
        d, mm, y = m.groups()
        bln = BULAN_LOOKUP.get(mm.lower())
        if bln:
            try:
                return pd.Timestamp(datetime(int(y), bln, int(d)))
            except Exception:
                return pd.NaT

    return pd.NaT

def to_int_idr(val: object) -> float | int | np.nan:
    """
    Konversi string rupiah ke integer:
      - "Rp6,1JT" -> 6100000
      - "Rp79,2RB" -> 79200
      - "Rp43.900" -> 43900
      - "43900" -> 43900
    NA -> np.nan
    """
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return np.nan

    s = str(val).strip().upper()
    if not s:
        return np.nan

    # normalisasi angka
    if "JT" in s:
        num = s.replace("RP", "").replace("JT", "").replace(".", "").replace(" ", "").replace(",", ".")
        try:
            return int(round(float(num) * 1_000_000))
        except Exception:
            return np.nan
    if "RB" in s:
        num = s.replace("RP", "").replace("RB", "").replace(".", "").replace(" ", "").replace(",", ".")
        try:
            return int(round(float(num) * 1_000))
        except Exception:
            return np.nan

    digits = re.sub(r"[^\d]", "", s)
    return int(digits) if digits else np.nan

# ===================== Util Kolom =====================
def find_col_any(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    """
    Cari nama kolom asli berdasarkan kandidat (case-insensitive, trim),
    cocok eksak lebih dulu, lalu "contains".
    Return None jika tak ketemu.
    """
    if df is None or df.empty or not list(df.columns):
        return None

    lower_map = {c: c.lower().strip() for c in df.columns}
    wants = [w.lower().strip() for w in candidates]

    # eksak
    for want in wants:
        for orig, low in lower_map.items():
            if low == want:
                return orig

    # contains
    for want in wants:
        for orig, low in lower_map.items():
            if want in low:
                return orig

    return None

def norm_key_series(s: pd.Series) -> pd.Series:
    """Normalisasi username untuk join-insensitif: lower, trim, NaN -> ''."""
    return s.fillna("").astype(str).str.strip().str.lower()

# ===================== Data Loading =====================
@dataclass
class LoadResult:
    df: pd.DataFrame
    sheet_name: str

def load_excel_any(upload, label_singular: str) -> Optional[LoadResult]:
    """Baca Excel (semua sheet -> pilih), return DataFrame & nama sheet, atau None jika gagal."""
    if not upload:
        return None
    try:
        xls = pd.read_excel(upload, sheet_name=None, dtype=str)
    except Exception as e:
        st.error(f"Gagal membaca {label_singular}: {e}")
        return None

    if not xls:
        st.warning(f"Tidak ada sheet pada {label_singular}.")
        return None

    names = list(xls.keys())
    chosen = names[0] if len(names) == 1 else st.selectbox(f"Pilih Sheet {label_singular}", names)
    df = xls[chosen].copy()
    if df.empty:
        st.warning(f"Sheet '{chosen}' kosong pada {label_singular}.")
    return LoadResult(df=df, sheet_name=chosen)

# ===================== UI & Alur =====================
st.title("🗂️ Ringkasan Laporan Affiliator")

# -------- 1) Unggah Database Affiliator --------
up1 = st.file_uploader("Unggah Database Affiliator (Excel)", type=["xlsx", "xls"], label_visibility="visible")
db = load_excel_any(up1, "Database Affiliator")
if db is None or db.df.empty:
    st.stop()

df_db = db.df

# Pilih kolom target minimal
TARGETS = [
    "tanggal input data", "username", "perjanjian video", "perjanjian live",
    "sku utama", "variasi", "harga produk", "tanggal barang diterima",
]
col_map = {c: c.lower().strip() for c in df_db.columns}
kolom_asli = [col for col, low in col_map.items() if low in TARGETS]
if not kolom_asli:
    st.error("Tidak menemukan kolom target minimum pada Database Affiliator.")
    st.stop()

df_tampil = df_db[kolom_asli].copy()

# Parsing tanggal & harga jika ada
col_tgl_input = find_col_any(df_tampil, ["tanggal input data"])
col_tbd = find_col_any(df_tampil, ["tanggal barang diterima"])
if col_tgl_input:
    df_tampil[col_tgl_input] = pd.to_datetime(df_tampil[col_tgl_input].map(parse_dt_any))
if col_tbd:
    df_tampil[col_tbd] = pd.to_datetime(df_tampil[col_tbd].map(parse_dt_any))

col_harga = find_col_any(df_tampil, ["harga produk"])
if col_harga:
    df_tampil[col_harga] = df_tampil[col_harga].map(to_int_idr)

# Filter SKU
col_sku = find_col_any(df_tampil, ["sku utama"])
if col_sku:
    sku_options = sorted([x for x in df_tampil[col_sku].dropna().astype(str).unique()])
    pilihan_sku = st.multiselect("Filter SKU Utama", options=sku_options)
    if pilihan_sku:
        df_tampil = df_tampil[df_tampil[col_sku].isin(pilihan_sku)]

# Filter Bulan (berdasarkan Tanggal Input Data)
if col_tgl_input:
    col_period = df_tampil[col_tgl_input].dropna()
    if not col_period.empty:
        periods = col_period.dt.to_period("M").sort_values().unique()
        label_bulan = [f"{BULAN_ID[p.month - 1]} {p.year}" for p in periods]
        pilihan_bulan = st.multiselect("Filter Bulan (Tanggal Input Data)", options=label_bulan)
        if pilihan_bulan:
            mask = df_tampil[col_tgl_input].apply(
                lambda d: f"{BULAN_ID[d.month - 1]} {d.year}" if pd.notna(d) else None
            ).isin(pilihan_bulan)
            df_tampil = df_tampil[mask.fillna(False)]

# Renumber setelah filter
df_tampil.index = range(1, len(df_tampil) + 1)

# Tampilkan Database Affiliator
fmt_map_aff: Dict[str, callable] = {}
if col_tgl_input:
    fmt_map_aff[col_tgl_input] = fmt_tgl_id
if col_tbd:
    fmt_map_aff[col_tbd] = fmt_tgl_id
if col_harga:
    fmt_map_aff[col_harga] = fmt_rp

st.subheader("📑 Database Affiliator")
st.dataframe(df_tampil.style.format(fmt_map_aff, na_rep="-"), width="stretch")

# -------- 2) Unggah Data Penjualan & Join --------
up2 = st.file_uploader("Unggah Data Penjualan Affiliator (Excel)", type=["xlsx", "xls"], label_visibility="visible")
if up2:
    sales = load_excel_any(up2, "Data Penjualan")
    if sales and not sales.df.empty:
        df_sales_raw = sales.df

        col_user_sales = find_col_any(
            df_sales_raw,
            ["creator username", "username", "creator_username", "creator user", "creator"],
        )
        col_gmv = find_col_any(df_sales_raw, ["gmv"])
        col_items = find_col_any(df_sales_raw, ["items sold", "items_sold", "items"])
        col_vid = find_col_any(df_sales_raw, ["videos", "video"])
        col_live = find_col_any(df_sales_raw, ["live streams", "live_streams", "live"])
        col_comm = find_col_any(df_sales_raw, ["est. commission", "est commission", "commission"])

        if not col_user_sales:
            st.error("Kolom username pada Data Penjualan tidak ditemukan (cari 'Creator Username').")
        else:
            # siapkan df numerik untuk agregasi
            df_sales_num = pd.DataFrame()
            df_sales_num[col_user_sales] = df_sales_raw[col_user_sales].astype(str)

            if col_gmv:
                df_sales_num["GMV"] = df_sales_raw[col_gmv].map(to_int_idr)
            if col_comm:
                df_sales_num["Est. commission"] = df_sales_raw[col_comm].map(to_int_idr)
            if col_items:
                df_sales_num["Items sold"] = pd.to_numeric(df_sales_raw[col_items], errors="coerce")
            if col_vid:
                df_sales_num["Videos"] = pd.to_numeric(df_sales_raw[col_vid], errors="coerce")
            if col_live:
                df_sales_num["LIVE streams"] = pd.to_numeric(df_sales_raw[col_live], errors="coerce")

            # agregasi sales per username
            agg_sales = (
                df_sales_num.groupby(col_user_sales, dropna=False)
                .sum(numeric_only=True)
                .reset_index()
            )

            # agregasi database affiliator: jumlah sample & max(tanggal barang diterima)
            col_user_db = find_col_any(df_tampil, ["username"])
            if not col_user_db:
                st.error("Kolom 'Username' pada Database Affiliator tidak ditemukan.")
                st.stop()

            cols_for_grp = [col_user_db] + ([col_tbd] if col_tbd else [])
            df_grp = df_tampil[cols_for_grp].copy()

            agg_kwargs = {"Jumlah_Sample": (col_user_db, "size")}
            if col_tbd:
                agg_kwargs["Tanggal Barang Diterima"] = (col_tbd, "max")

            agg_db = df_grp.groupby(col_user_db, dropna=False).agg(**agg_kwargs).reset_index()
            if "Tanggal Barang Diterima" not in agg_db.columns:
                agg_db["Tanggal Barang Diterima"] = pd.NaT

            # join case-insensitive
            agg_db["_key"] = norm_key_series(agg_db[col_user_db])
            agg_sales["_key"] = norm_key_series(agg_sales[col_user_sales])

            hasil = pd.merge(
                agg_db[[col_user_db, "Jumlah_Sample", "Tanggal Barang Diterima", "_key"]],
                agg_sales.drop(columns=[col_user_sales]),
                on="_key",
                how="left",
            ).drop(columns=["_key"])

            hasil.rename(columns={col_user_db: "Username"}, inplace=True)

            # pastikan kolom numerik tersedia
            for c in ["GMV", "Items sold", "Videos", "LIVE streams", "Est. commission"]:
                if c not in hasil.columns:
                    hasil[c] = np.nan
                hasil[c] = pd.to_numeric(hasil[c], errors="coerce")

            hasil.index = range(1, len(hasil) + 1)

            styler_join = hasil.style.format(
                {
                    "GMV": fmt_rp,
                    "Est. commission": fmt_rp,
                    "Items sold": fmt_int_disp,
                    "Videos": fmt_int_disp,
                    "LIVE streams": fmt_int_disp,
                    "Tanggal Barang Diterima": fmt_tgl_id,
                },
                na_rep="-",
            )

            st.subheader("💰 Data Penjualan Affiliator")
            st.dataframe(styler_join, width="stretch")
    else:
        st.info("Unggah file penjualan untuk melihat gabungan database + penjualan.")