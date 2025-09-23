from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from typing import Iterable, Optional, Dict

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
    if x is None or pd.isna(x):
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

def to_int_safe(val: object) -> float | int | np.nan:
    """Parser int aman (ambil digit saja, cocok utk '2x', '3 / bulan', dsb.)."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return np.nan
    try:
        s = re.sub(r"[^\d-]", "", str(val))
        return int(s) if s not in ("", "-") else np.nan
    except Exception:
        return np.nan

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
    df_tampil[col_tgl_input] = df_tampil[col_tgl_input].map(parse_dt_any)
if col_tbd:
    df_tampil[col_tbd] = df_tampil[col_tbd].map(parse_dt_any)

col_harga = find_col_any(df_tampil, ["harga produk"])
if col_harga:
    df_tampil[col_harga] = df_tampil[col_harga].map(to_int_idr)

# Perjanjian Video/Live -> int
col_pj_video = find_col_any(df_tampil, ["perjanjian video"])
col_pj_live  = find_col_any(df_tampil, ["perjanjian live"])
if col_pj_video:
    df_tampil[col_pj_video] = df_tampil[col_pj_video].map(to_int_safe)
if col_pj_live:
    df_tampil[col_pj_live] = df_tampil[col_pj_live].map(to_int_safe)

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
if col_pj_video:
    fmt_map_aff[col_pj_video] = fmt_int_disp
if col_pj_live:
    fmt_map_aff[col_pj_live] = fmt_int_disp

st.subheader("📑 Database Affiliator")
st.dataframe(df_tampil.style.format(fmt_map_aff, na_rep="-"), width="stretch")

# ====== TOTAL PENGELUARAN SAMPLE (berdasarkan filter aktif) ======
if col_harga:
    total_pengeluaran = pd.to_numeric(df_tampil[col_harga], errors="coerce").sum(skipna=True)

    # Opsi A: ringkas & rapi pakai st.metric
    # st.markdown("#### 💸 Total Pengeluaran Sample")
    # st.metric(label="Total Pengeluaran Sample", value=fmt_rp(total_pengeluaran))

    # --- Jika ingin tampilan kartu yang lebih menonjol, pakai Opsi B di bawah dan hapus Opsi A ---
    st.markdown(
        f"""
        <div style="margin-top:0.5rem;padding:16px;border:1px solid #e5e7eb;
                    border-radius:12px;background:#f8fafc">
            <div style="font-size:16px;font-weight:600;margin-bottom:4px;">
                💸 Total Pengeluaran Sample
            </div>
            <div style="font-size:28px;font-weight:800;line-height:1;">
                {fmt_rp(total_pengeluaran)}
            </div>
            <div style="color:#6b7280;font-size:12px;margin-top:4px;">
                Berdasarkan filter aktif pada tabel di atas
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
else:
    st.info("Kolom 'Harga Produk' tidak ditemukan, sehingga total pengeluaran sample tidak dapat dihitung.")


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

            # agregasi database affiliator: jumlah sample & max(tanggal barang diterima) + perjanjian
            col_user_db = find_col_any(df_tampil, ["username"])
            if not col_user_db:
                st.error("Kolom 'Username' pada Database Affiliator tidak ditemukan.")
                st.stop()

            cols_for_grp = [col_user_db]
            if col_tbd:
                cols_for_grp.append(col_tbd)
            if col_pj_video:
                cols_for_grp.append(col_pj_video)
            if col_pj_live:
                cols_for_grp.append(col_pj_live)
            df_grp = df_tampil[cols_for_grp].copy()

            agg_kwargs = {"Jumlah_Sample": (col_user_db, "size")}
            if col_tbd:
                agg_kwargs["Tanggal Barang Diterima"] = (col_tbd, "max")
            if col_pj_video:
                agg_kwargs["Perjanjian Video"] = (col_pj_video, "max")  # bisa diganti "sum"
            if col_pj_live:
                agg_kwargs["Perjanjian Live"] = (col_pj_live, "max")    # bisa diganti "sum"

            agg_db = df_grp.groupby(col_user_db, dropna=False).agg(**agg_kwargs).reset_index()

            # pastikan kolom ada walau tidak ditemukan
            if "Tanggal Barang Diterima" not in agg_db.columns:
                agg_db["Tanggal Barang Diterima"] = pd.NaT
            if "Perjanjian Video" not in agg_db.columns:
                agg_db["Perjanjian Video"] = np.nan
            if "Perjanjian Live" not in agg_db.columns:
                agg_db["Perjanjian Live"] = np.nan

            # join case-insensitive
            agg_db["_key"] = norm_key_series(agg_db[col_user_db])
            agg_sales["_key"] = norm_key_series(agg_sales[col_user_sales])

            hasil = pd.merge(
                agg_db[[col_user_db, "Jumlah_Sample", "Tanggal Barang Diterima", "Perjanjian Video", "Perjanjian Live", "_key"]],
                agg_sales.drop(columns=[col_user_sales]),
                on="_key",
                how="left",
            ).drop(columns=["_key"])

            hasil.rename(columns={col_user_db: "Username"}, inplace=True)

            # pastikan kolom numerik tersedia
            for c in ["GMV", "Items sold", "Videos", "LIVE streams", "Est. commission",
                      "Perjanjian Video", "Perjanjian Live"]:
                if c not in hasil.columns:
                    hasil[c] = np.nan
                hasil[c] = pd.to_numeric(hasil[c], errors="coerce")

            # Reorder kolom agar nyaman dibaca
            desired_order = [
                "Username",
                "Jumlah_Sample",
                "Perjanjian Video", "Videos",
                "Perjanjian Live", "LIVE streams",
                "GMV", "Est. commission", "Items sold",
                "Tanggal Barang Diterima",
            ]
            cols_now = list(hasil.columns)
            ordered = [c for c in desired_order if c in cols_now] + [c for c in cols_now if c not in desired_order]
            hasil = hasil[ordered]

            hasil.index = range(1, len(hasil) + 1)

            # ===== Styling warna target =====
            def _style_target_row(row: pd.Series):
                styles = [""] * len(row)
                cols = list(row.index)

                # Videos vs Perjanjian Video
                if ("Videos" in cols) and ("Perjanjian Video" in cols):
                    idx = cols.index("Videos")
                    v = row.get("Videos", np.nan)
                    t = row.get("Perjanjian Video", np.nan)
                    if pd.notna(v) and pd.notna(t):
                        styles[idx] = "background-color: #d1fae5;" if v >= t else "background-color: #fee2e2;"

                # LIVE streams vs Perjanjian Live
                if ("LIVE streams" in cols) and ("Perjanjian Live" in cols):
                    idx = cols.index("LIVE streams")
                    v = row.get("LIVE streams", np.nan)
                    t = row.get("Perjanjian Live", np.nan)
                    if pd.notna(v) and pd.notna(t):
                        styles[idx] = "background-color: #d1fae5;" if v >= t else "background-color: #fee2e2;"

                return styles

            styler_join = (
                hasil.style
                .format(
                    {
                        "GMV": fmt_rp,
                        "Est. commission": fmt_rp,
                        "Items sold": fmt_int_disp,
                        "Videos": fmt_int_disp,
                        "LIVE streams": fmt_int_disp,
                        "Perjanjian Video": fmt_int_disp,
                        "Perjanjian Live": fmt_int_disp,
                        "Tanggal Barang Diterima": fmt_tgl_id,
                    },
                    na_rep="-",
                )
                .apply(_style_target_row, axis=1)
            )

            # --- Ganti label header untuk tampilan saja ---
            DISP_RENAME = {
                "Videos": "Video dibuat",
                "LIVE streams": "Live dilakukan",
            }

            # DataFrame untuk ditampilkan (header sudah diganti)
            df_display = hasil.rename(columns=DISP_RENAME)

            # Nama kolom setelah di-rename (dipakai untuk format & styling)
            col_vid_disp  = DISP_RENAME.get("Videos", "Videos")
            col_live_disp = DISP_RENAME.get("LIVE streams", "LIVE streams")

            # ===== Styling warna target (pakai nama kolom tampilan) =====
            def _style_target_row_disp(row: pd.Series):
                styles = [""] * len(row)
                cols = list(row.index)

                # ===== Target: Video vs Perjanjian Video =====
                if (col_vid_disp in cols) and ("Perjanjian Video" in cols):
                    idx_v = cols.index(col_vid_disp)
                    v = row.get(col_vid_disp, np.nan)
                    t = row.get("Perjanjian Video", np.nan)
                    if pd.notna(v) and pd.notna(t):
                        styles[idx_v] = "background-color: #d1fae5;" if v >= t else "background-color: #fee2e2;"

                # ===== Target: Live vs Perjanjian Live =====
                if (col_live_disp in cols) and ("Perjanjian Live" in cols):
                    idx_l = cols.index(col_live_disp)
                    v = row.get(col_live_disp, np.nan)
                    t = row.get("Perjanjian Live", np.nan)
                    if pd.notna(v) and pd.notna(t):
                        styles[idx_l] = "background-color: #d1fae5;" if v >= t else "background-color: #fee2e2;"

                # ===== Umur tanggal "Tanggal Barang Diterima" =====
                # >= 30 hari lalu  -> hijau
                # 7..29 hari lalu  -> oranye
                # 0..6 hari lalu   -> merah
                # (opsional) tanggal masa depan / NaT -> dibiarkan default
                if "Tanggal Barang Diterima" in cols:
                    idx_tbd = cols.index("Tanggal Barang Diterima")
                    val = row.get("Tanggal Barang Diterima", pd.NaT)
                    d = pd.to_datetime(val, errors="coerce")
                    if pd.notna(d):
                        today = pd.Timestamp.today().normalize()
                        delta_days = (today - d.normalize()).days
                        if delta_days >= 30:
                            styles[idx_tbd] = "background-color: #d1fae5;"  # hijau lembut
                        elif delta_days >= 7:
                            styles[idx_tbd] = "background-color: #fed7aa;"  # oranye lembut
                        elif delta_days >= 0:
                            styles[idx_tbd] = "background-color: #fecaca;"  # merah lembut
                        # else (tanggal ke depan): biarkan default
                    # NaT: biarkan default (tetap "-")
                return styles


            # ===== Formatter kolom tampilan =====
            fmt_map_join = {
                "GMV": fmt_rp,
                "Est. commission": fmt_rp,
                "Items sold": fmt_int_disp,
                col_vid_disp: fmt_int_disp,
                col_live_disp: fmt_int_disp,
                "Perjanjian Video": fmt_int_disp,
                "Perjanjian Live": fmt_int_disp,
                "Tanggal Barang Diterima": fmt_tgl_id,
            }

            styler_join = (
                df_display.style
                .format(fmt_map_join, na_rep="-")
                .apply(_style_target_row_disp, axis=1)
            )

            st.subheader("💰 Data Penjualan Affiliator")
            st.dataframe(styler_join, width="stretch")

            # -------- 3) Tabel Top GMV (maks. 10 baris) --------
            # Catatan: filter > 0 agar hanya yang benar-benar ada GMV.
            # Jika ingin memasukkan GMV=0, ubah jadi >= 0.
            if "GMV" in df_display.columns:
                mask_gmv = pd.to_numeric(df_display["GMV"], errors="coerce").gt(0)
                top_gmv = (
                    df_display.loc[mask_gmv]
                    .sort_values("GMV", ascending=False)
                    .head(10)
                    .copy()
                )

                if top_gmv.empty:
                    st.info("Belum ada baris dengan GMV > 0 pada hasil gabungan.")
                else:
                    # Index rapi mulai dari 1
                    top_gmv.index = range(1, len(top_gmv) + 1)

                    # Pakai formatter & styling yang sama seperti tabel kedua
                    styler_top = (
                        top_gmv.style
                        .format(fmt_map_join, na_rep="-")
                        .apply(_style_target_row_disp, axis=1)
                    )

                    st.subheader(f"🏆 Top {len(top_gmv)} GMV")
                    st.dataframe(styler_top, width="stretch")
            else:
                st.info("Kolom GMV tidak ditemukan pada hasil gabungan.")

    else:
        st.info("Unggah file penjualan untuk melihat gabungan database + penjualan.")