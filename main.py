# streamlit_app.py
import streamlit as st
import pandas as pd
import numpy as np
import re
import io

# =========================
# KONFIGURASI HEADER & DATA
# =========================
# (index pandas 0-based; Excel rows: 6,7,52,53,8)
ROW_PRODUSEN   = 6  - 1   # row 6 -> index 5
ROW_KEMASAN    = 7  - 1
ROW_MERK       = 52 - 1
ROW_HOLDING    = 53 - 1
ROW_DATA_START = 8  - 1

BASE_COLS = [
    "Tahun","Bulan","nbulan","Daerah","Pulau","Produsen",
    "Total","Kemasan","Negara","Holding","Merk"
]

# Mapping pulau & nama bulan
daerah_to_pulau = {
    "D.I. Aceh":"Sumatera","Sumut":"Sumatera","Sumbar":"Sumatera","Riau":"Sumatera",
    "Kepulauan Riau":"Sumatera","Jambi":"Sumatera","Sumsel":"Sumatera",
    "Bangka - Belitung":"Sumatera","Bengkulu":"Sumatera","Lampung":"Sumatera",
    "D. K. I. Jakarta":"Jawa","Banten":"Jawa","Jabar":"Jawa","Jateng":"Jawa","D. I. Y.":"Jawa","Jatim":"Jawa",
    "Kalbar":"Kalimantan","Kalsel":"Kalimantan","Kalteng":"Kalimantan","Kaltim":"Kalimantan","Kaltara":"Kalimantan",
    "Sultera":"Sulawesi","Sulsel":"Sulawesi","Sulbar":"Sulawesi","Sulteng":"Sulawesi","Sulut":"Sulawesi","Gorontalo":"Sulawesi",
    "Bali":"Bali Nusra","N. T. B.":"Bali Nusra","N. T. T.":"Bali Nusra",
    "Maluku":"Ind. Timur","Maluku Utara":"Ind. Timur","Papua Barat":"Ind. Timur","Papua":"Ind. Timur"
}
bulan_map = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"Mei",6:"Jun",7:"Jul",8:"Agt",9:"Sep",10:"Okt",11:"Nov",12:"Des"}

# ==========
# UTILITIES
# ==========
def header_text(df: pd.DataFrame, r: int, c: int) -> str:
    """Ambil teks grid; NaN => ''."""
    try:
        v = df.iat[r, c]
    except Exception:
        return ""
    return "" if pd.isna(v) else str(v)

def clean_text(s: str) -> str:
    return str(s).strip()

def clean_kemasan(s: str) -> str:
    s = clean_text(s)
    return "Bulk" if s.lower() == "curah" else s

def to_number(v) -> float:
    if pd.isna(v): return 0.0
    s = str(v).strip()
    if s in ("", "-"): return 0.0
    # buang pemisah ribuan umum
    s = re.sub(r"[.,](?=\d{3}\b)", "", s)
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        try:
            return float(re.sub(r"[^\d.-]", "", s))
        except Exception:
            return 0.0

def stop_at_this_column(df: pd.DataFrame, col: int) -> bool:
    """True jika sel data pertama (ROW_DATA_START) kosong / '-'."""
    v = header_text(df, ROW_DATA_START, col)
    return (v.strip() == "" or v.strip() == "-")

def find_col_provinsi(df: pd.DataFrame, max_col: int):
    """Cari kolom 'Provinsi' fleksibel di row 6/7/52."""
    for c in range(0, max_col+1):
        t6  = header_text(df, ROW_PRODUSEN, c).replace(" ", "").upper()
        t7  = header_text(df, ROW_KEMASAN,  c).replace(" ", "").upper()
        t52 = header_text(df, ROW_MERK,     c).replace(" ", "").upper()
        if "PROVINSI" in t6 or "PROVINSI" in t7 or "PROVINSI" in t52:
            return c
    return None

def to_numeric_series(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
         .str.replace(r"[.,](?=\d{3}\b)", "", regex=True)
         .str.replace("-", "0")
         .replace({"nan":"0","None":"0"})
         .astype(float)
    )

def safe_select(df: pd.DataFrame, cols: list) -> pd.DataFrame:
    return df[[c for c in cols if c in df.columns]].copy()

# ==============================
# UNPIVOT: PRODUSEN-HOLDING-MERK
# ==============================
def unpivot_produsen_holding_merk(xlsx_bytes: bytes, sheet_name=0) -> pd.DataFrame:
    """
    Return DataFrame long-format:
    ['Daerah','Kemasan','Produsen','Holding','Merk','Total','OrderKey']
    """
    # baca grid mentah tanpa header
    df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet_name, header=None, engine="openpyxl", dtype=str)
    if df.shape[1] == 0:
        raise ValueError("Sheet kosong.")

    # batas kolom → pakai baris kemasan
    max_col = int(df.iloc[ROW_KEMASAN].last_valid_index()) if df.shape[1] else -1
    if max_col < 0:
        raise ValueError("Baris kemasan (row 7) kosong / tidak ditemukan.")

    col_prov = find_col_provinsi(df, max_col)
    if col_prov is None:
        raise ValueError("Kolom 'Provinsi' tidak ditemukan di row 6/7/52.")

    first_data_col = col_prov + 1

    # Urutan produsen untuk OrderKey (kiri->kanan)
    produsen_order = []
    for c in range(first_data_col, max_col + 1):
        if stop_at_this_column(df, c):
            break
        produsen = clean_text(header_text(df, ROW_PRODUSEN, c))
        kemasan  = clean_kemasan(header_text(df, ROW_KEMASAN,  c))
        if produsen and kemasan in ("Bag", "Bulk"):
            if produsen not in produsen_order:
                produsen_order.append(produsen)
    produsen_to_idx = {p: i+1 for i, p in enumerate(produsen_order)}

    # Build records: Bag dulu, Bulk belakangan
    records = []
    for pass_type in ("Bag", "Bulk"):
        blank_run = 0
        r = ROW_DATA_START
        max_row = df.shape[0] - 1
        while r <= max_row:
            daerah = clean_text(header_text(df, r, col_prov))
            # Stop bawah:
            if daerah.upper().startswith("CATATAN"):
                break
            if daerah == "":
                blank_run += 1
                if blank_run >= 2:
                    break
                r += 1
                continue
            else:
                blank_run = 0

            if daerah.upper().startswith("TOTAL"):
                r += 1
                continue

            # Loop kolom data
            for c in range(first_data_col, max_col + 1):
                if stop_at_this_column(df, c):
                    break

                merk    = clean_text(header_text(df, ROW_MERK,     c))
                prod    = clean_text(header_text(df, ROW_PRODUSEN,  c))
                holding = clean_text(header_text(df, ROW_HOLDING,   c))
                kemasan = clean_kemasan(header_text(df, ROW_KEMASAN, c))

                if prod and kemasan in ("Bag", "Bulk") and kemasan == pass_type:
                    val = to_number(header_text(df, r, c))
                    type_rank = 0 if pass_type == "Bag" else 100
                    if prod in produsen_to_idx:
                        order_key = type_rank + produsen_to_idx[prod]
                        records.append({
                            "Daerah": daerah,
                            "Kemasan": kemasan,
                            "Produsen": prod,
                            "Holding": holding,
                            "Merk": merk,
                            "Total": val,
                            "OrderKey": order_key
                        })
            r += 1

    out = pd.DataFrame.from_records(records)
    if not out.empty:
        out = out.sort_values(["Daerah", "OrderKey"], kind="mergesort").reset_index(drop=True)
    return out

# ========================
# HITUNG MS & PERTUMBUHAN
# ========================
def calc_ms_and_growth(df: pd.DataFrame) -> pd.DataFrame:
    df = (df.groupby(BASE_COLS, as_index=False)["Total"].sum()
            .sort_values(["Tahun","nbulan","Daerah","Merk"]))
    total_per_period = df.groupby(["Tahun","Bulan","Daerah"])["Total"].transform("sum")
    df["MS"] = df["Total"] / total_per_period

    df = df.sort_values(["Merk","Daerah","Kemasan","Tahun","nbulan"]).copy()
    df["MoM Growth %"] = df.groupby(["Merk","Daerah","Kemasan"])["MS"].pct_change(1)
    df["YoY Growth %"] = df.groupby(["Merk","Daerah","Kemasan"])["MS"].pct_change(12)

    df["MS_YTD"] = df.groupby(["Merk","Daerah","Kemasan","Tahun"])["MS"].cumsum()
    df["YtD Growth %"] = df.groupby(["Merk","Daerah","Kemasan"])["MS_YTD"].pct_change(12)

    # MSY (market share year-to-date berbasis Total)
    df = df.sort_values(["Daerah","Merk","Tahun","nbulan"]).copy()
    df["Total Merk YtD"] = df.groupby(["Daerah","Merk","Tahun"])["Total"].cumsum()
    total_all = (
        df.groupby(["Daerah","Tahun","nbulan"])["Total"].sum()
          .groupby(level=["Daerah","Tahun"]).cumsum().reset_index(name="Total All YtD")
    )
    df = df.merge(total_all, on=["Daerah","Tahun","nbulan"], how="left")
    df["MSY"] = df["Total Merk YtD"] / df["Total All YtD"]

    for col in ["MoM Growth %","YoY Growth %","YtD Growth %"]:
        df[col] = df[col].replace([float("inf"), float("-inf")], 1.0)
    return df

# ======================
# STREAMLIT: APP LAYOUT
# ======================
st.set_page_config(page_title="Automasi Market Share & Mapping", layout="wide")
st.title("Automasi Market Share & Mapping")

# Periode
with st.expander("Set Periode Data Bulan Ini", expanded=True):
    tahun_input = st.number_input("Tahun", min_value=2000, max_value=2100, step=1, value=2025)
    bulan_input = st.selectbox("Bulan (1–12)", list(range(1, 13)))

# Uploads
uploaded_current = st.file_uploader("Upload Data Bulan Ini (Excel)", type=["xlsx"])
uploaded_db      = st.file_uploader("Upload Database (Excel)", type=["xlsx"])
uploaded_map     = st.file_uploader("Upload Mapping (Excel)", type=["xlsx"])

# Helper: dapatkan bytes dari UploadedFile (agar bisa dibaca berulang)
def get_bytes(uploaded_file) -> bytes:
    return uploaded_file.getvalue() if uploaded_file is not None else None

# Pilih sheet untuk Data Bulan Ini (opsional)
df_long = None
if uploaded_current is not None:
    try:
        cur_bytes = get_bytes(uploaded_current)
        xls = pd.ExcelFile(io.BytesIO(cur_bytes), engine="openpyxl")
        sheet_names = xls.sheet_names
        sheet_sel = st.selectbox("Pilih Sheet • Data Bulan Ini", sheet_names, index=0)
        df_long = unpivot_produsen_holding_merk(cur_bytes, sheet_name=sheet_sel)
        st.success(f"Unpivot OK • Baris: {len(df_long):,}")
        st.dataframe(df_long.head(30), use_container_width=True)
    except Exception as e:
        st.error(f"Gagal unpivot Data Bulan Ini: {e}")

# Tombol start aktif kalau tiga file ada & df_long sudah berhasil
start = st.button(
    "Start Proses",
    type="primary",
    disabled=not (df_long is not None and uploaded_db is not None and uploaded_map is not None)
)

if not (uploaded_current and uploaded_db and uploaded_map):
    st.info("Upload tiga file: Data Bulan Ini, Database, dan Mapping.")

if start:
    try:
        # Baca DB & Mapping (sheet pertama, atau bisa kasih selector seperti Current)
        db_bytes  = get_bytes(uploaded_db)
        map_bytes = get_bytes(uploaded_map)
        db = pd.read_excel(io.BytesIO(db_bytes), engine="openpyxl")
        mapping_df = pd.read_excel(io.BytesIO(map_bytes), engine="openpyxl")

        current = df_long.copy()

        # Set periode & kolom turunan Data Bulan Ini
        current["Tahun"]  = int(tahun_input)
        current["nbulan"] = int(bulan_input)
        current["Bulan"]  = current["nbulan"].astype(int).map(bulan_map)
        current["Negara"] = "Domestik"
        if "Daerah" in current.columns:
            current["Pulau"] = current["Daerah"].map(daerah_to_pulau).fillna("Lainnya")

        # Pastikan Total numerik
        if "Total" in current.columns:
            current["Total"] = to_numeric_series(current["Total"])
        if "Total" in db.columns:
            db["Total"] = to_numeric_series(db["Total"])

        # Map Segment / Area AP bila tersedia di mapping_df
        current_core = safe_select(current, BASE_COLS)
        if {"Merk","Daerah","Segment"}.issubset(mapping_df.columns):
            seg_map = mapping_df.drop_duplicates(["Merk","Daerah"])[["Merk","Daerah","Segment"]]
            current_core = current_core.merge(seg_map, on=["Merk","Daerah"], how="left")
        if {"Daerah","Area AP"}.issubset(mapping_df.columns):
            area_map = mapping_df.drop_duplicates(["Daerah"])[["Daerah","Area AP"]]
            current_core = current_core.merge(area_map, on="Daerah", how="left")

        # Samakan kolom & append
        keep_cols = BASE_COLS + [c for c in ["Segment","Area AP"]
                                 if c in (set(db.columns) | set(current_core.columns))]
        keep_cols = [c for c in keep_cols if c in (set(db.columns) | set(current_core.columns))]
        db_aligned = safe_select(db, keep_cols)
        current_aligned = safe_select(current_core, keep_cols)

        # Optional: replace mode (hapus record periode yang sama di DB)
        if {"Tahun","nbulan"}.issubset(current_aligned.columns):
            y_now = int(current_aligned["Tahun"].max())
            m_now = current_aligned.loc[current_aligned["Tahun"].eq(y_now), "nbulan"].max()
            db_clean = db_aligned[~((db_aligned["Tahun"] == y_now) & (db_aligned["nbulan"] == m_now))]
        else:
            db_clean = db_aligned

        combined = pd.concat([db_clean, current_aligned], ignore_index=True)

        # Hitung MS & Growth
        result = calc_ms_and_growth(combined)

        final_cols = keep_cols + ["MS","MoM Growth %","YoY Growth %","YtD Growth %",
                                  "Total Merk YtD","Total All YtD","MSY"]
        final_cols = [c for c in final_cols if c in result.columns]
        final = (result[final_cols]
                 .sort_values(["Tahun","nbulan","Merk"])
                 .reset_index(drop=True))

        # Optional key gabungan (kalau perlu)
        final["X"] = (
            final["Tahun"].astype(str)
            + final["Bulan"].astype(str)
            + final["Daerah"].astype(str)
            + final["Merk"].astype(str)
            + final["Kemasan"].astype(str)
        )

        # Urut kolom output
        desired = ["X","Tahun","Bulan","Daerah","Pulau","Produsen","Total",
                   "Kemasan","Negara","Holding","Merk","nbulan",
                   "MS","MoM Growth %","YoY Growth %","YtD Growth %",
                   "Total Merk YtD","Total All YtD","MSY"]
        final = safe_select(final, desired)

        st.success(f"Selesai! Baris hasil: {len(final):,}")
        st.dataframe(final.head(50), use_container_width=True)

        # Download Excel
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            final.to_excel(w, index=False, sheet_name="Result")
        buf.seek(0)
        st.download_button(
            "Download Data_Hasil.xlsx", buf,
            file_name="Data_Hasil.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Gagal memproses: {e}")
