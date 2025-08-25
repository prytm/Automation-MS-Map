import streamlit as st
import pandas as pd
import io

st.title("Automasi Market Share & Mapping")

# ==== Input Periode Data Bulan Ini ====
with st.expander("Set Periode Data Bulan Ini", expanded=True):
    tahun_input = st.number_input("Tahun", min_value=2000, max_value=2100, step=1)
    bulan_input = st.selectbox("Bulan (1–12)", list(range(1, 13)))

# ==== Helper: pilih sheet ====
def read_sheet_with_picker(uploaded_file, label, default_idx=0):
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    sheet_names = xls.sheet_names
    default = default_idx if default_idx < len(sheet_names) else 0
    sheet_name = st.selectbox(f"Pilih Sheet • {label}", sheet_names, index=default, key=f"{label}_sheet")
    return pd.read_excel(xls, sheet_name=sheet_name)

# ==== Mapping Daerah ke Pulau ====
daerah_to_pulau = {
    # Sumatera
    "D.I. Aceh": "Sumatera", "Sumut": "Sumatera", "Sumbar": "Sumatera",
    "Riau": "Sumatera", "Kepulauan Riau": "Sumatera", "Jambi": "Sumatera",
    "Sumsel": "Sumatera", "Bangka - Belitung": "Sumatera",
    "Bengkulu": "Sumatera", "Lampung": "Sumatera",

    # Jawa
    "D. K. I. Jakarta": "Jawa", "Banten": "Jawa", "Jabar": "Jawa",
    "Jateng": "Jawa", "D. I. Y.": "Jawa", "Jatim": "Jawa",

    # Kalimantan
    "Kalbar": "Kalimantan", "Kalsel": "Kalimantan", "Kalteng": "Kalimantan",
    "Kaltim": "Kalimantan", "Kaltara": "Kalimantan",

    # Sulawesi
    "Sultera": "Sulawesi", "Sulsel": "Sulawesi", "Sulbar": "Sulawesi",
    "Sulteng": "Sulawesi", "Sulut": "Sulawesi", "Gorontalo": "Sulawesi",

    # Bali Nusra
    "Bali": "Bali Nusra", "N. T. B.": "Bali Nusra", "N. T. T.": "Bali Nusra",

    # Indonesia Timur
    "Maluku": "Ind. Timur", "Maluku Utara": "Ind. Timur",
    "Papua Barat": "Ind. Timur", "Papua": "Ind. Timur"
}

# Mapping nbulan → nama bulan Indo
bulan_map = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "Mei", 6: "Jun",
    7: "Jul", 8: "Agt", 9: "Sep", 10: "Okt", 11: "Nov", 12: "Des"
}

BASE_COLS = [
    "Tahun","Bulan","nbulan","Daerah","Pulau","Produsen",
    "Total","Kemasan","Negara","Holding","Merk"
]

def safe_select(df, cols):
    return df[[c for c in cols if c in df.columns]].copy()

def to_numeric_series(s):
    return (
        s.astype(str)
         .str.replace(r"[.,](?=\d{3}\b)", "", regex=True)  # hilangkan pemisah ribuan
         .str.replace("-", "0")
         .replace({"nan": "0", "None": "0"})
         .astype(float)
    )

def calc_ms_and_growth(df):
    # agregasi dasar
    df = (df.groupby(BASE_COLS, as_index=False)["Total"].sum()
            .sort_values(["Tahun","nbulan","Daerah","Merk"]))
    # market share per periode-daerah
    total_per_period = df.groupby(["Tahun","Bulan","Daerah"])["Total"].transform("sum")
    df["MS"] = df["Total"] / total_per_period

    # growth
    df = df.sort_values(["Merk","Daerah","Kemasan","Tahun","nbulan"]).copy()
    df["MoM Growth %"] = df.groupby(["Merk","Daerah","Kemasan"])["MS"].pct_change(1)
    df["YoY Growth %"] = df.groupby(["Merk","Daerah","Kemasan"])["MS"].pct_change(12)

    # YTD share & growth
    df["MS_YTD"] = df.groupby(["Merk","Daerah","Kemasan","Tahun"])["MS"].cumsum()
    df["YtD Growth %"] = df.groupby(["Merk","Daerah","Kemasan"])["MS_YTD"].pct_change(12)

    # MSY (YTD berdasarkan total)
    df = df.sort_values(["Daerah","Merk","Tahun","nbulan"]).copy()
    df["Total Merk YtD"] = df.groupby(["Daerah","Merk","Tahun"])["Total"].cumsum()
    total_all = (
        df.groupby(["Daerah","Tahun","nbulan"])["Total"].sum()
          .groupby(level=["Daerah","Tahun"]).cumsum().reset_index(name="Total All YtD")
    )
    df = df.merge(total_all, on=["Daerah","Tahun","nbulan"], how="left")
    df["MSY"] = df["Total Merk YtD"] / df["Total All YtD"]

    # bersihkan inf
    for col in ["MoM Growth %","YoY Growth %","YtD Growth %"]:
        df[col] = df[col].replace([float("inf"), float("-inf")], 1.0)
    return df

# ==== Uploads ====
uploaded_current = st.file_uploader("Upload Data Bulan Ini (Excel)", type=["xlsx"])
uploaded_db = st.file_uploader("Upload Database (Excel)", type=["xlsx"])
uploaded_map = st.file_uploader("Upload Mapping (Excel)", type=["xlsx"])

if uploaded_current and uploaded_db and uploaded_map:
    # Data Bulan Ini default ke sheet ke-2 (index 1)
    current = read_sheet_with_picker(uploaded_current, "Data Bulan Ini", default_idx=1)
    # Database & Mapping default sheet pertama
    db = read_sheet_with_picker(uploaded_db, "Database", default_idx=0)
    mapping_df = read_sheet_with_picker(uploaded_map, "Mapping", default_idx=0)

    # Terapkan periode & kolom turunan ke seluruh baris Data Bulan Ini
    current["Tahun"] = int(tahun_input)
    current["nbulan"] = int(bulan_input)                 # 1..12 reset tiap tahun
    current["Bulan"] = bulan_map[current["nbulan"]]      # singkatan Indo
    current["Negara"] = "Domestik"                       # force Domestik
    if "Daerah" in current.columns:
        current["Pulau"] = current["Daerah"].map(daerah_to_pulau).fillna("Lainnya")

    # pastikan Total numerik
    if "Total" in current.columns:
        current["Total"] = to_numeric_series(current["Total"])
    if "Total" in db.columns:
        db["Total"] = to_numeric_series(db["Total"])

    # --- MAP HANYA DATA BULAN INI (Segment / Area AP jika ada di mapping) ---
    current_core = safe_select(current, BASE_COLS)

    seg_map = (mapping_df.drop_duplicates(["Merk","Daerah"])[["Merk","Daerah","Segment"]]
               if {"Merk","Daerah","Segment"}.issubset(mapping_df.columns) else None)
    area_map = (mapping_df.drop_duplicates(["Daerah"])[["Daerah","Area AP"]]
               if {"Daerah","Area AP"}.issubset(mapping_df.columns) else None)

    if seg_map is not None:
        current_core = current_core.merge(seg_map, on=["Merk","Daerah"], how="left")
    if area_map is not None:
        current_core = current_core.merge(area_map, on="Daerah", how="left")

    # --- ALIGN KOLUM & APPEND ---
    keep_cols = BASE_C
