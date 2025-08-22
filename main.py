import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Automasi Market Share & Mapping", layout="wide")
st.title("Automasi Market Share & Mapping")

# ============== Helpers ==============
BULAN2NUM = {
    "Januari": 1, "Februari": 2, "Maret": 3, "April": 4, "Mei": 5, "Juni": 6,
    "Juli": 7, "Agustus": 8, "September": 9, "Oktober": 10, "November": 11, "Desember": 12
}

# Kolom inti yang kita pakai di seluruh perhitungan
BASE_COLS = [
    "Tahun", "Bulan", "nbulan", "Daerah", "Pulau", "Produsen",
    "Total", "Kemasan", "Negara", "Holding", "Merk"
]

def ensure_nbulan(df):
    """Pastikan kolom nbulan ada (1..12). Kalau belum ada, bikin dari Bulan."""
    if "nbulan" not in df.columns:
        if "Bulan" in df.columns:
            df["nbulan"] = df["Bulan"].map(BULAN2NUM).fillna(df["Bulan"])
        else:
            st.error("Kolom 'nbulan' atau 'Bulan' wajib ada.")
            st.stop()
    return df

def safe_select(df, cols):
    """Ambil kolom yang ada saja agar aman saat align."""
    return df[[c for c in cols if c in df.columns]].copy()

def to_numeric_series(s):
    """Koersi angka dengan aman (hapus pemisah ribuan, ubah '-' ke 0)."""
    return (
        s.astype(str)
         .str.replace(r"[.,](?=\d{3}\b)", "", regex=True)  # hapus thousand sep sederhana
         .str.replace("-", "0")
         .replace({"nan": "0", "None": "0"})
         .astype(float)
    )

def calc_ms_and_growth(df):
    """Hitung MS, Growth, YTD, dan MSY. df harus sudah bersih & terurut."""
    # agregasi per baris unik (kalau ada duplikat input)
    df = (df.groupby(BASE_COLS, as_index=False)["Total"].sum()
            .sort_values(["Tahun", "nbulan", "Daerah", "Merk"]).reset_index(drop=True))

    # Market Share per (Tahun,Bulan,Daerah)
    total_per_period = df.groupby(["Tahun", "Bulan", "Daerah"])["Total"].transform("sum")
    df["MS"] = df["Total"] / total_per_period

    # Growth berbasis MS
    df["MoM Growth %"] = (
        df.sort_values(["Merk", "Daerah", "Kemasan", "Tahun", "nbulan"])
          .groupby(["Merk", "Daerah", "Kemasan"])["MS"]
          .pct_change(1)
    )
    df["YoY Growth %"] = (
        df.sort_values(["Merk", "Daerah", "Kemasan", "Tahun", "nbulan"])
          .groupby(["Merk", "Daerah", "Kemasan"])["MS"]
          .pct_change(12)
    )

    # MS_YTD (akumulasi MS per tahun) + YtD Growth YoY
    df["MS_YTD"] = (
        df.sort_values(["Merk", "Daerah", "Kemasan", "Tahun", "nbulan"])
          .groupby(["Merk", "Daerah", "Kemasan", "Tahun"])["MS"]
          .cumsum()
    )
    df["YtD Growth %"] = (
        df.sort_values(["Merk", "Daerah", "Kemasan", "Tahun", "nbulan"])
          .groupby(["Merk", "Daerah", "Kemasan"])["MS_YTD"]
          .pct_change(12)
    )

    # MSY (market share YTD by Total)
    df = df.sort_values(["Daerah", "Merk", "Tahun", "nbulan"]).copy()
    df["Total Merk YtD"] = (
        df.groupby(["Daerah", "Merk", "Tahun"])["Total"].cumsum()
    )
    total_all = (
        df.groupby(["Daerah", "Tahun", "nbulan"])["Total"]
          .sum()
          .groupby(level=["Daerah", "Tahun"])
          .cumsum()
          .reset_index(name="Total All YtD")
    )
    df = df.merge(total_all, on=["Daerah", "Tahun", "nbulan"], how="left")
    df["MSY"] = df["Total Merk YtD"] / df["Total All YtD"]

    # Bersihkan inf/-inf → 1.0 (100%) sesuai preferensi kamu)
    for col in ["MoM Growth %", "YoY Growth %", "YtD Growth %"]:
        df[col] = df[col].replace([float("inf"), float("-inf")], 1.0)

    return df

# ============== Inputs ==============
uploaded_current = st.file_uploader("Upload **Data Bulan Ini** (Excel)", type=["xlsx"])
uploaded_db = st.file_uploader("Upload **Database (semua bulan)** (Excel)", type=["xlsx"])
uploaded_map = st.file_uploader("Upload **Mapping (Excel)**", type=["xlsx"])

if uploaded_current and uploaded_db and uploaded_map:
    # --- Read files ---
    current = pd.read_excel(uploaded_current)
    db = pd.read_excel(uploaded_db)
    mapping_df = pd.read_excel(uploaded_map)

    # --- Bersihkan & selaraskan kolom inti ---
    # pastikan nbulan ada
    current = ensure_nbulan(current)
    db = ensure_nbulan(db)

    # pastikan Total numerik
    if "Total" in current.columns:
        current["Total"] = to_numeric_series(current["Total"])
    if "Total" in db.columns:
        db["Total"] = to_numeric_series(db["Total"])

    # Ambil hanya kolom inti yang tersedia
    current_core = safe_select(current, BASE_COLS)
    db_core = safe_select(db, BASE_COLS)

    # --- Mapping ke data BULAN INI dulu ---
    # (mapping minimal: Merk+Daerah -> Segment, Daerah -> Area AP)
    # kalau di database historis sudah ada Segment/Area AP, nanti tetap ikut saat concat bila kolomnya ada
    if {"Merk", "Daerah"}.issubset(mapping_df.columns):
        seg_map = mapping_df.drop_duplicates(subset=["Merk", "Daerah"])[["Merk", "Daerah", "Segment"]] \
                            if "Segment" in mapping_df.columns else None
    else:
        seg_map = None

    area_map = mapping_df.drop_duplicates(subset=["Daerah"])[["Daerah", "Area AP"]] \
                if "Area AP" in mapping_df.columns else None

    if seg_map is not None:
        current_core = current_core.merge(seg_map, on=["Merk", "Daerah"], how="left")
    if area_map is not None:
        current_core = current_core.merge(area_map, on="Daerah", how="left")

    # --- Align kolom yang sama untuk append ---
    # kita ambil irisan kolom supaya aman (hanya kolom yang match)
    common_cols = list(set(current_core.columns) & set(db_core.columns))
    # kolom tambahan hasil mapping (Segment, Area AP) kalau ada
    extra_cols = [c for c in ["Segment", "Area AP"] if c in current_core.columns]
    keep_cols = sorted(set(common_cols + extra_cols), key=lambda x: (x not in BASE_COLS, x))

    current_aligned = safe_select(current_core, keep_cols)
    db_aligned = safe_select(db_core, keep_cols)

    # pastikan kolom extra di db ada juga (isi NaN) agar concat tidak error
    for col in keep_cols:
        if col not in db_aligned.columns:
            db_aligned[col] = pd.NA
    db_aligned = db_aligned[keep_cols]

    # --- Gabung (append) ---
    combined = pd.concat([db_aligned, current_aligned], ignore_index=True)

    # --- Perhitungan MS & Growth di GABUNGAN ---
    combined = ensure_nbulan(combined)
    result = calc_ms_and_growth(combined)

    # --- Urutan & tampilan akhir ---
    final_cols = BASE_COLS + [c for c in ["Segment", "Area AP"] if c in result.columns] + \
                 ["MS", "MoM Growth %", "YoY Growth %", "YtD Growth %",
                  "Total Merk YtD", "Total All YtD", "MSY"]

    final_cols = [c for c in final_cols if c in result.columns]
    final = (result[final_cols]
             .sort_values(["Tahun", "nbulan", "Merk"])
             .reset_index(drop=True))
    

    st.success(f"Berhasil! Baris: {len(final):,}")

    # --- Download ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final.to_excel(writer, index=False, sheet_name="Result")
    output.seek(0)

    st.download_button(
        "Download Data Hasil",
        data=output,
        file_name="Data_Hasil.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Opsional: preview 50 baris
    st.dataframe(final.head(50))
else:
    st.info("Upload **tiga file**: Data Bulan Ini, Database, dan Mapping.")
