import streamlit as st
import pandas as pd
import io

st.title("Automasi Market Share & Mapping")

# ==== Input Tahun & Bulan untuk Data Bulan Ini 
tahun_input = st.number_input("Masukkan Tahun Data Bulan Ini", min_value=2000, max_value=2100, step=1)
bulan_input = st.selectbox("Pilih Bulan Data Bulan Ini", list(range(1,13)), format_func=lambda x: f"{x:02d}")

BASE_COLS = [
    "Tahun","Bulan","nbulan","Daerah","Pulau","Produsen",
    "Total","Kemasan","Negara","Holding","Merk"
]

def safe_select(df, cols):
    return df[[c for c in cols if c in df.columns]].copy()

def to_numeric_series(s):
    return (
        s.astype(str)
         .str.replace(r"[.,](?=\d{3}\b)", "", regex=True)
         .str.replace("-", "0")
         .replace({"nan": "0", "None": "0"})
         .astype(float)
    )

def calc_ms_and_growth(df):
    df = (df.groupby(BASE_COLS, as_index=False)["Total"].sum()
            .sort_values(["Tahun","nbulan","Daerah","Merk"]))
    total_per_period = df.groupby(["Tahun","Bulan","Daerah"])["Total"].transform("sum")
    df["MS"] = df["Total"] / total_per_period

    df = df.sort_values(["Merk","Daerah","Kemasan","Tahun","nbulan"]).copy()
    df["MoM Growth %"] = df.groupby(["Merk","Daerah","Kemasan"])["MS"].pct_change(1)
    df["YoY Growth %"] = df.groupby(["Merk","Daerah","Kemasan"])["MS"].pct_change(12)

    df["MS_YTD"] = df.groupby(["Merk","Daerah","Kemasan","Tahun"])["MS"].cumsum()
    df["YtD Growth %"] = df.groupby(["Merk","Daerah","Kemasan"])["MS_YTD"].pct_change(12)

    # MSY by Total (YTD)
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

# ==== Uploads ====
uploaded_current = st.file_uploader("Upload Data Bulan Ini (Excel)", type=["xlsx"])
uploaded_db = st.file_uploader("Upload Database (Excel)", type=["xlsx"])
uploaded_map = st.file_uploader("Upload Mapping (Excel)", type=["xlsx"])

if uploaded_current and uploaded_db and uploaded_map:
    current = pd.read_excel(uploaded_current)
    db = pd.read_excel(uploaded_db)
    mapping_df = pd.read_excel(uploaded_map)

    # pastikan Total numerik
    if "Total" in current.columns: current["Total"] = to_numeric_series(current["Total"])
    if "Total" in db.columns: db["Total"] = to_numeric_series(db["Total"])

    # --- MAP HANYA DATA BULAN INI ---
    current_core = safe_select(current, BASE_COLS)

    seg_map = (mapping_df.drop_duplicates(["Merk","Daerah"])[["Merk","Daerah","Segment"]]
               if {"Merk","Daerah","Segment"}.issubset(mapping_df.columns) else None)
    area_map = (mapping_df.drop_duplicates(["Daerah"])[["Daerah","Area AP"]]
               if {"Daerah","Area AP"}.issubset(mapping_df.columns) else None)

    if seg_map is not None:
        current_core = current_core.merge(seg_map, on=["Merk","Daerah"], how="left")
    if area_map is not None:
        current_core = current_core.merge(area_map, on="Daerah", how="left")

    # Tambahkan kolom Tahun & nbulan ke seluruh row Data Bulan Ini
    current["Tahun"] = tahun_input
    current["Bulan"] = bulan_input
    current["nbulan"] = ((current["Tahun"].astype(int) - current["Tahun"].min()) * 12 
                             + current["Bulan"].astype(int))

    # Negara otomatis Domestik
    current["Negara"] = "Domestik"

    # --- ALIGN KOLUM & APPEND ---
    keep_cols = BASE_COLS + [c for c in ["Segment","Area AP"] if c in (set(db.columns)|set(current_core.columns))]
    keep_cols = [c for c in keep_cols if c in (set(db.columns)|set(current_core.columns))]

    db_aligned = safe_select(db, keep_cols)
    current_aligned = safe_select(current_core, keep_cols)

    # OPTIONAL: REPLACE MODE (hindari duplikat bulan yang sama di DB)
    y_now = int(current_aligned["Tahun"].max())
    m_now = current_aligned.loc[current_aligned["Tahun"].eq(y_now), "nbulan"].max()
    db_clean = db_aligned[~( (db_aligned["Tahun"]==y_now) & (db_aligned["nbulan"]==m_now) )]

    combined = pd.concat([db_clean, current_aligned], ignore_index=True)

    # --- HITUNG ---
    result = calc_ms_and_growth(combined)

    final_cols = keep_cols + ["MS","MoM Growth %","YoY Growth %","YtD Growth %","Total Merk YtD","Total All YtD","MSY"]
    final_cols = [c for c in final_cols if c in result.columns]
    final = (result[final_cols]
             .sort_values(["Tahun","nbulan","Merk"])
             .reset_index(drop=True))

    st.success(f"Ok! Baris: {len(final):,}")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        final.to_excel(w, index=False, sheet_name="Result")
    buf.seek(0)
    st.download_button("Download Data Hasil", buf, "Data_Hasil.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.dataframe(final.head(50))
else:
    st.info("Upload tiga file: Data Bulan Ini, Database, dan Mapping.")
