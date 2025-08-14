import streamlit as st
import pandas as pd
import io

st.title("Automasi Market Share & Mapping")

# Upload file data utama
uploaded_data = st.file_uploader("Upload file Data (Excel)", type=["xlsx"])

if uploaded_data is not None:
    # Baca file data utama
    data = pd.read_excel(uploaded_data)

    # Baca mapping langsung dari repo
    mapping_df = pd.read_excel("Mapping.xlsx")  # Pastikan Mapping.xlsx ada di repo

    # --- Prosesnya sama kayak script kamu ---
    data_new = data[['Tahun', 'Bulan', 'Daerah', 'Pulau', 'Produsen', 'Total', 'Kemasan', 'Negara', 'Holding', 'Merk', 'nbulan']]
    data_new_grouped = data_new.groupby(['Tahun', 'Bulan', 'Daerah', 'Pulau', 'Produsen', 'Total', 'Kemasan', 'Holding', 'Merk', 'nbulan'], as_index=False)['Total'].sum()

    total_per_periode = data_new_grouped.groupby(['Tahun', 'Bulan', 'Daerah'])['Total'].transform('sum')
    data_new_grouped['MS'] = data_new_grouped['Total'] / total_per_periode
    data_new_grouped = data_new_grouped.sort_values(['Tahun', 'nbulan', 'Merk']).reset_index(drop=True)
    data_copy = data_new_grouped.copy()

    data_copy['MoM Growth %'] = data_copy.groupby(['Merk', 'Daerah','Kemasan'])['MS'].pct_change(1)
    data_copy['YoY Growth %'] = data_copy.groupby(['Merk', 'Daerah','Kemasan'])['MS'].pct_change(12)
    data_copy['MS_YTD'] = data_copy.groupby(['Merk', 'Daerah', 'Kemasan', 'Tahun'])['MS'].cumsum()
    data_copy['YtD Growth %'] = data_copy.groupby(['Merk', 'Daerah', 'Kemasan'])['MS_YTD'].pct_change(12)

    def calc_ytd_market_share(df):
        df = df.sort_values(['Daerah', 'Merk', 'Tahun', 'nbulan']).copy()
        df['Total Merk YtD'] = df.groupby(['Daerah', 'Merk', 'Tahun'])['Total'].cumsum()
        total_all = (
            df.groupby(['Daerah', 'Tahun', 'nbulan'])['Total']
              .sum()
              .groupby(level=['Daerah', 'Tahun'])
              .cumsum()
              .reset_index(name='Total All YtD')
        )
        df = df.merge(total_all, on=['Daerah', 'Tahun', 'nbulan'], how='left')
        df['MSY'] = df['Total Merk YtD'] / df['Total All YtD']
        return df

    full_data = calc_ytd_market_share(data_copy)

    mapping_df['Merk'] = mapping_df['Merk'].str.strip()
    mapping_df['Daerah'] = mapping_df['Daerah'].str.strip()
    full_data['Merk'] = full_data['Merk'].str.strip()
    full_data['Daerah'] = full_data['Daerah'].str.strip()

    segment_map = {(row['Merk'], row['Daerah']): row['Segment'] for _, row in mapping_df.iterrows()}
    area_ap_map = {row['Daerah']: row['Area AP'] for _, row in mapping_df.iterrows()}

    full_data['Segment'] = full_data.apply(lambda x: segment_map.get((x['Merk'], x['Daerah']), None), axis=1)
    full_data['Area AP'] = full_data['Daerah'].map(area_ap_map)
    full_data = full_data.sort_values(['Tahun', 'nbulan', 'Merk']).reset_index(drop=True)

    full_data = full_data[['Tahun', 'Bulan', 'Daerah', 'Pulau', 'Produsen', 'Kemasan',
           'Holding', 'Merk', 'nbulan', 'Total', 'Segment', 'Area AP', 'MS', 'MoM Growth %',
           'YoY Growth %', 'MS_YTD', 'YtD Growth %', 'Total Merk YtD',
           'Total All YtD', 'MSY']]
    del full_data['MS_YTD']

    # Simpan hasil ke buffer memori
    output = io.BytesIO()
    full_data.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    
    # Tombol download di Streamlit
    st.download_button(
        label="Download Data Hasil",
        data=output,
        file_name='Data_Hasil.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
