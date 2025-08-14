import pandas as pd
import numpy as np

data = pd.read_excel('Database ASI.xlsx')
mapping_df = pd.read_excel('Mapping.xlsx')

data_new = data[['Tahun', 'Bulan', 'Daerah', 'Pulau', 'Produsen', 'Total', 'Kemasan', 'Negara', 'Holding', 'Merk', 'nbulan']]

# Group berdasarkan Tahun, Bulan, Merk, Daerah
data_new_grouped = data_new.groupby(['Tahun', 'Bulan', 'Daerah', 'Pulau', 'Produsen', 'Total', 'Kemasan', 'Holding', 'Merk', 'nbulan'], as_index=False)['Total'].sum()

total_per_periode = data_new_grouped.groupby(['Tahun', 'Bulan', 'Daerah'])['Total'].transform('sum')
data_new_grouped['MS'] = data_new_grouped['Total'] / total_per_periode
data_new_grouped['MS'] = data_new_grouped['MS']

# Urutin
data_new_grouped = data_new_grouped.sort_values(['Tahun', 'nbulan', 'Merk']).reset_index(drop=True)

data_copy = data_new_grouped.copy()

# MoM Growth
data_copy['MoM Growth %'] = (
    data_copy
    .groupby(['Merk', 'Daerah','Kemasan'])['MS']
    .pct_change(1)
)

# YoY Growth
data_copy['YoY Growth %'] = (
    data_copy
    .groupby(['Merk', 'Daerah','Kemasan'])['MS']
    .pct_change(12)
)

# YTD dari MS (akumulasi Jan bulan ini per tahun)
data_copy['MS_YTD'] = (
    data_copy
    .groupby(['Merk', 'Daerah', 'Kemasan', 'Tahun'])['MS']
    .cumsum()
)

# Growth YTD vs bulan yang sama tahun lalu (pakai MS_YTD)
data_copy['YtD Growth %'] = (
    data_copy
    .groupby(['Merk', 'Daerah', 'Kemasan'])['MS_YTD']
    .pct_change(12)
)

# Perhitungan Market Share YTD
def calc_ytd_market_share(df):
    df = df.sort_values(['Daerah', 'Merk', 'Tahun', 'nbulan']).copy()

    # Numerator: YTD per merk
    df['Total Merk YtD'] = (
        df.groupby(['Daerah', 'Merk', 'Tahun'])['Total']
          .cumsum()
    )

    # Denominator: YTD semua merk (per bulan, tanpa pecah merk)
    total_all = (
        df.groupby(['Daerah', 'Tahun', 'nbulan'])['Total']
          .sum()
          .groupby(level=['Daerah', 'Tahun'])
          .cumsum()
          .reset_index(name='Total All YtD')
    )

    # Merge biar nilai total_all_ytd sama untuk semua merk di bulan itu
    df = df.merge(total_all, on=['Daerah', 'Tahun', 'nbulan'], how='left')

    # Market Share YTD
    df['MSY'] = df['Total Merk YtD'] / df['Total All YtD']

    return df

full_data = calc_ytd_market_share(data_copy)

# Mapping Segment
# Buat dictionary mapping Segment berdasarkan Merk & Daerah
segment_map = {
    (row['Merk'], row['Daerah']): row['Segment']
    for _, row in mapping_df.iterrows()
}

# Buat dictionary mapping Area AP berdasarkan Daerah
area_ap_map = {
    row['Daerah']: row['Area AP']
    for _, row in mapping_df.iterrows()
}

# Mapping ke dataframe utama
full_data['Segment'] = full_data.apply(
    lambda x: segment_map.get((x['Merk'], x['Daerah']), None), axis=1
)

full_data['Area AP'] = full_data['Daerah'].map(area_ap_map)

# Sort
full_data = full_data.sort_values(['Tahun', 'nbulan', 'Merk']).reset_index(drop=True)

full_data = full_data[['Tahun', 'Bulan', 'Daerah', 'Pulau', 'Produsen', 'Kemasan',
       'Holding', 'Merk', 'nbulan', 'Total', 'Segment', 'Area AP', 'MS', 'MoM Growth %',
       'YoY Growth %', 'MS_YTD', 'YtD Growth %', 'Total Merk YtD',
       'Total All YtD', 'MSY']]
del full_data['MS_YTD']
full_data.to_excel('Data Hasil.xlsx')
