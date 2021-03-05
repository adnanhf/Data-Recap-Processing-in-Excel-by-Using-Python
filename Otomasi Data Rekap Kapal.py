# -*- coding: utf-8 -*-
"""
Program ini membuat n-sheet di excel
n-sheet tersebut berisi klasifikasi data
berdasarkan kriteria tertentu

Milestone program ini adalah:
    - Membaca data (file berformat .xlsx)
    - Memilih data berdasarkan kriteria tertentu
    - Mengurutkan data yang sudah dipilih berdasarkan tanggal
    - Mengatur properti tertentu seperti ukuran font dan jenis alinea
    - Menulis data di baris dan kolom tertentu
"""

''' ----------------------Kegiatan pada 04 Maret 2021----------------------'''
'''
# Impor library
import pandas as pd

# Membaca data, variabel rekap bunyu bertipe DataFrame
rekap_bunyu = pd.read_excel('SAMPLE DATA REKAP.xlsx',
                            sheet_name='REKAP',
                            header=4,
                            usecols=[x for x in range(26)])

# Menghapus baris data yang kosong
rekap_bunyu.dropna(how="all",inplace=True)

# Mengganti nama header kolom di sheet rekap bunyu
rekap_bunyu.rename(columns={'NO SPB':'NO SPB KELUAR','TANGGAL':'TANGGAL TIBA',
                            'JAM':'JAM TIBA','NO SPB.1':'NO SPB MASUK',
                            'TANGGAL.1':'TANGGAL TOLAK','JAM.1':'JAM TOLAK'},
                   inplace=True)

# Melihat isi kolom
print(pd.to_numeric(rekap_bunyu.GT))
'''
''' ------------------------------Akhir Kegiatan---------------------------'''

''' ----------------------Kegiatan pada 05 Maret 2021----------------------'''

# Impor library
import pandas as pd

# Fungsi untuk membaca seluruh sheet di excel
def read_excel_file(filename='SAMPLE DATA REKAP.xlsx',col_title=4):
    excel_df=pd.read_excel(filename,sheet_name=None,header=col_title)
    print('terdapat',len(excel_df),'sheet di dalam file ini')
    return excel_df

# Fungsi untuk menghapus kolom data yang kosong pada seluruh sheet
def drop_nan_cols(dfname):
    for i in dfname:
        dfname[i].dropna(axis=1,how='all',inplace=True)
    return dfname

# Fungsi untuk menghapus baris data yang kosong pada seluruh sheet
def drop_nan_rows(dfname):
    for i in dfname:
        dfname[i].dropna(axis=0,how='all',inplace=True)
    return dfname

def main():
    df_to_work=read_excel_file()
    df_to_work=drop_nan_cols(df_to_work)
    df_to_work=drop_nan_rows(df_to_work)
    print(df_to_work)
    
if __name__ == '__main__':
    main()

''' ------------------------------Akhir Kegiatan---------------------------'''