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

''' ----------------------Code Writing at 11th March 2021----------------------'''

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

# Fungsi mengubah format data ke list
def conv_df_to_list(dfname):
    if type(dfname) is not pd.DataFrame:
        nested_list=[]
        for i in dfname:
            child_list = [dfname[i].columns.values.tolist()] + dfname[i].values.tolist()
            nested_list.append(child_list)
        return nested_list
    elif type(dfname) is pd.DataFrame:
        main_list = [dfname.columns.values.tolist()] + dfname.values.tolist()
        return main_list

# Fungsi memilih data untuk SIB Kecil
def choose_for_small_sib(listdf):
    if len(listdf) > 1:
        
        dict_small_sib = {
            'Nama Kapal':[],
            'Nama Nahkoda':[],
            'Berat Kotor (GT)':[],
            'Berat Bersih (NT)':[],
            'Tanda Selar Menurut Pas Tahunan':[],
            'Tempat Kedudukan Kapal':[],
            'Tanggal Tiba':[],
            'Terakhir Singgah dari':[],
            'Kode Muatan Datang':[],
            'Tanggal Berangkat':[],
            'Persinggahan Pertama':[],
            'Kode Muatan Berangkat':[],
            'Keagenan/Kepemilikan':[],
            'KET.':[]}
        
        for i in range(len(listdf)):
            for j in range(len(listdf[i])):
                if j == 0:
                    pass
                elif j != 0:
                    print(len(listdf[i][j]))

# Fungsi utama untuk menjalankan program
def main():
    df_to_work=read_excel_file()
    df_to_work=drop_nan_cols(df_to_work)
    df_to_work=drop_nan_rows(df_to_work)
    list_df=conv_df_to_list(df_to_work)
    list_df_small_sib = choose_for_small_sib(list_df)
    
if __name__ == '__main__':
    main()

''' ------------------------------End of The Code Writing---------------------------'''
