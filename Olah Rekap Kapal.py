# -*- coding: utf-8 -*-
"""
Program ini membuat n-sheet di excel
n-sheet tersebut berisi klasifikasi data
berdasarkan kriteria tertentu

Milestone program ini adalah:
    - Membaca data (file berformat .xlsx) [Tercapai]
    - Memilih data berdasarkan kriteria tertentu [Memperbarui Metode]
    - Mengurutkan data yang sudah dipilih berdasarkan tanggal [Tercapai]
    - Mengatur properti tertentu seperti ukuran font dan jenis alinea [ X ]
    - Menulis data di baris dan kolom tertentu [Tercapai]
    
This code creates n-sheets excel as output
filled with classified data
upon certain criterias

The Milestone:
    - Read data (file format: xlsx) [Done]
    - Selecting data based on certain criterias [Updating]
    - Sort selected data based on date [Done]
    - Set certain properties ie. fonts & alignments [ X ]
    - Write data on certain row(s) or column(s) [Done]
"""

''' ----------------------Code Writing at 16th November 2021----------------------'''

# Importing library
import pandas as pd
import numpy as np
from systemtools.number import *

# functions for general purpose
#1 read data
def read_excel_file(filename='None',col_title=4):
    dfname=pd.read_excel(filename,sheet_name=None,header=col_title)
    #printing number of sheet in an excel file
    print('terdapat',len(dfname),'sheet di dalam file ini')
    
    #declare blank list for converting Pandas dataframe to list
    nested_list=[]

    for i in dfname:
        #the dropna, to delete any empty cells, based on row
        dfname[i].dropna(axis=0,how='any',inplace=True)
        #converting dataframe into list
        child_list = [dfname[i].columns.values.tolist()] + dfname[i].values.tolist()
        #printing number of row from each sheet, make sure no empty cells involving in dropna
        print('Sheet',i,':',len(dfname[i]),'baris')
        nested_list.append(child_list)

    return nested_list

#2 find root index in a nested list
def nestedlist_rootindex(thelist, char1, char2):
    for i in range(len(thelist)):
        if char1 == thelist[i][0] and char2 == thelist[i][1]:
            return i

# core functions, for monthly report
#1 SIB report functions
def sib_based(data=None):
    #change writing style at variable initiation
    dict_sibkecil = {'Nomor':[],'Nama Kapal':[],'Nama Nahkoda':[],
                    'GT':[],'NT':[],'Tanda Selar':[],'Tempat Pendaftaran':[],
                    'Tanggal Tiba':[],'Asal Kapal':[],'Kode Muatan Tiba':[],'Tanggal Tolak':[],
                    'Jam Tolak':[],'Tujuan Kapal':[],'Kode Muatan Tolak':[],'Keagenan':[],'Keterangan':[]
                    }

    dict_sibgede = {'Nomor':[],'Nama Kapal':[],'Bendera':[],'Nama Nahkoda':[],'Tempat Pendaftaran':[],
                    'GT':[],'NT':[],'Tanggal Tiba':[],'Asal Kapal':[],'Kode Muatan Tiba':[],'Tanggal Tolak':[],
                    'Jam Tolak':[],'Tujuan Kapal':[],'Kode Muatan Tolak':[],'Keagenan':[],'Keterangan':[],
                    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if data[i][j][5] <= 500:
               dict_sibkecil['Nomor'].append(j)
               dict_sibkecil['Nama Kapal'].append(data[i][j][4])
               dict_sibkecil['Nama Nahkoda'].append(data[i][j][9])
               dict_sibkecil['GT'].append(data[i][j][5])
               dict_sibkecil['NT'].append(data[i][j][6])
               dict_sibkecil['Tanda Selar'].append(data[i][j][7])
               dict_sibkecil['Tempat Pendaftaran'].append(data[i][j][8])
               dict_sibkecil['Tanggal Tiba'].append(data[i][j][15])
               dict_sibkecil['Asal Kapal'].append(data[i][j][14])
               dict_sibkecil['Kode Muatan Tiba'].append(data[i][j][17])
               dict_sibkecil['Tanggal Tolak'].append(data[i][j][26])
               dict_sibkecil['Jam Tolak'].append(data[i][j][27])
               dict_sibkecil['Tujuan Kapal'].append(data[i][j][25])
               dict_sibkecil['Kode Muatan Tolak'].append(data[i][j][28])
               dict_sibkecil['Keagenan'].append(data[i][j][13])
               dict_sibkecil['Keterangan'].append(data[i][j][34])
            elif data[i][j][5] > 500:
               dict_sibgede['Nomor'].append(j)
               dict_sibgede['Nama Kapal'].append(data[i][j][4])
               dict_sibgede['Bendera'].append(data[i][j][12])
               dict_sibgede['Nama Nahkoda'].append(data[i][j][9])
               dict_sibgede['Tempat Pendaftaran'].append(data[i][j][8])
               dict_sibgede['GT'].append(data[i][j][5])
               dict_sibgede['NT'].append(data[i][j][6])
               dict_sibgede['Tanggal Tiba'].append(data[i][j][15])
               dict_sibgede['Asal Kapal'].append(data[i][j][14])
               dict_sibgede['Kode Muatan Tiba'].append(data[i][j][17])
               dict_sibgede['Tanggal Tolak'].append(data[i][j][26])
               dict_sibgede['Jam Tolak'].append(data[i][j][27])
               dict_sibgede['Tujuan Kapal'].append(data[i][j][25])
               dict_sibgede['Kode Muatan Tolak'].append(data[i][j][28])
               dict_sibgede['Keagenan'].append(data[i][j][13])
               dict_sibgede['Keterangan'].append(data[i][j][34])

    #create dataframe from dicts
    sib_kecil = pd.DataFrame.from_dict(dict_sibkecil)
    sib_gede = pd.DataFrame.from_dict(dict_sibgede)
    
    #sorting on dates
    sib_kecil.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    sib_gede.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    
    #filling the right row number
    sib_kecil['Nomor'] = range(1,len(sib_kecil)+1)
    sib_gede['Nomor'] = range(1,len(sib_gede)+1)

    return sib_kecil,sib_gede

#2 TKII style report
def tkii_based(data=None):
    #this function covers previous four functions, because they have the same format
    dict_combined = {'Nomor':[],'Kode Kapal':[],'Nama Kapal':[],'Bendera':[],'Keagenan':[],'GT':[],
                     'Tanggal Tiba':[],'Jam Tiba':[],'Asal Kapal':[],'Tanggal Tambat':[],'Jam Tambat':[],
                     'Tanggal Tolak':[],'Tujuan Kapal':[],'Muatan Tiba':[],'Jml Muatan Tiba':[],
                     'Jns Muatan Tiba':[],'Muatan Tolak':[],'Jml Muatan Tolak':[],'Jns Muatan Tolak':[],
                     'Jam Tolak':[],'Lokasi Tolak':[],'Lokasi Bongkar':[],'Lokasi Muat':[],'Kategori':[]
                    }

    #main Process
    for i in range(len(data)):
        for j in range(1,len(data[i])):
            
            #uniting some labels into one and only kind
            if 'PERTAMINA BUNYU' in data[i][j][22] and 'PERTAMINA BUNYU' not in data[i][j][33]:
                 data[i][j][22] = 'TUKS PERTAMINA BUNYU'
            elif 'PERTAMINA BUNYU' not in data[i][j][22] and 'PERTAMINA BUNYU' in data[i][j][33]:
                 data[i][j][33] = 'TUKS PERTAMINA BUNYU'
            elif 'PERTAMINA BUNYU' in data[i][j][22] and 'PERTAMINA BUNYU' in data[i][j][33]:
                 data[i][j][22],data[i][j][33] = 'TUKS PERTAMINA BUNYU','TUKS PERTAMINA BUNYU'

            #some conditions to suit with the excel format in the usual monthly report
            if ';' in data[i][j][18] and ';' in data[i][j][29]:
                arr_load = data[i][j][18].split('; ')
                arr_num,arr_mu = data[i][j][19].split('; '),data[i][j][20].split('; ')

                depar_load = data[i][j][29].split('; ')
                depar_num,depar_mu = data[i][j][30].split('; '),data[i][j][31].split('; ')

                if len(arr_load) == len(depar_load):
                    for p in range(len(arr_load)):
                        if isinstance(arr_load[p],str) and p == 0:
                            dict_combined['Nomor'].append(j)
                            dict_combined['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_combined['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_combined['Bendera'].append(data[i][j][12])
                            dict_combined['Keagenan'].append(data[i][j][13])
                            dict_combined['GT'].append(data[i][j][5])
                            dict_combined['Tanggal Tiba'].append(data[i][j][15])
                            dict_combined['Jam Tiba'].append(data[i][j][16])
                            dict_combined['Asal Kapal'].append(data[i][j][14])
                            dict_combined['Tanggal Tambat'].append(data[i][j][15])
                            dict_combined['Jam Tambat'].append(data[i][j][16])
                            dict_combined['Tanggal Tolak'].append(data[i][j][26])
                            dict_combined['Tujuan Kapal'].append(data[i][j][25])
                            dict_combined['Muatan Tiba'].append(arr_load[p])
                            dict_combined['Jml Muatan Tiba'].append(arr_num[p])
                            dict_combined['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_combined['Muatan Tolak'].append(depar_load[p])
                            dict_combined['Jml Muatan Tolak'].append(depar_num[p])
                            dict_combined['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_combined['Jam Tolak'].append(data[i][j][27])
                            dict_combined['Lokasi Tolak'].append(data[i][j][24])
                            dict_combined['Lokasi Bongkar'].append(data[i][j][22])
                            dict_combined['Lokasi Muat'].append(data[i][j][33])
                            dict_combined['Kategori'].append(data[i][j][35])
                        else:
                            dict_combined['Nomor'].append(None)
                            dict_combined['Kode Kapal'].append(None)
                            dict_combined['Nama Kapal'].append(None)
                            dict_combined['Bendera'].append(None)
                            dict_combined['Keagenan'].append(None)
                            dict_combined['GT'].append(None)
                            dict_combined['Tanggal Tiba'].append(None)
                            dict_combined['Jam Tiba'].append(None)
                            dict_combined['Asal Kapal'].append(None)
                            dict_combined['Tanggal Tambat'].append(None)
                            dict_combined['Jam Tambat'].append(None)
                            dict_combined['Tanggal Tolak'].append(data[i][j][26])
                            dict_combined['Tujuan Kapal'].append(None)
                            dict_combined['Muatan Tiba'].append(arr_load[p])
                            dict_combined['Jml Muatan Tiba'].append(arr_num[p])
                            dict_combined['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_combined['Muatan Tolak'].append(depar_load[p])
                            dict_combined['Jml Muatan Tolak'].append(depar_num[p])
                            dict_combined['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_combined['Jam Tolak'].append(data[i][j][27])
                            dict_combined['Lokasi Tolak'].append(data[i][j][24])
                            dict_combined['Lokasi Bongkar'].append(data[i][j][22])
                            dict_combined['Lokasi Muat'].append(data[i][j][33])
                            dict_combined['Kategori'].append(data[i][j][35])
                        
                elif len(arr_load) < len(depar_load):
                    arr_load.extend(np.full([len(depar_load)-len(arr_load),1],None))
                    arr_num.extend(np.full([len(depar_num)-len(arr_num),1],None))
                    arr_mu.extend(np.full([len(depar_mu)-len(arr_mu),1],None))

                    for p in range(len(depar_load)):
                        if isinstance(arr_load[p],str) and p == 0:
                            dict_combined['Nomor'].append(j)
                            dict_combined['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_combined['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_combined['Bendera'].append(data[i][j][12])
                            dict_combined['Keagenan'].append(data[i][j][13])
                            dict_combined['GT'].append(data[i][j][5])
                            dict_combined['Tanggal Tiba'].append(data[i][j][15])
                            dict_combined['Jam Tiba'].append(data[i][j][16])
                            dict_combined['Asal Kapal'].append(data[i][j][14])
                            dict_combined['Tanggal Tambat'].append(data[i][j][15])
                            dict_combined['Jam Tambat'].append(data[i][j][16])
                            dict_combined['Tanggal Tolak'].append(data[i][j][26])
                            dict_combined['Tujuan Kapal'].append(data[i][j][25])
                            dict_combined['Muatan Tiba'].append(arr_load[p])
                            dict_combined['Jml Muatan Tiba'].append(arr_num[p])
                            dict_combined['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_combined['Muatan Tolak'].append(depar_load[p])
                            dict_combined['Jml Muatan Tolak'].append(depar_num[p])
                            dict_combined['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_combined['Jam Tolak'].append(data[i][j][27])
                            dict_combined['Lokasi Tolak'].append(data[i][j][24])
                            dict_combined['Lokasi Bongkar'].append(data[i][j][22])
                            dict_combined['Lokasi Muat'].append(data[i][j][33])
                            dict_combined['Kategori'].append(data[i][j][35])
                        elif isinstance(arr_load[p],str) and p != 0:
                            dict_combined['Nomor'].append(None)
                            dict_combined['Kode Kapal'].append(None)
                            dict_combined['Nama Kapal'].append(None)
                            dict_combined['Bendera'].append(None)
                            dict_combined['Keagenan'].append(None)
                            dict_combined['GT'].append(None)
                            dict_combined['Tanggal Tiba'].append(None)
                            dict_combined['Jam Tiba'].append(None)
                            dict_combined['Asal Kapal'].append(None)
                            dict_combined['Tanggal Tambat'].append(None)
                            dict_combined['Jam Tambat'].append(None)
                            dict_combined['Tanggal Tolak'].append(data[i][j][26])
                            dict_combined['Tujuan Kapal'].append(None)
                            dict_combined['Muatan Tiba'].append(arr_load[p])
                            dict_combined['Jml Muatan Tiba'].append(arr_num[p])
                            dict_combined['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_combined['Muatan Tolak'].append(depar_load[p])
                            dict_combined['Jml Muatan Tolak'].append(depar_num[p])
                            dict_combined['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_combined['Jam Tolak'].append(data[i][j][27])
                            dict_combined['Lokasi Tolak'].append(data[i][j][24])
                            dict_combined['Lokasi Bongkar'].append(data[i][j][22])
                            dict_combined['Lokasi Muat'].append(data[i][j][33])
                            dict_combined['Kategori'].append(data[i][j][35])
                        elif not isinstance(arr_load[p],str) and p != 0:
                            dict_combined['Nomor'].append(None)
                            dict_combined['Kode Kapal'].append(None)
                            dict_combined['Nama Kapal'].append(None)
                            dict_combined['Bendera'].append(None)
                            dict_combined['Keagenan'].append(None)
                            dict_combined['GT'].append(None)
                            dict_combined['Tanggal Tiba'].append(None)
                            dict_combined['Jam Tiba'].append(None)
                            dict_combined['Asal Kapal'].append(None)
                            dict_combined['Tanggal Tambat'].append(None)
                            dict_combined['Jam Tambat'].append(None)
                            dict_combined['Tanggal Tolak'].append(data[i][j][26])
                            dict_combined['Tujuan Kapal'].append(None)
                            dict_combined['Muatan Tiba'].append(None)
                            dict_combined['Jml Muatan Tiba'].append(None)
                            dict_combined['Jns Muatan Tiba'].append(None)
                            dict_combined['Muatan Tolak'].append(depar_load[p])
                            dict_combined['Jml Muatan Tolak'].append(depar_num[p])
                            dict_combined['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_combined['Jam Tolak'].append(data[i][j][27])
                            dict_combined['Lokasi Tolak'].append(data[i][j][24])
                            dict_combined['Lokasi Bongkar'].append(data[i][j][22])
                            dict_combined['Lokasi Muat'].append(data[i][j][33])
                            dict_combined['Kategori'].append(data[i][j][35])

                elif len(arr_load) > len(depar_load):
                    depar_load.extend(np.full([1,len(arr_load)-len(depar_load)],None))
                    depar_num.extend(np.full([1,len(arr_num)-len(depar_num)],None))
                    depar_mu.extend(np.full([1,len(arr_mu)-len(depar_mu)],None))

                    for p in range(len(arr_load)):
                        if isinstance(depar_load[p],str) and p == 0:
                            dict_combined['Nomor'].append(j)
                            dict_combined['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_combined['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_combined['Bendera'].append(data[i][j][12])
                            dict_combined['Keagenan'].append(data[i][j][13])
                            dict_combined['GT'].append(data[i][j][5])
                            dict_combined['Tanggal Tiba'].append(data[i][j][15])
                            dict_combined['Jam Tiba'].append(data[i][j][16])
                            dict_combined['Asal Kapal'].append(data[i][j][14])
                            dict_combined['Tanggal Tambat'].append(data[i][j][15])
                            dict_combined['Jam Tambat'].append(data[i][j][16])
                            dict_combined['Tanggal Tolak'].append(data[i][j][26])
                            dict_combined['Tujuan Kapal'].append(data[i][j][25])
                            dict_combined['Muatan Tiba'].append(arr_load[p])
                            dict_combined['Jml Muatan Tiba'].append(arr_num[p])
                            dict_combined['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_combined['Muatan Tolak'].append(depar_load[p])
                            dict_combined['Jml Muatan Tolak'].append(depar_num[p])
                            dict_combined['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_combined['Jam Tolak'].append(data[i][j][27])
                            dict_combined['Lokasi Tolak'].append(data[i][j][24])
                            dict_combined['Lokasi Bongkar'].append(data[i][j][22])
                            dict_combined['Lokasi Muat'].append(data[i][j][33])
                            dict_combined['Kategori'].append(data[i][j][35])
                        elif isinstance(depar_load[p],str) and p != 0:
                            dict_combined['Nomor'].append(None)
                            dict_combined['Kode Kapal'].append(None)
                            dict_combined['Nama Kapal'].append(None)
                            dict_combined['Bendera'].append(None)
                            dict_combined['Keagenan'].append(None)
                            dict_combined['GT'].append(None)
                            dict_combined['Tanggal Tiba'].append(None)
                            dict_combined['Jam Tiba'].append(None)
                            dict_combined['Asal Kapal'].append(None)
                            dict_combined['Tanggal Tambat'].append(None)
                            dict_combined['Jam Tambat'].append(None)
                            dict_combined['Tanggal Tolak'].append(data[i][j][26])
                            dict_combined['Tujuan Kapal'].append(None)
                            dict_combined['Muatan Tiba'].append(arr_load[p])
                            dict_combined['Jml Muatan Tiba'].append(arr_num[p])
                            dict_combined['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_combined['Muatan Tolak'].append(depar_load[p])
                            dict_combined['Jml Muatan Tolak'].append(depar_num[p])
                            dict_combined['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_combined['Jam Tolak'].append(data[i][j][27])
                            dict_combined['Lokasi Tolak'].append(data[i][j][24])
                            dict_combined['Lokasi Bongkar'].append(data[i][j][22])
                            dict_combined['Lokasi Muat'].append(data[i][j][33])
                            dict_combined['Kategori'].append(data[i][j][35])
                        elif not isinstance(depar_load[p],str) and p != 0:
                            dict_combined['Nomor'].append(None)
                            dict_combined['Kode Kapal'].append(None)
                            dict_combined['Nama Kapal'].append(None)
                            dict_combined['Bendera'].append(None)
                            dict_combined['Keagenan'].append(None)
                            dict_combined['GT'].append(None)
                            dict_combined['Tanggal Tiba'].append(None)
                            dict_combined['Jam Tiba'].append(None)
                            dict_combined['Asal Kapal'].append(None)
                            dict_combined['Tanggal Tambat'].append(None)
                            dict_combined['Jam Tambat'].append(None)
                            dict_combined['Tanggal Tolak'].append(data[i][j][26])
                            dict_combined['Tujuan Kapal'].append(None)
                            dict_combined['Muatan Tiba'].append(arr_load[p])
                            dict_combined['Jml Muatan Tiba'].append(arr_num[p])
                            dict_combined['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_combined['Muatan Tolak'].append(None)
                            dict_combined['Jml Muatan Tolak'].append(None)
                            dict_combined['Jns Muatan Tolak'].append(None)
                            dict_combined['Jam Tolak'].append(data[i][j][27])
                            dict_combined['Lokasi Tolak'].append(data[i][j][24])
                            dict_combined['Lokasi Bongkar'].append(data[i][j][22])
                            dict_combined['Lokasi Muat'].append(data[i][j][33])
                            dict_combined['Kategori'].append(data[i][j][35])

            elif ';' in data[i][j][18] and ';' not in data[i][j][29]:
                arr_load = data[i][j][18].split('; ')
                arr_num = data[i][j][19].split('; ')
                arr_mu = data[i][j][20].split('; ')

                for p in range(len(arr_load)):
                    if p == 0:
                        dict_combined['Nomor'].append(j)
                        dict_combined['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                        dict_combined['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                        dict_combined['Bendera'].append(data[i][j][12])
                        dict_combined['Keagenan'].append(data[i][j][13])
                        dict_combined['GT'].append(data[i][j][5])
                        dict_combined['Tanggal Tiba'].append(data[i][j][15])
                        dict_combined['Jam Tiba'].append(data[i][j][16])
                        dict_combined['Asal Kapal'].append(data[i][j][14])
                        dict_combined['Tanggal Tambat'].append(data[i][j][15])
                        dict_combined['Jam Tambat'].append(data[i][j][16])
                        dict_combined['Tanggal Tolak'].append(data[i][j][26])
                        dict_combined['Tujuan Kapal'].append(data[i][j][25])
                        dict_combined['Muatan Tiba'].append(arr_load[p])
                        dict_combined['Jml Muatan Tiba'].append(arr_num[p])
                        dict_combined['Jns Muatan Tiba'].append(arr_mu[p])
                        dict_combined['Muatan Tolak'].append(data[i][j][29])
                        dict_combined['Jml Muatan Tolak'].append(data[i][j][30])
                        dict_combined['Jns Muatan Tolak'].append(data[i][j][31])
                        dict_combined['Jam Tolak'].append(data[i][j][27])
                        dict_combined['Lokasi Tolak'].append(data[i][j][24])
                        dict_combined['Lokasi Bongkar'].append(data[i][j][22])
                        dict_combined['Lokasi Muat'].append(data[i][j][33])
                        dict_combined['Kategori'].append(data[i][j][35])
                    elif p != 0:
                        dict_combined['Nomor'].append(None)
                        dict_combined['Kode Kapal'].append(None)
                        dict_combined['Nama Kapal'].append(None)
                        dict_combined['Bendera'].append(None)
                        dict_combined['Keagenan'].append(None)
                        dict_combined['GT'].append(None)
                        dict_combined['Tanggal Tiba'].append(None)
                        dict_combined['Jam Tiba'].append(None)
                        dict_combined['Asal Kapal'].append(None)
                        dict_combined['Tanggal Tambat'].append(None)
                        dict_combined['Jam Tambat'].append(None)
                        dict_combined['Tanggal Tolak'].append(data[i][j][26])
                        dict_combined['Tujuan Kapal'].append(None)
                        dict_combined['Muatan Tiba'].append(arr_load[p])
                        dict_combined['Jml Muatan Tiba'].append(arr_num[p])
                        dict_combined['Jns Muatan Tiba'].append(arr_mu[p])
                        dict_combined['Muatan Tolak'].append(None)
                        dict_combined['Jml Muatan Tolak'].append(None)
                        dict_combined['Jns Muatan Tolak'].append(None)
                        dict_combined['Jam Tolak'].append(data[i][j][27])
                        dict_combined['Lokasi Tolak'].append(data[i][j][24])
                        dict_combined['Lokasi Bongkar'].append(data[i][j][22])
                        dict_combined['Lokasi Muat'].append(data[i][j][33])
                        dict_combined['Kategori'].append(data[i][j][35])

            elif ';' not in data[i][j][18] and ';' in data[i][j][29]:
                depar_load = data[i][j][29].split('; ')
                depar_num = data[i][j][30].split('; ')
                depar_mu = data[i][j][31].split('; ')

                for p in range(len(depar_load)):
                    if p == 0:
                        dict_combined['Nomor'].append(j)
                        dict_combined['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                        dict_combined['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                        dict_combined['Bendera'].append(data[i][j][12])
                        dict_combined['Keagenan'].append(data[i][j][13])
                        dict_combined['GT'].append(data[i][j][5])
                        dict_combined['Tanggal Tiba'].append(data[i][j][15])
                        dict_combined['Jam Tiba'].append(data[i][j][16])
                        dict_combined['Asal Kapal'].append(data[i][j][14])
                        dict_combined['Tanggal Tambat'].append(data[i][j][15])
                        dict_combined['Jam Tambat'].append(data[i][j][16])
                        dict_combined['Tanggal Tolak'].append(data[i][j][26])
                        dict_combined['Tujuan Kapal'].append(data[i][j][25])
                        dict_combined['Muatan Tiba'].append(data[i][j][18])
                        dict_combined['Jml Muatan Tiba'].append(data[i][j][19])
                        dict_combined['Jns Muatan Tiba'].append(data[i][j][20])
                        dict_combined['Muatan Tolak'].append(depar_load[p])
                        dict_combined['Jml Muatan Tolak'].append(depar_num[p])
                        dict_combined['Jns Muatan Tolak'].append(depar_mu[p])
                        dict_combined['Jam Tolak'].append(data[i][j][27])
                        dict_combined['Lokasi Tolak'].append(data[i][j][24])
                        dict_combined['Lokasi Bongkar'].append(data[i][j][22])
                        dict_combined['Lokasi Muat'].append(data[i][j][33])
                        dict_combined['Kategori'].append(data[i][j][35])
                    elif p != 0:
                        dict_combined['Nomor'].append(None)
                        dict_combined['Kode Kapal'].append(None)
                        dict_combined['Nama Kapal'].append(None)
                        dict_combined['Bendera'].append(None)
                        dict_combined['Keagenan'].append(None)
                        dict_combined['GT'].append(None)
                        dict_combined['Tanggal Tiba'].append(None)
                        dict_combined['Jam Tiba'].append(None)
                        dict_combined['Asal Kapal'].append(None)
                        dict_combined['Tanggal Tambat'].append(None)
                        dict_combined['Jam Tambat'].append(None)
                        dict_combined['Tanggal Tolak'].append(data[i][j][26])
                        dict_combined['Tujuan Kapal'].append(None)
                        dict_combined['Muatan Tiba'].append(None)
                        dict_combined['Jml Muatan Tiba'].append(None)
                        dict_combined['Jns Muatan Tiba'].append(None)
                        dict_combined['Muatan Tolak'].append(depar_load[p])
                        dict_combined['Jml Muatan Tolak'].append(depar_num[p])
                        dict_combined['Jns Muatan Tolak'].append(depar_mu[p])
                        dict_combined['Jam Tolak'].append(data[i][j][27])
                        dict_combined['Lokasi Tolak'].append(data[i][j][24])
                        dict_combined['Lokasi Bongkar'].append(data[i][j][22])
                        dict_combined['Lokasi Muat'].append(data[i][j][33])
                        dict_combined['Kategori'].append(data[i][j][35])

            elif ';' not in data[i][j][18] and ';' not in data[i][j][29]:
                dict_combined['Nomor'].append(j)
                dict_combined['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_combined['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_combined['Bendera'].append(data[i][j][12])
                dict_combined['Keagenan'].append(data[i][j][13])
                dict_combined['GT'].append(data[i][j][5])
                dict_combined['Tanggal Tiba'].append(data[i][j][15])
                dict_combined['Jam Tiba'].append(data[i][j][16])
                dict_combined['Asal Kapal'].append(data[i][j][14])
                dict_combined['Tanggal Tambat'].append(data[i][j][15])
                dict_combined['Jam Tambat'].append(data[i][j][16])
                dict_combined['Tanggal Tolak'].append(data[i][j][26])
                dict_combined['Tujuan Kapal'].append(data[i][j][25])
                dict_combined['Muatan Tiba'].append(data[i][j][18])
                dict_combined['Jml Muatan Tiba'].append(data[i][j][19])
                dict_combined['Jns Muatan Tiba'].append(data[i][j][20])
                dict_combined['Muatan Tolak'].append(data[i][j][29])
                dict_combined['Jml Muatan Tolak'].append(data[i][j][30])
                dict_combined['Jns Muatan Tolak'].append(data[i][j][31])
                dict_combined['Jam Tolak'].append(data[i][j][27])
                dict_combined['Lokasi Tolak'].append(data[i][j][24])
                dict_combined['Lokasi Bongkar'].append(data[i][j][22])
                dict_combined['Lokasi Muat'].append(data[i][j][33])
                dict_combined['Kategori'].append(data[i][j][35])

    #creating dataframe from dicts and sort it by dates
    precombi = pd.DataFrame.from_dict(dict_combined)
    precombi.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    #creating new dataframes based on some categories
    combi1 = precombi.loc[precombi['Kategori'] == 'DOMESTIK']
    combi2 = precombi.loc[precombi['Kategori'] == 'EKSPOR']
    combined = combi1.append(combi2)

    bunyu = precombi.loc[precombi['Lokasi Tolak'] == 'BUNYU']
    albunyu = precombi.loc[(precombi['Lokasi Bongkar'] == 'TUKS PERTAMINA BUNYU') | (precombi['Lokasi Muat'] == 'TUKS PERTAMINA BUNYU')]
    
    tanmua1 = precombi.loc[precombi['Kode Kapal'].isin(['TB','TK','OB','BG'])]
    tanmua2 = precombi.loc[(precombi['Muatan Tiba'] == 'NIHIL') & (precombi['Muatan Tolak'] == 'NIHIL')]
    tanmua = tanmua1.append(tanmua2)
    tanmua.drop_duplicates(inplace=True)
    tanmua.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    #minor refining for the output
    foutput = [combined,bunyu,albunyu,tanmua]
    for funcs in range(len(foutput)):
        counter = 1
        for i in range(len(foutput[funcs])):
            if np.isnan(foutput[funcs].iat[i,0]) == False:
                foutput[funcs].iat[i,0] = counter
                counter+=1
            elif np.isnan(foutput[funcs].iat[i,0]) == True:
                foutput[funcs].iat[i,11] = None
                foutput[funcs].iat[i,12] = None
        foutput[funcs].drop(['Jam Tolak','Lokasi Tolak','Lokasi Bongkar','Lokasi Muat','Kategori'],axis=1,inplace=True)
    return foutput

#3 Domestic & Export style report
def domexp_based(data=None):
    #this function covers previous thirteen functions, because they have the same format
    dict_domcateg = {'Nomor':[],'Kode Kapal':[],'Nama Kapal':[],'Keagenan':[],'Bendera':[],'GT':[],
                     'Trayek':[],'Tanggal Tiba':[],'Tanggal Tolak':[],'Muatan Tiba':[],'Jml Muatan Tiba':[],
                     'Jns Muatan Tiba':[],'Asal Kapal':[],'Muatan Tolak':[],'Jml Muatan Tolak':[],
                     'Jns Muatan Tolak':[],'Tujuan Kapal':[],'Jam Tolak':[],'Kode Muatan Tiba':[],
                     'Kode Muatan Tolak':[],'Kategori':[]
                    }

    #main Process
    for i in range(len(data)):
        for j in range(1,len(data[i])):
            #some nested conditions to suit with the excel format in the usual monthly report
            if ';' in data[i][j][18] and ';' in data[i][j][29]:
                arr_load = data[i][j][18].split('; ')
                arr_num,arr_mu = data[i][j][19].split('; '),data[i][j][20].split('; ')

                depar_load = data[i][j][29].split('; ')
                depar_num,depar_mu = data[i][j][30].split('; '),data[i][j][31].split('; ')

                if len(arr_load) == len(depar_load):
                    for p in range(len(arr_load)):
                        if isinstance(arr_load[p],str) and p == 0:
                            dict_domcateg['Nomor'].append(j)
                            dict_domcateg['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_domcateg['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_domcateg['Keagenan'].append(data[i][j][13])
                            dict_domcateg['Bendera'].append(data[i][j][12])
                            dict_domcateg['GT'].append(data[i][j][5])
                            dict_domcateg['Trayek'].append('T')
                            dict_domcateg['Tanggal Tiba'].append(data[i][j][15])
                            dict_domcateg['Tanggal Tolak'].append(data[i][j][26])
                            dict_domcateg['Muatan Tiba'].append(arr_load[p])
                            dict_domcateg['Jml Muatan Tiba'].append(arr_num[p])
                            dict_domcateg['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_domcateg['Asal Kapal'].append(data[i][j][14])
                            dict_domcateg['Muatan Tolak'].append(depar_load[p])
                            dict_domcateg['Jml Muatan Tolak'].append(depar_num[p])
                            dict_domcateg['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_domcateg['Tujuan Kapal'].append(data[i][j][25])
                            dict_domcateg['Jam Tolak'].append(data[i][j][27])
                            dict_domcateg['Kode Muatan Tiba'].append(data[i][j][21])
                            dict_domcateg['Kode Muatan Tolak'].append(data[i][j][32])
                            dict_domcateg['Kategori'].append(data[i][j][35])
                        else:
                            dict_domcateg['Nomor'].append(None)
                            dict_domcateg['Kode Kapal'].append(None)
                            dict_domcateg['Nama Kapal'].append(None)
                            dict_domcateg['Keagenan'].append(None)
                            dict_domcateg['Bendera'].append(None)
                            dict_domcateg['GT'].append(None)
                            dict_domcateg['Trayek'].append(None)
                            dict_domcateg['Tanggal Tiba'].append(None)
                            dict_domcateg['Tanggal Tolak'].append(data[i][j][24])
                            dict_domcateg['Muatan Tiba'].append(arr_load[p])
                            dict_domcateg['Jml Muatan Tiba'].append(arr_num[p])
                            dict_domcateg['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_domcateg['Asal Kapal'].append(None)
                            dict_domcateg['Muatan Tolak'].append(depar_load[p])
                            dict_domcateg['Jml Muatan Tolak'].append(depar_num[p])
                            dict_domcateg['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_domcateg['Tujuan Kapal'].append(None)
                            dict_domcateg['Jam Tolak'].append(data[i][j][27])
                            dict_domcateg['Kode Muatan Tiba'].append(data[i][j][21])
                            dict_domcateg['Kode Muatan Tolak'].append(data[i][j][32])
                            dict_domcateg['Kategori'].append(data[i][j][35])

                elif len(arr_load) < len(depar_load):
                    arr_load.extend(np.full([len(depar_load)-len(arr_load),1],None))
                    arr_num.extend(np.full([len(depar_num)-len(arr_num),1],None))
                    arr_mu.extend(np.full([len(depar_mu)-len(arr_mu),1],None))

                    for p in range(len(depar_load)):
                        if isinstance(arr_load[p],str) and p == 0:
                            dict_domcateg['Nomor'].append(j)
                            dict_domcateg['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_domcateg['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_domcateg['Keagenan'].append(data[i][j][13])
                            dict_domcateg['Bendera'].append(data[i][j][12])
                            dict_domcateg['GT'].append(data[i][j][5])
                            dict_domcateg['Trayek'].append('T')
                            dict_domcateg['Tanggal Tiba'].append(data[i][j][15])
                            dict_domcateg['Tanggal Tolak'].append(data[i][j][26])
                            dict_domcateg['Muatan Tiba'].append(arr_load[p])
                            dict_domcateg['Jml Muatan Tiba'].append(arr_num[p])
                            dict_domcateg['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_domcateg['Asal Kapal'].append(data[i][j][14])
                            dict_domcateg['Muatan Tolak'].append(depar_load[p])
                            dict_domcateg['Jml Muatan Tolak'].append(depar_num[p])
                            dict_domcateg['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_domcateg['Tujuan Kapal'].append(data[i][j][25])
                            dict_domcateg['Jam Tolak'].append(data[i][j][27])
                            dict_domcateg['Kode Muatan Tiba'].append(data[i][j][21])
                            dict_domcateg['Kode Muatan Tolak'].append(data[i][j][32])
                            dict_domcateg['Kategori'].append(data[i][j][35])
                        elif isinstance(arr_load[p],str) and p != 0:
                            dict_domcateg['Nomor'].append(None)
                            dict_domcateg['Kode Kapal'].append(None)
                            dict_domcateg['Nama Kapal'].append(None)
                            dict_domcateg['Keagenan'].append(None)
                            dict_domcateg['Bendera'].append(None)
                            dict_domcateg['GT'].append(None)
                            dict_domcateg['Trayek'].append(None)
                            dict_domcateg['Tanggal Tiba'].append(None)
                            dict_domcateg['Tanggal Tolak'].append(data[i][j][24])
                            dict_domcateg['Muatan Tiba'].append(arr_load[p])
                            dict_domcateg['Jml Muatan Tiba'].append(arr_num[p])
                            dict_domcateg['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_domcateg['Asal Kapal'].append(None)
                            dict_domcateg['Muatan Tolak'].append(depar_load[p])
                            dict_domcateg['Jml Muatan Tolak'].append(depar_num[p])
                            dict_domcateg['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_domcateg['Tujuan Kapal'].append(None)
                            dict_domcateg['Jam Tolak'].append(data[i][j][27])
                            dict_domcateg['Kode Muatan Tiba'].append(data[i][j][21])
                            dict_domcateg['Kode Muatan Tolak'].append(data[i][j][32])
                            dict_domcateg['Kategori'].append(data[i][j][35])
                        elif not isinstance(arr_load[p],str) and p != 0:
                            dict_domcateg['Nomor'].append(None)
                            dict_domcateg['Kode Kapal'].append(None)
                            dict_domcateg['Nama Kapal'].append(None)
                            dict_domcateg['Keagenan'].append(None)
                            dict_domcateg['Bendera'].append(None)
                            dict_domcateg['GT'].append(None)
                            dict_domcateg['Trayek'].append(None)
                            dict_domcateg['Tanggal Tiba'].append(None)
                            dict_domcateg['Tanggal Tolak'].append(data[i][j][24])
                            dict_domcateg['Muatan Tiba'].append(None)
                            dict_domcateg['Jml Muatan Tiba'].append(None)
                            dict_domcateg['Jns Muatan Tiba'].append(None)
                            dict_domcateg['Asal Kapal'].append(None)
                            dict_domcateg['Muatan Tolak'].append(depar_load[p])
                            dict_domcateg['Jml Muatan Tolak'].append(depar_num[p])
                            dict_domcateg['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_domcateg['Tujuan Kapal'].append(None)
                            dict_domcateg['Jam Tolak'].append(data[i][j][27])
                            dict_domcateg['Kode Muatan Tiba'].append(data[i][j][21])
                            dict_domcateg['Kode Muatan Tolak'].append(data[i][j][32])
                            dict_domcateg['Kategori'].append(data[i][j][35])

                elif len(arr_load) > len(depar_load):
                    depar_load.extend(np.full([1,len(arr_load)-len(depar_load)],None))
                    depar_num.extend(np.full([1,len(arr_num)-len(depar_num)],None))
                    depar_mu.extend(np.full([1,len(arr_mu)-len(depar_mu)],None))

                    for p in range(len(arr_load)):
                        if isinstance(depar_load[p],str) and p == 0:
                            dict_domcateg['Nomor'].append(j)
                            dict_domcateg['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_domcateg['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_domcateg['Keagenan'].append(data[i][j][13])
                            dict_domcateg['Bendera'].append(data[i][j][12])
                            dict_domcateg['GT'].append(data[i][j][5])
                            dict_domcateg['Trayek'].append('T')
                            dict_domcateg['Tanggal Tiba'].append(data[i][j][15])
                            dict_domcateg['Tanggal Tolak'].append(data[i][j][26])
                            dict_domcateg['Muatan Tiba'].append(arr_load[p])
                            dict_domcateg['Jml Muatan Tiba'].append(arr_num[p])
                            dict_domcateg['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_domcateg['Asal Kapal'].append(data[i][j][14])
                            dict_domcateg['Muatan Tolak'].append(depar_load[p])
                            dict_domcateg['Jml Muatan Tolak'].append(depar_num[p])
                            dict_domcateg['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_domcateg['Tujuan Kapal'].append(data[i][j][25])
                            dict_domcateg['Jam Tolak'].append(data[i][j][27])
                            dict_domcateg['Kode Muatan Tiba'].append(data[i][j][21])
                            dict_domcateg['Kode Muatan Tolak'].append(data[i][j][32])
                            dict_domcateg['Kategori'].append(data[i][j][35])
                        elif isinstance(depar_load[p],str) and p != 0:
                            dict_domcateg['Nomor'].append(None)
                            dict_domcateg['Kode Kapal'].append(None)
                            dict_domcateg['Nama Kapal'].append(None)
                            dict_domcateg['Keagenan'].append(None)
                            dict_domcateg['Bendera'].append(None)
                            dict_domcateg['GT'].append(None)
                            dict_domcateg['Trayek'].append(None)
                            dict_domcateg['Tanggal Tiba'].append(None)
                            dict_domcateg['Tanggal Tolak'].append(data[i][j][24])
                            dict_domcateg['Muatan Tiba'].append(arr_load[p])
                            dict_domcateg['Jml Muatan Tiba'].append(arr_num[p])
                            dict_domcateg['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_domcateg['Asal Kapal'].append(None)
                            dict_domcateg['Muatan Tolak'].append(depar_load[p])
                            dict_domcateg['Jml Muatan Tolak'].append(depar_num[p])
                            dict_domcateg['Jns Muatan Tolak'].append(depar_mu[p])
                            dict_domcateg['Tujuan Kapal'].append(None)
                            dict_domcateg['Jam Tolak'].append(data[i][j][27])
                            dict_domcateg['Kode Muatan Tiba'].append(data[i][j][21])
                            dict_domcateg['Kode Muatan Tolak'].append(data[i][j][32])
                            dict_domcateg['Kategori'].append(data[i][j][35])
                        elif not isinstance(depar_load[p],str) and p != 0:
                            dict_domcateg['Nomor'].append(None)
                            dict_domcateg['Kode Kapal'].append(None)
                            dict_domcateg['Nama Kapal'].append(None)
                            dict_domcateg['Keagenan'].append(None)
                            dict_domcateg['Bendera'].append(None)
                            dict_domcateg['GT'].append(None)
                            dict_domcateg['Trayek'].append(None)
                            dict_domcateg['Tanggal Tiba'].append(None)
                            dict_domcateg['Tanggal Tolak'].append(data[i][j][24])
                            dict_domcateg['Muatan Tiba'].append(arr_load[p])
                            dict_domcateg['Jml Muatan Tiba'].append(arr_num[p])
                            dict_domcateg['Jns Muatan Tiba'].append(arr_mu[p])
                            dict_domcateg['Asal Kapal'].append(None)
                            dict_domcateg['Muatan Tolak'].append(None)
                            dict_domcateg['Jml Muatan Tolak'].append(None)
                            dict_domcateg['Jns Muatan Tolak'].append(None)
                            dict_domcateg['Tujuan Kapal'].append(None)
                            dict_domcateg['Jam Tolak'].append(data[i][j][27])
                            dict_domcateg['Kode Muatan Tiba'].append(data[i][j][21])
                            dict_domcateg['Kode Muatan Tolak'].append(data[i][j][32])
                            dict_domcateg['Kategori'].append(data[i][j][35])

            elif ';' in data[i][j][18] and ';' not in data[i][j][29]:
                arr_load = data[i][j][18].split('; ')
                arr_num = data[i][j][19].split('; ')
                arr_mu = data[i][j][20].split('; ')

                for p in range(len(arr_load)):
                    if p == 0:
                        dict_domcateg['Nomor'].append(j)
                        dict_domcateg['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                        dict_domcateg['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                        dict_domcateg['Keagenan'].append(data[i][j][13])
                        dict_domcateg['Bendera'].append(data[i][j][12])
                        dict_domcateg['GT'].append(data[i][j][5])
                        dict_domcateg['Trayek'].append('T')
                        dict_domcateg['Tanggal Tiba'].append(data[i][j][15])
                        dict_domcateg['Tanggal Tolak'].append(data[i][j][26])
                        dict_domcateg['Muatan Tiba'].append(arr_load[p])
                        dict_domcateg['Jml Muatan Tiba'].append(arr_num[p])
                        dict_domcateg['Jns Muatan Tiba'].append(arr_mu[p])
                        dict_domcateg['Asal Kapal'].append(data[i][j][14])
                        dict_domcateg['Muatan Tolak'].append(data[i][j][29])
                        dict_domcateg['Jml Muatan Tolak'].append(data[i][j][30])
                        dict_domcateg['Jns Muatan Tolak'].append(data[i][j][31])
                        dict_domcateg['Tujuan Kapal'].append(data[i][j][25])
                        dict_domcateg['Jam Tolak'].append(data[i][j][27])
                        dict_domcateg['Kode Muatan Tiba'].append(data[i][j][21])
                        dict_domcateg['Kode Muatan Tolak'].append(data[i][j][32])
                        dict_domcateg['Kategori'].append(data[i][j][35])
                    elif p != 0:
                        dict_domcateg['Nomor'].append(None)
                        dict_domcateg['Kode Kapal'].append(None)
                        dict_domcateg['Nama Kapal'].append(None)
                        dict_domcateg['Keagenan'].append(None)
                        dict_domcateg['Bendera'].append(None)
                        dict_domcateg['GT'].append(None)
                        dict_domcateg['Trayek'].append(None)
                        dict_domcateg['Tanggal Tiba'].append(None)
                        dict_domcateg['Tanggal Tolak'].append(data[i][j][26])
                        dict_domcateg['Muatan Tiba'].append(arr_load[p])
                        dict_domcateg['Jml Muatan Tiba'].append(arr_num[p])
                        dict_domcateg['Jns Muatan Tiba'].append(arr_mu[p])
                        dict_domcateg['Asal Kapal'].append(None)
                        dict_domcateg['Muatan Tolak'].append(None)
                        dict_domcateg['Jml Muatan Tolak'].append(None)
                        dict_domcateg['Jns Muatan Tolak'].append(None)
                        dict_domcateg['Tujuan Kapal'].append(None)
                        dict_domcateg['Jam Tolak'].append(data[i][j][27])
                        dict_domcateg['Kode Muatan Tiba'].append(data[i][j][21])
                        dict_domcateg['Kode Muatan Tolak'].append(data[i][j][32])
                        dict_domcateg['Kategori'].append(data[i][j][35])

            elif ';' not in data[i][j][18] and ';' in data[i][j][29]:
                depar_load = data[i][j][29].split('; ')
                depar_num = data[i][j][30].split('; ')
                depar_mu = data[i][j][31].split('; ')

                for p in range(len(depar_load)):
                    if p == 0:
                        dict_domcateg['Nomor'].append(j)
                        dict_domcateg['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                        dict_domcateg['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                        dict_domcateg['Keagenan'].append(data[i][j][13])
                        dict_domcateg['Bendera'].append(data[i][j][12])
                        dict_domcateg['GT'].append(data[i][j][5])
                        dict_domcateg['Trayek'].append('T')
                        dict_domcateg['Tanggal Tiba'].append(data[i][j][15])
                        dict_domcateg['Tanggal Tolak'].append(data[i][j][26])
                        dict_domcateg['Muatan Tiba'].append(data[i][j][18])
                        dict_domcateg['Jml Muatan Tiba'].append(data[i][j][19])
                        dict_domcateg['Jns Muatan Tiba'].append(data[i][j][20])
                        dict_domcateg['Asal Kapal'].append(data[i][j][14])
                        dict_domcateg['Muatan Tolak'].append(depar_load[p])
                        dict_domcateg['Jml Muatan Tolak'].append(depar_num[p])
                        dict_domcateg['Jns Muatan Tolak'].append(depar_mu[p])
                        dict_domcateg['Tujuan Kapal'].append(data[i][j][25])
                        dict_domcateg['Jam Tolak'].append(data[i][j][27])
                        dict_domcateg['Kode Muatan Tiba'].append(data[i][j][21])
                        dict_domcateg['Kode Muatan Tolak'].append(data[i][j][32])
                        dict_domcateg['Kategori'].append(data[i][j][35])
                    elif p != 0:
                        dict_domcateg['Nomor'].append(None)
                        dict_domcateg['Kode Kapal'].append(None)
                        dict_domcateg['Nama Kapal'].append(None)
                        dict_domcateg['Keagenan'].append(None)
                        dict_domcateg['Bendera'].append(None)
                        dict_domcateg['GT'].append(None)
                        dict_domcateg['Trayek'].append(None)
                        dict_domcateg['Tanggal Tiba'].append(None)
                        dict_domcateg['Tanggal Tolak'].append(data[i][j][26])
                        dict_domcateg['Muatan Tiba'].append(None)
                        dict_domcateg['Jml Muatan Tiba'].append(None)
                        dict_domcateg['Jns Muatan Tiba'].append(None)
                        dict_domcateg['Asal Kapal'].append(None)
                        dict_domcateg['Muatan Tolak'].append(depar_load[p])
                        dict_domcateg['Jml Muatan Tolak'].append(depar_num[p])
                        dict_domcateg['Jns Muatan Tolak'].append(depar_mu[p])
                        dict_domcateg['Tujuan Kapal'].append(None)
                        dict_domcateg['Jam Tolak'].append(data[i][j][27])
                        dict_domcateg['Kode Muatan Tiba'].append(data[i][j][21])
                        dict_domcateg['Kode Muatan Tolak'].append(data[i][j][32])
                        dict_domcateg['Kategori'].append(data[i][j][35])

            elif ';' not in data[i][j][18] and ';' not in data[i][j][29]:
                dict_domcateg['Nomor'].append(j)
                dict_domcateg['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_domcateg['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_domcateg['Keagenan'].append(data[i][j][13])
                dict_domcateg['Bendera'].append(data[i][j][12])
                dict_domcateg['GT'].append(data[i][j][5])
                dict_domcateg['Trayek'].append('T')
                dict_domcateg['Tanggal Tiba'].append(data[i][j][15])
                dict_domcateg['Tanggal Tolak'].append(data[i][j][26])
                dict_domcateg['Muatan Tiba'].append(data[i][j][18])
                dict_domcateg['Jml Muatan Tiba'].append(data[i][j][19])
                dict_domcateg['Jns Muatan Tiba'].append(data[i][j][20])
                dict_domcateg['Asal Kapal'].append(data[i][j][14])
                dict_domcateg['Muatan Tolak'].append(data[i][j][29])
                dict_domcateg['Jml Muatan Tolak'].append(data[i][j][30])
                dict_domcateg['Jns Muatan Tolak'].append(data[i][j][31])
                dict_domcateg['Tujuan Kapal'].append(data[i][j][25])
                dict_domcateg['Jam Tolak'].append(data[i][j][27])
                dict_domcateg['Kode Muatan Tiba'].append(data[i][j][21])
                dict_domcateg['Kode Muatan Tolak'].append(data[i][j][32])
                dict_domcateg['Kategori'].append(data[i][j][35])

    #creating dataframe from dicts and sort it by dates
    predomcat = pd.DataFrame.from_dict(dict_domcateg)
    predomcat.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    #create new dataframes by certain categories
    domes = predomcat.loc[predomcat['Kategori'] == 'DOMESTIK']
    expor = predomcat.loc[predomcat['Kategori'] == 'EKSPOR']

    palm = domes.loc[(domes['Kode Muatan Tiba'].str.contains('(SWT)',case=False)) | (domes['Kode Muatan Tolak'].str.contains('(SWT)',case=False))]
    coal = domes.loc[(domes['Kode Muatan Tiba'].str.contains('(BABA)',case=False)) | (domes['Kode Muatan Tolak'].str.contains('(BABA)',case=False))]
    raco = domes.loc[(domes['Kode Muatan Tiba'].str.contains('(GEAR)',case=False)) | (domes['Kode Muatan Tolak'].str.contains('(GEAR)',case=False))]
    tuah = domes.loc[(domes['Kode Muatan Tiba'].str.contains('(BAPE)',case=False)) | (domes['Kode Muatan Tolak'].str.contains('(BAPE)',case=False))]
    coil = domes.loc[(domes['Kode Muatan Tiba'].str.contains('(CRIL)',case=False)) | (domes['Kode Muatan Tolak'].str.contains('(CRIL)',case=False))]
    atat = domes.loc[(domes['Kode Muatan Tiba'].str.contains('(ALBE)',case=False)) | (domes['Kode Muatan Tolak'].str.contains('(ALBE)',case=False))]
    fuel = domes.loc[(domes['Kode Muatan Tiba'].str.contains('(BBM)',case=False)) | (domes['Kode Muatan Tolak'].str.contains('(BBM)',case=False))]
    vehi = domes.loc[(domes['Kode Muatan Tiba'].str.contains('(KNDR)',case=False)) | (domes['Kode Muatan Tolak'].str.contains('(KNDR)',case=False))]
    wood = domes.loc[(domes['Kode Muatan Tiba'].str.contains('(KAYU)',case=False)) | (domes['Kode Muatan Tolak'].str.contains('(KAYU)',case=False))]
    sand = domes.loc[(domes['Kode Muatan Tiba'].str.contains('(TNH)',case=False)) | (domes['Kode Muatan Tolak'].str.contains('(TNH)',case=False))]
    mixx = domes.loc[(domes['Kode Muatan Tiba'].str.contains('(CMPR)',case=False)) | (domes['Kode Muatan Tolak'].str.contains('(CMPR)',case=False))]

    #minor refining for the output
    foutput = [domes,expor,palm,coal,raco,tuah,coil,atat,fuel,vehi,wood,sand,mixx]
    for funcs in range(len(foutput)):
        counter = 1
        for i in range(len(foutput[funcs])):
            if np.isnan(foutput[funcs].iat[i,0]) == False:
                foutput[funcs].iat[i,0] = counter
                counter+=1
            elif np.isnan(foutput[funcs].iat[i,0]) == True:
                foutput[funcs].iat[i,11] = None
                foutput[funcs].iat[i,12] = None
        foutput[funcs].drop(['Jam Tolak','Kode Muatan Tiba','Kode Muatan Tolak','Kategori'],axis=1,inplace=True)

    return foutput

#4 Port Clearance style report
def port_clr(data=None):
    #change writing style for initiating a dictionary
    dict_port = {'Nomor':[],'Kode SPB':[],'Nomor SPB':[],'Nomor Reg':[],'Kode Kapal':[],'Nama Kapal':[],
                 'Nama Nahkoda':[],'Bendera':[],'GT':[],'SIPI':[],'SIKPI':[],'SLO':[],'Asal Kapal':[],
                 'Tanggal Tiba':[],'Kru Kapal':[],'Tujuan Kapal':[],'Tanggal Tolak':[],'Muatan Tolak':[],
                 'Jml Muatan Tolak':[],'Jns Muatan Tolak':[],'Keagenan':[],'Jam Tolak':[]
                }

    #main process
    for i in range(len(data)):
        for j in range(1,len(data[i])):
            #nested conditioning to suit with the excel file content
            if ';' in data[i][j][29]:
                depar_load = data[i][j][29].split('; ')
                depar_num = data[i][j][30].split('; ')
                depar_mu = data[i][j][31].split('; ')

                for p in range(len(depar_load)):
                    if p == 0:
                        dict_port['Nomor'].append(j)
                        dict_port['Kode SPB'].append('T58')
                        dict_port['Nomor SPB'].append(data[i][j][1])
                        if data[i][j][2] != '--' and data[i][j][3] == '--':
                            dict_port['Nomor Reg'].append(data[i][j][2])
                        elif data[i][j][2] == '--' and data[i][j][3] != '--':
                            dict_port['Nomor Reg'].append(data[i][j][3])
                        dict_port['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                        dict_port['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                        dict_port['Nama Nahkoda'].append(data[i][j][9])
                        dict_port['Bendera'].append(data[i][j][12])
                        dict_port['GT'].append(data[i][j][5])
                        dict_port['SIPI'].append('--')
                        dict_port['SIKPI'].append('--')
                        dict_port['SLO'].append('--')
                        dict_port['Asal Kapal'].append(data[i][j][14])
                        dict_port['Tanggal Tiba'].append(data[i][j][15])
                        dict_port['Kru Kapal'].append(data[i][j][11])
                        dict_port['Tujuan Kapal'].append(data[i][j][25])
                        dict_port['Tanggal Tolak'].append(data[i][j][26])
                        dict_port['Muatan Tolak'].append(depar_load[p])
                        dict_port['Jml Muatan Tolak'].append(depar_num[p])
                        dict_port['Jns Muatan Tolak'].append(depar_mu[p])
                        dict_port['Keagenan'].append(data[i][j][13])
                        dict_port['Jam Tolak'].append(data[i][j][27])
                    elif p != 0:
                        dict_port['Nomor'].append(None)
                        dict_port['Kode SPB'].append(None)
                        dict_port['Nomor SPB'].append(None)
                        dict_port['Nomor Reg'].append(None)
                        dict_port['Kode Kapal'].append(None)
                        dict_port['Nama Kapal'].append(None)
                        dict_port['Nama Nahkoda'].append(None)
                        dict_port['Bendera'].append(None)
                        dict_port['GT'].append(None)
                        dict_port['SIPI'].append(None)
                        dict_port['SIKPI'].append(None)
                        dict_port['SLO'].append(None)
                        dict_port['Asal Kapal'].append(None)
                        dict_port['Tanggal Tiba'].append(None)
                        dict_port['Kru Kapal'].append(None)
                        dict_port['Tujuan Kapal'].append(None)
                        dict_port['Tanggal Tolak'].append(data[i][j][26])
                        dict_port['Muatan Tolak'].append(depar_load[p])
                        dict_port['Jml Muatan Tolak'].append(depar_num[p])
                        dict_port['Jns Muatan Tolak'].append(depar_mu[p])
                        dict_port['Keagenan'].append(None)
                        dict_port['Jam Tolak'].append(data[i][j][27])

            elif ';' not in data[i][j][29]:
                dict_port['Nomor'].append(j)
                dict_port['Kode SPB'].append('T58')
                dict_port['Nomor SPB'].append(data[i][j][1])
                if data[i][j][2] != '--' and data[i][j][3] == '--':
                    dict_port['Nomor Reg'].append(data[i][j][2])
                elif data[i][j][2] == '--' and data[i][j][3] != '--':
                    dict_port['Nomor Reg'].append(data[i][j][3])
                dict_port['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_port['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_port['Nama Nahkoda'].append(data[i][j][9])
                dict_port['Bendera'].append(data[i][j][12])
                dict_port['GT'].append(data[i][j][5])
                dict_port['SIPI'].append('--')
                dict_port['SIKPI'].append('--')
                dict_port['SLO'].append('--')
                dict_port['Asal Kapal'].append(data[i][j][14])
                dict_port['Tanggal Tiba'].append(data[i][j][15])
                dict_port['Kru Kapal'].append(data[i][j][11])
                dict_port['Tujuan Kapal'].append(data[i][j][25])
                dict_port['Tanggal Tolak'].append(data[i][j][26])
                dict_port['Muatan Tolak'].append(data[i][j][29])
                dict_port['Jml Muatan Tolak'].append(data[i][j][30])
                dict_port['Jns Muatan Tolak'].append(data[i][j][31])
                dict_port['Keagenan'].append(data[i][j][13])
                dict_port['Jam Tolak'].append(data[i][j][27])
    
    #convert dict into dataframe, sort by dates, and drop unused columns
    portclr = pd.DataFrame.from_dict(dict_port)
    portclr.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    portclr.drop(['Jam Tolak'],axis=1,inplace=True)

    counter = 1
    for i in range(len(portclr)):
        if np.isnan(portclr.iat[i,0]) == False:
            portclr.iat[i,0] = counter
            counter+=1
        elif np.isnan(portclr.iat[i,0]) == True:
            portclr.iat[i,11] = None
            portclr.iat[i,12] = None
    
    return portclr

#5 Loads Counter
def goods(data=None):
    #function for summarizing shiploads, splitted to two categories: Domestic & Export
    dogarname, dogarnum, dogarmea, dodepname, dodepnum, dodepmea = [],[],[],[],[],[]
    exdepname, exdepnum, exdepmea = [],[],[],[],[],[]

    #inserting shipload's names and numbers
    for i in range(len(data)):
        for j in range(1,len(data[i])):
            #inserting data in domestic operation
            if data[i][j][35] == 'DOMESTIK':
                if ';' in data[i][j][18] and ';' in data[i][j][29]:
                    garr_temp,narr_temp,marr_temp = data[i][j][18].split('; '),data[i][j][19].split('; '),data[i][j][20].split('; ')
                    gdep_temp,ndep_temp,mdep_temp = data[i][j][29].split('; '),data[i][j][30].split('; '),data[i][j][31].split('; ')

                    dogarname.extend(garr_temp)
                    dogarnum.extend(narr_temp)
                    dogarmea.extend(marr_temp)

                    dodepname.extend(gdep_temp)
                    dodepnum.extend(ndep_temp)
                    dodepmea.extend(mdep_temp)

                elif ';' in data[i][j][18] and ';' not in data[i][j][29]:
                    garr_temp,narr_temp,marr_temp = data[i][j][18].split('; '),data[i][j][19].split('; '),data[i][j][20].split('; ')

                    dogarname.extend(garr_temp)
                    dogarnum.extend(narr_temp)
                    dogarmea.extend(marr_temp)

                elif ';' not in data[i][j][18] and ';' in data[i][j][29]:
                    gdep_temp,ndep_temp,mdep_temp = data[i][j][29].split('; '),data[i][j][30].split('; '),data[i][j][31].split('; ')

                    dodepname.extend(gdep_temp)
                    dodepnum.extend(ndep_temp)
                    dodepmea.extend(mdep_temp)

                elif ';' not in data[i][j][18] and ';' not in data[i][j][29]:
                    dogarname.append(data[i][j][18])
                    dogarnum.append(data[i][j][19])
                    dogarmea.append(data[i][j][20])

                    dodepname.append(data[i][j][29])
                    dodepnum.append(data[i][j][30])
                    dodepmea.append(data[i][j][31])
                    
            #inserting data for exporting operation
            elif data[i][j][35] == 'EKSPOR':
                if ';' in data[i][j][29]:
                    gdep_temp,ndep_temp,mdep_temp = data[i][j][29].split('; '),data[i][j][30].split('; '),data[i][j][31].split('; ')

                    exdepname.extend(gdep_temp)
                    exdepnum.extend(ndep_temp)
                    exdepmea.extend(mdep_temp)

                elif ';' not in data[i][j][29]:
                    exdepname.append(data[i][j][29])
                    exdepnum.append(data[i][j][30])
                    exdepmea.append(data[i][j][31])
                
                else:
                    pass
                
    #data cleaning from certain conditions ie. nan and data type
    dogarname = [x for x in dogarname if x != 'NIHIL']
    dogarnum = list(map(str,[x for x in dogarnum]))
    dogarnum = list(map(float,[parseNumber(x) for x in dogarnum if x != '--']))
    dogarmea = [x for x in dogarmea if x != '--']

    dodepname = [x for x in dodepname if x != 'NIHIL']
    dodepnum = list(map(str,[x for x in dodepnum]))
    dodepnum = list(map(float,[parseNumber(x) for x in dodepnum if x != '--']))
    dodepmea = [x for x in dodepmea if x != '--']

    exdepname = [x for x in exdepname if x != 'NIHIL']
    exdepnum = list(map(str,[x for x in exdepnum]))
    exdepnum = list(map(float,[parseNumber(x) for x in exdepnum if x != '--']))
    exdepmea = [x for x in exdepmea if x != '--']
    
    #calculating shipload's summary
    domarr = pd.DataFrame(list(zip(dogarname,dogarnum,dogarmea)), columns = ['Nama', 'Jumlah','Satuan'])
    domdep = pd.DataFrame(list(zip(dodepname,dodepnum,dodepmea)), columns = ['Nama', 'Jumlah','Satuan'])

    expdep = pd.DataFrame(list(zip(exdepname,exdepnum,exdepmea)), columns = ['Nama', 'Jumlah','Satuan'])

    goardom = domarr.groupby(['Nama','Satuan'],as_index=False).sum('Jumlah')
    godedom = domdep.groupby(['Nama','Satuan'],as_index=False).sum('Jumlah')
    domarr,domdep = goardom.values.tolist(),godedom.values.tolist()

    godeexp = expdep.groupby(['Nama','Satuan'],as_index=False).sum('Jumlah')
    expdep = godeexp.values.tolist()
    
    #rounding numbers to 3 decimal places
    for glist in [domarr,domdep,expdep]:
        for i in range(len(glist)):
            peek = str(glist[i][2])
            if len(peek[peek.find('.')+1:]) > 3:
                glist[i][2] = float(round(glist[i][2],3))
            elif len(peek[peek.find('.')+1:]) <= 3:
                pass
            else:
                pass
            
    #collecting shipload's names
    sumcag = [glist[i][0]+'-'+glist[i][1] for glist in [domarr,domdep] for i in range(len(glist))]
    sumcag = np.unique(sumcag).tolist()
    
    #writing summary, first for domestic operation
    sumnam,sumar,sumde = [],[],[]
    for i in range(len(sumcag)):
        gname,gmea = sumcag[i].split('-')
        sumnam.append(gname)
        try:
            j = nestedlist_rootindex(domarr, gname, gmea)
            sumar.append(str(domarr[j][2])+' '+domarr[j][1])
        except TypeError:
            sumar.append('--')

        try:
            j = nestedlist_rootindex(domdep, gname, gmea)
            sumde.append(str(domdep[j][2])+' '+domdep[j][1])
        except TypeError:
            sumde.append('--')
    
    #second, for exporting operation
    expna,expar,expde = [expdep[0][0]],['--'],[str(expdep[0][2])+' '+expdep[0][1]]
    
    #finalizing format for summaries
    goods_dom = pd.DataFrame(list(zip(sumnam,sumar,sumde)), columns = ['Nama', 'Bongkar','Muat'])
    goods_exp = pd.DataFrame(list(zip(expna,expar,expde)), columns = ['Nama', 'Bongkar','Muat'])

    return goods_dom,goods_exp

#master function
def main():
    datadf = read_excel_file(filename='data input.xlsx')
    sibk,sibg = sib_based(data=datadf)
    gabungan,bunyu,albunyu,nihil = tkii_based(data=datadf)
    dom,exp,swt,baba,gear,bape,cril,albe,fuel,kndr,wood,sand,cmpr = domexp_based(data=datadf)
    spb = port_clr(data=datadf)
    port_clr(data=datadf)
    domsum,expsum = goods(data=datadf)

    filewriter = pd.ExcelWriter('Untuk Hasil Olah Data.xlsx')

    sibk.to_excel(filewriter,'SIB Kecil',index=False)
    sibg.to_excel(filewriter,'SIB Besar',index=False)
    gabungan.to_excel(filewriter,'TK.II UPT',index=False)
    dom.to_excel(filewriter,'Domestik',index=False)
    exp.to_excel(filewriter,'Ekspor',index=False)
    bunyu.to_excel(filewriter,'Bunyu',index=False)
    albunyu.to_excel(filewriter,'Bunyu AL',index=False)
    nihil.to_excel(filewriter,'Tanpa Muatan',index=False)
    swt.to_excel(filewriter,'Sawit',index=False)
    baba.to_excel(filewriter,'Batubara',index=False)
    gear.to_excel(filewriter,'General Cargo',index=False)
    bape.to_excel(filewriter,'Batu Pecah',index=False)
    cril.to_excel(filewriter,'Crude Oil',index=False)
    albe.to_excel(filewriter,'Alat Berat',index=False)
    fuel.to_excel(filewriter,'BBM',index=False)
    kndr.to_excel(filewriter,'Mobil',index=False)
    wood.to_excel(filewriter,'Kayu',index=False)
    sand.to_excel(filewriter,'Tanah',index=False)
    cmpr.to_excel(filewriter,'Campuran',index=False)
    spb.to_excel(filewriter,'SPB',index=False)
    domsum.to_excel(filewriter,'Ringkasan Muatan',index=False,startcol=0)
    expsum.to_excel(filewriter,'Ringkasan Muatan',index=False,startcol=4)

    filewriter.save()
    
if __name__ == '__main__':
    main()

''' ------------------------------End of The Code Writing---------------------------'''
