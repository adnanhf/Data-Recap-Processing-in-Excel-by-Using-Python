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

''' ----------------------Code Writing at 7th September 2021----------------------'''

# Importing library
import pandas as pd

# Read data function
def read_excel_file(filename='None',col_title=4):
    dfname=pd.read_excel(filename,sheet_name=None,header=col_title)
    #printing number of sheet in an excel file
    print('terdapat',len(dfname),'sheet di dalam file ini')
    
    #declare blank list for converting Pandas dataframe to list
    nested_list=[]

    for i in dfname:
        #delete any empty cells, based on row
        dfname[i].dropna(axis=0,how='any',inplace=True)
        #converting dataframe into list
        child_list = [dfname[i].columns.values.tolist()] + dfname[i].values.tolist()
        #printing number of row from each sheet
        print('Sheet',i,'ada',len(dfname[i]),'baris')
        nested_list.append(child_list)

    return nested_list

def sib_based(data=None):
    
    #declare dictionaries for output
    dict_sibkecil = {
        'Nama Kapal':[],
        'Nama Nahkoda':[],
        'GT':[],
        'NT':[],
        'Tanda Selar':[],
        'Tempat Pendaftaran':[],
        'Tanggal Tiba':[],
        'Asal Kapal':[],
        'Kode Muatan Tiba':[],
        'Tanggal Tolak':[],
        'Jam Tolak':[],
        'Tujuan Kapal':[],
        'Kode Muatan Tolak':[],
        'Keagenan':[],
        'Keterangan':[],
    }

    dict_sibgede = {
        'Nama Kapal':[],
        'Bendera':[],
        'Nama Nahkoda':[],
        'Tempat Pendaftaran':[],
        'GT':[],
        'NT':[],
        'Tanggal Tiba':[],
        'Asal Kapal':[],
        'Kode Muatan Tiba':[],
        'Tanggal Tolak':[],
        'Jam Tolak':[],
        'Tujuan Kapal':[],
        'Kode Muatan Tolak':[],
        'Keagenan':[],
        'Keterangan':[],
    }

    #loop through list of excel data and append it to the dictionaries
    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if data[i][j][5] <= 500:
               dict_sibkecil['Nama Kapal'].append(data[i][j][4])
               dict_sibkecil['Nama Nahkoda'].append(data[i][j][9])
               dict_sibkecil['GT'].append(data[i][j][5])
               dict_sibkecil['NT'].append(data[i][j][6])
               dict_sibkecil['Tanda Selar'].append(data[i][j][7])
               dict_sibkecil['Tempat Pendaftaran'].append(data[i][j][8])
               dict_sibkecil['Tanggal Tiba'].append(data[i][j][15])
               dict_sibkecil['Asal Kapal'].append(data[i][j][14])
               dict_sibkecil['Kode Muatan Tiba'].append(data[i][j][17])
               dict_sibkecil['Tanggal Tolak'].append(data[i][j][22])
               dict_sibkecil['Jam Tolak'].append(data[i][j][23])
               dict_sibkecil['Tujuan Kapal'].append(data[i][j][21])
               dict_sibkecil['Kode Muatan Tolak'].append(data[i][j][24])
               dict_sibkecil['Keagenan'].append(data[i][j][13])
               dict_sibkecil['Keterangan'].append(data[i][j][26])
            elif data[i][j][5] > 500:
               dict_sibgede['Nama Kapal'].append(data[i][j][4])
               dict_sibgede['Bendera'].append(data[i][j][12])
               dict_sibgede['Nama Nahkoda'].append(data[i][j][9])
               dict_sibgede['Tempat Pendaftaran'].append(data[i][j][8])
               dict_sibgede['GT'].append(data[i][j][5])
               dict_sibgede['NT'].append(data[i][j][6])
               dict_sibgede['Tanggal Tiba'].append(data[i][j][15])
               dict_sibgede['Asal Kapal'].append(data[i][j][14])
               dict_sibgede['Kode Muatan Tiba'].append(data[i][j][17])
               dict_sibgede['Tanggal Tolak'].append(data[i][j][22])
               dict_sibgede['Jam Tolak'].append(data[i][j][23])
               dict_sibgede['Tujuan Kapal'].append(data[i][j][21])
               dict_sibgede['Kode Muatan Tolak'].append(data[i][j][24])
               dict_sibgede['Keagenan'].append(data[i][j][13])
               dict_sibgede['Keterangan'].append(data[i][j][26])
    
    #convert dictionaries into Pandas dataframe
    sib_kecil = pd.DataFrame.from_dict(dict_sibkecil)
    sib_gede = pd.DataFrame.from_dict(dict_sibgede)
    #sorting data in Pandas dataframe, based on date
    sib_kecil.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    sib_gede.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return sib_kecil,sib_gede

def combine_data(data=None):
    dict_combined = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Bendera':[],
        'Keagenan':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Jam Tiba':[],
        'Asal Kapal':[],
        'Tanggal Tambat':[],
        'Jam Tambat':[],
        'Tanggal Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
        'Muatan Tiba':[],
        'Muatan Tolak':[],
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if ';' in data[i][j][18]:
                arr_load = data[i][j][18].split('; ')

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
                        dict_combined['Tanggal Tolak'].append(data[i][j][24])
                        dict_combined['Tujuan Kapal'].append(data[i][j][23])
                        dict_combined['Jam Tolak'].append(data[i][j][25])
                        dict_combined['Muatan Tiba'].append(arr_load[i])
                        dict_combined['Muatan Tolak'].append(data[i][j][27])                        
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
                        dict_combined['Tanggal Tolak'].append(data[i][j][24])
                        dict_combined['Tujuan Kapal'].append(None)
                        dict_combined['Jam Tolak'].append(data[i][j][25])
                        dict_combined['Muatan Tiba'].append(arr_load[p])
                        dict_combined['Muatan Tolak'].append(None)

            elif ';' in data[i][j][27]:
                depar_load = data[i][j][27].split('; ')

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
                        dict_combined['Tanggal Tolak'].append(data[i][j][24])
                        dict_combined['Tujuan Kapal'].append(data[i][j][23])
                        dict_combined['Jam Tolak'].append(data[i][j][25])
                        dict_combined['Muatan Tiba'].append(data[i][j][18])
                        dict_combined['Muatan Tolak'].append(depar_load[i])
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
                        dict_combined['Tanggal Tolak'].append(data[i][j][24])
                        dict_combined['Tujuan Kapal'].append(None)
                        dict_combined['Jam Tolak'].append(data[i][j][25])
                        dict_combined['Muatan Tiba'].append(None)
                        dict_combined['Muatan Tolak'].append(depar_load[p])

            elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
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
                dict_combined['Tanggal Tolak'].append(data[i][j][24])
                dict_combined['Tujuan Kapal'].append(data[i][j][23])
                dict_combined['Jam Tolak'].append(data[i][j][25])
                dict_combined['Muatan Tiba'].append(data[i][j][18])
                dict_combined['Muatan Tolak'].append(data[i][j][27])

    combined = pd.DataFrame.from_dict(dict_combined)
    combined.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(combined)):
        if np.isnan(combined.iat[i,0]) == False:
            combined.iat[i,0] = counter
            counter+=1
        elif np.isnan(combined.iat[i,0]) == True:
            pass
    
    return combined

def dom_categ(data=None):
    dict_domcateg = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Keagenan':[],
        'Bendera':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Tanggal Tolak':[],
        'Muatan Tiba':[],
        'Asal Kapal':[],
        'Muatan Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if ';' in data[i][j][18]:
                arr_load = data[i][j][18].split('; ')

                for p in range(len(arr_load)):
                    if p == 0:
                        dict_domcateg['Nomor'].append(j)
                        dict_domcateg['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                        dict_domcateg['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                        dict_domcateg['Keagenan'].append(data[i][j][13])
                        dict_domcateg['Bendera'].append(data[i][j][12])
                        dict_domcateg['GT'].append(data[i][j][5])
                        dict_domcateg['Tanggal Tiba'].append(data[i][j][15])
                        dict_domcateg['Tanggal Tolak'].append(data[i][j][24])
                        dict_domcateg['Muatan Tiba'].append(arr_load[p])
                        dict_domcateg['Asal Kapal'].append(data[i][j][14])
                        dict_domcateg['Muatan Tolak'].append(data[i][j][27])
                        dict_domcateg['Tujuan Kapal'].append(data[i][j][23])
                        dict_domcateg['Jam Tolak'].append(data[i][j][25])
                    elif p != 0:
                        dict_domcateg['Nomor'].append(None)
                        dict_domcateg['Kode Kapal'].append(None)
                        dict_domcateg['Nama Kapal'].append(None)
                        dict_domcateg['Keagenan'].append(None)
                        dict_domcateg['Bendera'].append(None)
                        dict_domcateg['GT'].append(None)
                        dict_domcateg['Tanggal Tiba'].append(None)
                        dict_domcateg['Tanggal Tolak'].append(data[i][j][24])
                        dict_domcateg['Muatan Tiba'].append(arr_load[p])
                        dict_domcateg['Asal Kapal'].append(None)
                        dict_domcateg['Muatan Tolak'].append(None)
                        dict_domcateg['Tujuan Kapal'].append(None)
                        dict_domcateg['Jam Tolak'].append(data[i][j][25])

            elif ';' in data[i][j][27]:
                depar_load = data[i][j][27].split('; ')

                for p in range(len(depar_load)):
                    if p == 0:
                        dict_domcateg['Nomor'].append(j)
                        dict_domcateg['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                        dict_domcateg['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                        dict_domcateg['Keagenan'].append(data[i][j][13])
                        dict_domcateg['Bendera'].append(data[i][j][12])
                        dict_domcateg['GT'].append(data[i][j][5])
                        dict_domcateg['Tanggal Tiba'].append(data[i][j][15])
                        dict_domcateg['Tanggal Tolak'].append(data[i][j][24])
                        dict_domcateg['Muatan Tiba'].append(data[i][j][18])
                        dict_domcateg['Asal Kapal'].append(data[i][j][14])
                        dict_domcateg['Muatan Tolak'].append(depar_load[p])
                        dict_domcateg['Tujuan Kapal'].append(data[i][j][23])
                        dict_domcateg['Jam Tolak'].append(data[i][j][25])
                    elif p != 0:
                        dict_domcateg['Nomor'].append(None)
                        dict_domcateg['Kode Kapal'].append(None)
                        dict_domcateg['Nama Kapal'].append(None)
                        dict_domcateg['Keagenan'].append(None)
                        dict_domcateg['Bendera'].append(None)
                        dict_domcateg['GT'].append(None)
                        dict_domcateg['Tanggal Tiba'].append(None)
                        dict_domcateg['Tanggal Tolak'].append(data[i][j][24])
                        dict_domcateg['Muatan Tiba'].append(None)
                        dict_domcateg['Asal Kapal'].append(None)
                        dict_domcateg['Muatan Tolak'].append(depar_load[p])
                        dict_domcateg['Tujuan Kapal'].append(None)
                        dict_domcateg['Jam Tolak'].append(data[i][j][25])

            elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                dict_domcateg['Nomor'].append(j)
                dict_domcateg['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_domcateg['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_domcateg['Keagenan'].append(data[i][j][13])
                dict_domcateg['Bendera'].append(data[i][j][12])
                dict_domcateg['GT'].append(data[i][j][5])
                dict_domcateg['Tanggal Tiba'].append(data[i][j][15])
                dict_domcateg['Tanggal Tolak'].append(data[i][j][24])
                dict_domcateg['Muatan Tiba'].append(data[i][j][18])
                dict_domcateg['Asal Kapal'].append(data[i][j][14])
                dict_domcateg['Muatan Tolak'].append(data[i][j][27])
                dict_domcateg['Tujuan Kapal'].append(data[i][j][23])
                dict_domcateg['Jam Tolak'].append(data[i][j][25])

    domcateg = pd.DataFrame.from_dict(dict_domcateg)
    domcateg.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(domcateg)):
        if np.isnan(domcateg.iat[i,0]) == False:
            domcateg.iat[i,0] = counter
            counter+=1
        elif np.isnan(domcateg.iat[i,0]) == True:
            pass

    return domcateg

def bunyu(data=None):
    dict_bunyu = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Bendera':[],
        'Keagenan':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Jam Tiba':[],
        'Asal Kapal':[],
        'Tanggal Tambat':[],
        'Jam Tambat':[],
        'Tanggal Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
        'Muatan Tiba':[],
        'Muatan Tolak':[],
    }

    for i in range(len(data)-1):
        for j in range(1,len(data[i])):
            if ';' in data[i][j][18]:
                arr_load = data[i][j][18].split('; ')

                for p in range(len(arr_load)):
                    if p == 0:
                        dict_bunyu['Nomor'].append(j)
                        dict_bunyu['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                        dict_bunyu['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                        dict_bunyu['Bendera'].append(data[i][j][12])
                        dict_bunyu['Keagenan'].append(data[i][j][13])
                        dict_bunyu['GT'].append(data[i][j][5])
                        dict_bunyu['Tanggal Tiba'].append(data[i][j][15])
                        dict_bunyu['Jam Tiba'].append(data[i][j][16])
                        dict_bunyu['Asal Kapal'].append(data[i][j][14])
                        dict_bunyu['Tanggal Tambat'].append(data[i][j][15])
                        dict_bunyu['Jam Tambat'].append(data[i][j][16])
                        dict_bunyu['Tanggal Tolak'].append(data[i][j][24])
                        dict_bunyu['Tujuan Kapal'].append(data[i][j][23])
                        dict_bunyu['Jam Tolak'].append(data[i][j][25])
                        dict_bunyu['Muatan Tiba'].append(arr_load[p])
                        dict_bunyu['Muatan Tolak'].append(data[i][j][27])
                    elif p != 0:
                        dict_bunyu['Nomor'].append(None)
                        dict_bunyu['Kode Kapal'].append(None)
                        dict_bunyu['Nama Kapal'].append(None)
                        dict_bunyu['Bendera'].append(None)
                        dict_bunyu['Keagenan'].append(None)
                        dict_bunyu['GT'].append(None)
                        dict_bunyu['Tanggal Tiba'].append(None)
                        dict_bunyu['Jam Tiba'].append(None)
                        dict_bunyu['Asal Kapal'].append(None)
                        dict_bunyu['Tanggal Tambat'].append(None)
                        dict_bunyu['Jam Tambat'].append(None)
                        dict_bunyu['Tanggal Tolak'].append(data[i][j][24])
                        dict_bunyu['Tujuan Kapal'].append(None)
                        dict_bunyu['Jam Tolak'].append(data[i][j][25])
                        dict_bunyu['Muatan Tiba'].append(arr_load[p])
                        dict_bunyu['Muatan Tolak'].append(None)

            elif ';' in data[i][j][27]:
                depar_load = data[i][j][27].split('; ')
                depar_load[-1] = depar_load[-1][:depar_load[-1].find('(')-1]

                for p in range(len(depar_load)):
                    if p == 0:
                        dict_bunyu['Nomor'].append(j)
                        dict_bunyu['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                        dict_bunyu['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                        dict_bunyu['Bendera'].append(data[i][j][12])
                        dict_bunyu['Keagenan'].append(data[i][j][13])
                        dict_bunyu['GT'].append(data[i][j][5])
                        dict_bunyu['Tanggal Tiba'].append(data[i][j][15])
                        dict_bunyu['Jam Tiba'].append(data[i][j][16])
                        dict_bunyu['Asal Kapal'].append(data[i][j][14])
                        dict_bunyu['Tanggal Tambat'].append(data[i][j][15])
                        dict_bunyu['Jam Tambat'].append(data[i][j][16])
                        dict_bunyu['Tanggal Tolak'].append(data[i][j][24])
                        dict_bunyu['Tujuan Kapal'].append(data[i][j][23])
                        dict_bunyu['Jam Tolak'].append(data[i][j][25])
                        dict_bunyu['Muatan Tiba'].append(data[i][j][18])
                        dict_bunyu['Muatan Tolak'].append(depar_load[p])
                    elif p != 0:
                        dict_bunyu['Nomor'].append(None)
                        dict_bunyu['Kode Kapal'].append(None)
                        dict_bunyu['Nama Kapal'].append(None)
                        dict_bunyu['Bendera'].append(None)
                        dict_bunyu['Keagenan'].append(None)
                        dict_bunyu['GT'].append(None)
                        dict_bunyu['Tanggal Tiba'].append(None)
                        dict_bunyu['Jam Tiba'].append(None)
                        dict_bunyu['Asal Kapal'].append(None)
                        dict_bunyu['Tanggal Tambat'].append(None)
                        dict_bunyu['Jam Tambat'].append(None)
                        dict_bunyu['Tanggal Tolak'].append(data[i][j][24])
                        dict_bunyu['Tujuan Kapal'].append(None)
                        dict_bunyu['Jam Tolak'].append(data[i][j][25])
                        dict_bunyu['Muatan Tiba'].append(None)
                        dict_bunyu['Muatan Tolak'].append(depar_load[p])

            elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                dict_bunyu['Nomor'].append(j)
                dict_bunyu['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_bunyu['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_bunyu['Bendera'].append(data[i][j][12])
                dict_bunyu['Keagenan'].append(data[i][j][13])
                dict_bunyu['GT'].append(data[i][j][5])
                dict_bunyu['Tanggal Tiba'].append(data[i][j][15])
                dict_bunyu['Jam Tiba'].append(data[i][j][16])
                dict_bunyu['Asal Kapal'].append(data[i][j][14])
                dict_bunyu['Tanggal Tambat'].append(data[i][j][15])
                dict_bunyu['Jam Tambat'].append(data[i][j][16])
                dict_bunyu['Tanggal Tolak'].append(data[i][j][24])
                dict_bunyu['Tujuan Kapal'].append(data[i][j][23])
                dict_bunyu['Jam Tolak'].append(data[i][j][25])
                dict_bunyu['Muatan Tiba'].append(data[i][j][18])
                dict_bunyu['Muatan Tolak'].append(data[i][j][27])
    
    bunyuisland = pd.DataFrame.from_dict(dict_bunyu)
    bunyuisland.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(bunyuisland)):
        if np.isnan(bunyuisland.iat[i,0]) == False:
            bunyuisland.iat[i,0] = counter
            counter+=1
        elif np.isnan(bunyuisland.iat[i,0]) == True:
            pass

    return bunyuisland

def bunyu_al(data=None):
    dict_albunyu = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Bendera':[],
        'Keagenan':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Jam Tiba':[],
        'Asal Kapal':[],
        'Tanggal Tambat':[],
        'Jam Tambat':[],
        'Tanggal Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
        'Muatan Tiba':[],
        'Muatan Tolak':[],
    }

    for i in range(len(data)-1):
        for j in range(1,len(data[i])):
            if data[i][j][22] == 'BUNYU':
                if ';' in data[i][j][18]:
                    arr_load = data[i][j][18].split('; ')

                    for p in range(len(arr_load)):
                        if p == 0:
                            dict_albunyu['Nomor'].append(j)
                            dict_albunyu['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_albunyu['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_albunyu['Bendera'].append(data[i][j][12])
                            dict_albunyu['Keagenan'].append(data[i][j][13])
                            dict_albunyu['GT'].append(data[i][j][5])
                            dict_albunyu['Tanggal Tiba'].append(data[i][j][15])
                            dict_albunyu['Jam Tiba'].append(data[i][j][16])
                            dict_albunyu['Asal Kapal'].append(data[i][j][14])
                            dict_albunyu['Tanggal Tambat'].append(data[i][j][15])
                            dict_albunyu['Jam Tambat'].append(data[i][j][16])
                            dict_albunyu['Tanggal Tolak'].append(data[i][j][24])
                            dict_albunyu['Tujuan Kapal'].append(data[i][j][23])
                            dict_albunyu['Jam Tolak'].append(data[i][j][25])
                            dict_albunyu['Muatan Tiba'].append(arr_load[p])
                            dict_albunyu['Muatan Tolak'].append(data[i][j][27])
                        elif p != 0:
                            dict_albunyu['Nomor'].append(None)
                            dict_albunyu['Kode Kapal'].append(None)
                            dict_albunyu['Nama Kapal'].append(None)
                            dict_albunyu['Bendera'].append(None)
                            dict_albunyu['Keagenan'].append(None)
                            dict_albunyu['GT'].append(None)
                            dict_albunyu['Tanggal Tiba'].append(None)
                            dict_albunyu['Jam Tiba'].append(None)
                            dict_albunyu['Asal Kapal'].append(None)
                            dict_albunyu['Tanggal Tambat'].append(None)
                            dict_albunyu['Jam Tambat'].append(None)
                            dict_albunyu['Tanggal Tolak'].append(data[i][j][24])
                            dict_albunyu['Tujuan Kapal'].append(None)
                            dict_albunyu['Jam Tolak'].append(data[i][j][25])
                            dict_albunyu['Muatan Tiba'].append(arr_load[p])
                            dict_albunyu['Muatan Tolak'].append(None)

                elif ';' in data[i][j][27]:
                    depar_load = data[i][j][27].split('; ')

                    for p in range(len(depar_load)):
                        if p == 0:
                            dict_albunyu['Nomor'].append(j)
                            dict_albunyu['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_albunyu['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_albunyu['Bendera'].append(data[i][j][12])
                            dict_albunyu['Keagenan'].append(data[i][j][13])
                            dict_albunyu['GT'].append(data[i][j][5])
                            dict_albunyu['Tanggal Tiba'].append(data[i][j][15])
                            dict_albunyu['Jam Tiba'].append(data[i][j][16])
                            dict_albunyu['Asal Kapal'].append(data[i][j][14])
                            dict_albunyu['Tanggal Tambat'].append(data[i][j][15])
                            dict_albunyu['Jam Tambat'].append(data[i][j][16])
                            dict_albunyu['Tanggal Tolak'].append(data[i][j][24])
                            dict_albunyu['Tujuan Kapal'].append(data[i][j][23])
                            dict_albunyu['Jam Tolak'].append(data[i][j][25])
                            dict_albunyu['Muatan Tiba'].append(data[i][j][18])
                            dict_albunyu['Muatan Tolak'].append(depar_load[p])
                        elif p != 0:
                            dict_albunyu['Nomor'].append(None)
                            dict_albunyu['Kode Kapal'].append(None)
                            dict_albunyu['Nama Kapal'].append(None)
                            dict_albunyu['Bendera'].append(None)
                            dict_albunyu['Keagenan'].append(None)
                            dict_albunyu['GT'].append(None)
                            dict_albunyu['Tanggal Tiba'].append(None)
                            dict_albunyu['Jam Tiba'].append(None)
                            dict_albunyu['Asal Kapal'].append(None)
                            dict_albunyu['Tanggal Tambat'].append(None)
                            dict_albunyu['Jam Tambat'].append(None)
                            dict_albunyu['Tanggal Tolak'].append(data[i][j][24])
                            dict_albunyu['Tujuan Kapal'].append(None)
                            dict_albunyu['Jam Tolak'].append(data[i][j][25])
                            dict_albunyu['Muatan Tiba'].append(None)
                            dict_albunyu['Muatan Tolak'].append(depar_load[p])

                elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                    dict_albunyu['Nomor'].append(j)
                    dict_albunyu['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                    dict_albunyu['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                    dict_albunyu['Bendera'].append(data[i][j][12])
                    dict_albunyu['Keagenan'].append(data[i][j][13])
                    dict_albunyu['GT'].append(data[i][j][5])
                    dict_albunyu['Tanggal Tiba'].append(data[i][j][15])
                    dict_albunyu['Jam Tiba'].append(data[i][j][16])
                    dict_albunyu['Asal Kapal'].append(data[i][j][14])
                    dict_albunyu['Tanggal Tambat'].append(data[i][j][15])
                    dict_albunyu['Jam Tambat'].append(data[i][j][16])
                    dict_albunyu['Tanggal Tolak'].append(data[i][j][24])
                    dict_albunyu['Tujuan Kapal'].append(data[i][j][23])
                    dict_albunyu['Jam Tolak'].append(data[i][j][25])
                    dict_albunyu['Muatan Tiba'].append(data[i][j][18])
                    dict_albunyu['Muatan Tolak'].append(data[i][j][27])
    
    albunyu = pd.DataFrame.from_dict(dict_albunyu)
    albunyu.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(albunyu)):
        if np.isnan(albunyu.iat[i,0]) == False:
            albunyu.iat[i,0] = counter
            counter+=1
        elif np.isnan(albunyu.iat[i,0]) == True:
            pass

    return albunyu

def port_clr(data=None):
    dict_port = {
        'Nomor':[],
        'Kode SPB':[],
        'Nomor SPB':[],
        'Nomor Reg':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Nama Nahkoda':[],
        'Bendera':[],
        'GT':[],
        'SIPI':[],
        'SIKPI':[],
        'SLO':[],
        'Asal Kapal':[],
        'Tanggal Tiba':[],
        'Kru Kapal':[],
        'Tujuan Kapal':[],
        'Tanggal Tolak':[],
        'Muatan Tolak':[],
        'Keagenan':[],
        'Jam Tolak':[]
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if ';' in data[i][j][27]:
                depar_load = data[i][j][27].split('; ')

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
                        dict_port['Tujuan Kapal'].append(data[i][j][23])
                        dict_port['Tanggal Tolak'].append(data[i][j][24])
                        dict_port['Muatan Tolak'].append(depar_load[p])
                        dict_port['Keagenan'].append(data[i][j][13])
                        dict_port['Jam Tolak'].append(data[i][j][25])
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
                        dict_port['Tanggal Tolak'].append(data[i][j][24])
                        dict_port['Muatan Tolak'].append(depar_load[p])
                        dict_port['Keagenan'].append(None)
                        dict_port['Jam Tolak'].append(data[i][j][25])

            elif ';' not in data[i][j][27]:
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
                dict_port['Tujuan Kapal'].append(data[i][j][23])
                dict_port['Tanggal Tolak'].append(data[i][j][24])
                dict_port['Muatan Tolak'].append(data[i][j][27])
                dict_port['Keagenan'].append(data[i][j][13])
                dict_port['Jam Tolak'].append(data[i][j][25])
    
    portclr = pd.DataFrame.from_dict(dict_port)
    portclr.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(portclr)):
        if np.isnan(portclr.iat[i,0]) == False:
            portclr.iat[i,0] = counter
            counter+=1
        elif np.isnan(portclr.iat[i,0]) == True:
            pass

    return portclr

def wo_loads(data=None):
    dict_loads = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Bendera':[],
        'Keagenan':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Jam Tiba':[],
        'Asal Kapal':[],
        'Tanggal Tambat':[],
        'Jam Tambat':[],
        'Tanggal Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
    }

    ship_type = ['TB','OB','TK','BG']

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if data[i][j][4][:data[i][j][4].find('. ')] in ship_type:
                dict_loads['Nomor'].append(j)
                dict_loads['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_loads['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_loads['Bendera'].append(data[i][j][12])
                dict_loads['Keagenan'].append(data[i][j][13])
                dict_loads['GT'].append(data[i][j][5])
                dict_loads['Tanggal Tiba'].append(data[i][j][15])
                dict_loads['Jam Tiba'].append(data[i][j][16])
                dict_loads['Asal Kapal'].append(data[i][j][14])
                dict_loads['Tanggal Tambat'].append(data[i][j][15])
                dict_loads['Jam Tambat'].append(data[i][j][16])
                dict_loads['Tanggal Tolak'].append(data[i][j][24])
                dict_loads['Tujuan Kapal'].append(data[i][j][23])
                dict_loads['Jam Tolak'].append(data[i][j][25])
            elif data[i][j][18] == 'NIHIL' and data[i][j][27] == 'NIHIL':
                dict_loads['Nomor'].append(j)
                dict_loads['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_loads['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_loads['Bendera'].append(data[i][j][12])
                dict_loads['Keagenan'].append(data[i][j][13])
                dict_loads['GT'].append(data[i][j][5])
                dict_loads['Tanggal Tiba'].append(data[i][j][15])
                dict_loads['Jam Tiba'].append(data[i][j][16])
                dict_loads['Asal Kapal'].append(data[i][j][14])
                dict_loads['Tanggal Tambat'].append(data[i][j][15])
                dict_loads['Jam Tambat'].append(data[i][j][16])
                dict_loads['Tanggal Tolak'].append(data[i][j][24])
                dict_loads['Tujuan Kapal'].append(data[i][j][23])
                dict_loads['Jam Tolak'].append(data[i][j][25])
    
    loads= pd.DataFrame.from_dict(dict_loads)
    loads.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    loads['Nomor'] = range(1,len(loads)+1)

    return loads

def palm_oil(data=None):
    dict_palm = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Keagenan':[],
        'Bendera':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Tanggal Tolak':[],
        'Muatan Tiba':[],
        'Asal Kapal':[],
        'Muatan Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if '(SWT)' in data[i][j][19] or '(SWT)' in data[i][j][28]:
                if ';' in data[i][j][18]:
                    arr_load = data[i][j][18].split('; ')

                    for p in range(len(arr_load)):
                        if p == 0:
                            dict_palm['Nomor'].append(j)
                            dict_palm['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_palm['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_palm['Keagenan'].append(data[i][j][13])
                            dict_palm['Bendera'].append(data[i][j][12])
                            dict_palm['GT'].append(data[i][j][5])
                            dict_palm['Tanggal Tiba'].append(data[i][j][15])
                            dict_palm['Tanggal Tolak'].append(data[i][j][24])
                            dict_palm['Muatan Tiba'].append(arr_load[p])
                            dict_palm['Asal Kapal'].append(data[i][j][14])
                            dict_palm['Muatan Tolak'].append(data[i][j][27])
                            dict_palm['Tujuan Kapal'].append(data[i][j][23])
                            dict_palm['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_palm['Nomor'].append(None)
                            dict_palm['Kode Kapal'].append(None)
                            dict_palm['Nama Kapal'].append(None)
                            dict_palm['Keagenan'].append(None)
                            dict_palm['Bendera'].append(None)
                            dict_palm['GT'].append(None)
                            dict_palm['Tanggal Tiba'].append(None)
                            dict_palm['Tanggal Tolak'].append(data[i][j][24])
                            dict_palm['Muatan Tiba'].append(arr_load[p])
                            dict_palm['Asal Kapal'].append(None)
                            dict_palm['Muatan Tolak'].append(None)
                            dict_palm['Tujuan Kapal'].append(None)
                            dict_palm['Jam Tolak'].append(data[i][j][25])

                elif ';' in data[i][j][27]:
                    depar_load = data[i][j][27].split('; ')

                    for p in range(len(depar_load)):
                        if p == 0:
                            dict_palm['Nomor'].append(j)
                            dict_palm['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_palm['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_palm['Keagenan'].append(data[i][j][13])
                            dict_palm['Bendera'].append(data[i][j][12])
                            dict_palm['GT'].append(data[i][j][5])
                            dict_palm['Tanggal Tiba'].append(data[i][j][15])
                            dict_palm['Tanggal Tolak'].append(data[i][j][24])
                            dict_palm['Muatan Tiba'].append(data[i][j][18])
                            dict_palm['Asal Kapal'].append(data[i][j][14])
                            dict_palm['Muatan Tolak'].append(depar_load[p])
                            dict_palm['Tujuan Kapal'].append(data[i][j][23])
                            dict_palm['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_palm['Nomor'].append(None)
                            dict_palm['Kode Kapal'].append(None)
                            dict_palm['Nama Kapal'].append(None)
                            dict_palm['Keagenan'].append(None)
                            dict_palm['Bendera'].append(None)
                            dict_palm['GT'].append(None)
                            dict_palm['Tanggal Tiba'].append(None)
                            dict_palm['Tanggal Tolak'].append(data[i][j][24])
                            dict_palm['Muatan Tiba'].append(None)
                            dict_palm['Asal Kapal'].append(None)
                            dict_palm['Muatan Tolak'].append(depar_load[p])
                            dict_palm['Tujuan Kapal'].append(None)
                            dict_palm['Jam Tolak'].append(data[i][j][25])

                elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                    dict_palm['Nomor'].append(j)
                    dict_palm['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                    dict_palm['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                    dict_palm['Keagenan'].append(data[i][j][13])
                    dict_palm['Bendera'].append(data[i][j][12])
                    dict_palm['GT'].append(data[i][j][5])
                    dict_palm['Tanggal Tiba'].append(data[i][j][15])
                    dict_palm['Tanggal Tolak'].append(data[i][j][24])
                    dict_palm['Muatan Tiba'].append(data[i][j][18])
                    dict_palm['Asal Kapal'].append(data[i][j][14])
                    dict_palm['Muatan Tolak'].append(data[i][j][27])
                    dict_palm['Tujuan Kapal'].append(data[i][j][23])
                    dict_palm['Jam Tolak'].append(data[i][j][25])

    palm = pd.DataFrame.from_dict(dict_palm)
    palm.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(palm)):
        if np.isnan(palm.iat[i,0]) == False:
            palm.iat[i,0] = counter
            counter+=1
        elif np.isnan(palm.iat[i,0]) == True:
            pass

    return palm

def coals(data=None):
    dict_coals = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Keagenan':[],
        'Bendera':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Tanggal Tolak':[],
        'Muatan Tiba':[],
        'Asal Kapal':[],
        'Muatan Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if '(BABA)' in data[i][j][19] or '(BABA)' in data[i][j][28]:
                if ';' in data[i][j][18]:
                    arr_load = data[i][j][18].split('; ')

                    for p in range(len(arr_load)):
                        if p == 0:
                            dict_coals['Nomor'].append(j)
                            dict_coals['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_coals['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_coals['Keagenan'].append(data[i][j][13])
                            dict_coals['Bendera'].append(data[i][j][12])
                            dict_coals['GT'].append(data[i][j][5])
                            dict_coals['Tanggal Tiba'].append(data[i][j][15])
                            dict_coals['Tanggal Tolak'].append(data[i][j][24])
                            dict_coals['Muatan Tiba'].append(arr_load[p])
                            dict_coals['Asal Kapal'].append(data[i][j][14])
                            dict_coals['Muatan Tolak'].append(data[i][j][27])
                            dict_coals['Tujuan Kapal'].append(data[i][j][23])
                            dict_coals['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_coals['Nomor'].append(None)
                            dict_coals['Kode Kapal'].append(None)
                            dict_coals['Nama Kapal'].append(None)
                            dict_coals['Keagenan'].append(None)
                            dict_coals['Bendera'].append(None)
                            dict_coals['GT'].append(None)
                            dict_coals['Tanggal Tiba'].append(None)
                            dict_coals['Tanggal Tolak'].append(data[i][j][24])
                            dict_coals['Muatan Tiba'].append(arr_load[p])
                            dict_coals['Asal Kapal'].append(None)
                            dict_coals['Muatan Tolak'].append(None)
                            dict_coals['Tujuan Kapal'].append(None)
                            dict_coals['Jam Tolak'].append(data[i][j][25])

                elif ';' in data[i][j][27]:
                    depar_load = data[i][j][27].split('; ')

                    for p in range(len(depar_load)):
                        if p == 0:
                            dict_coals['Nomor'].append(j)
                            dict_coals['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_coals['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_coals['Keagenan'].append(data[i][j][13])
                            dict_coals['Bendera'].append(data[i][j][12])
                            dict_coals['GT'].append(data[i][j][5])
                            dict_coals['Tanggal Tiba'].append(data[i][j][15])
                            dict_coals['Tanggal Tolak'].append(data[i][j][24])
                            dict_coals['Muatan Tiba'].append(data[i][j][18])
                            dict_coals['Asal Kapal'].append(data[i][j][14])
                            dict_coals['Muatan Tolak'].append(depar_load[p])
                            dict_coals['Tujuan Kapal'].append(data[i][j][23])
                            dict_coals['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_coals['Nomor'].append(None)
                            dict_coals['Kode Kapal'].append(None)
                            dict_coals['Nama Kapal'].append(None)
                            dict_coals['Keagenan'].append(None)
                            dict_coals['Bendera'].append(None)
                            dict_coals['GT'].append(None)
                            dict_coals['Tanggal Tiba'].append(None)
                            dict_coals['Tanggal Tolak'].append(data[i][j][24])
                            dict_coals['Muatan Tiba'].append(None)
                            dict_coals['Asal Kapal'].append(None)
                            dict_coals['Muatan Tolak'].append(depar_load[p])
                            dict_coals['Tujuan Kapal'].append(None)
                            dict_coals['Jam Tolak'].append(data[i][j][25])

                elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                    dict_coals['Nomor'].append(j)
                    dict_coals['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                    dict_coals['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                    dict_coals['Keagenan'].append(data[i][j][13])
                    dict_coals['Bendera'].append(data[i][j][12])
                    dict_coals['GT'].append(data[i][j][5])
                    dict_coals['Tanggal Tiba'].append(data[i][j][15])
                    dict_coals['Tanggal Tolak'].append(data[i][j][24])
                    dict_coals['Muatan Tiba'].append(data[i][j][18])
                    dict_coals['Asal Kapal'].append(data[i][j][14])
                    dict_coals['Muatan Tolak'].append(data[i][j][27])
                    dict_coals['Tujuan Kapal'].append(data[i][j][23])
                    dict_coals['Jam Tolak'].append(data[i][j][25])

    coals = pd.DataFrame.from_dict(dict_coals)
    coals.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(coals)):
        if np.isnan(coals.iat[i,0]) == False:
            coals.iat[i,0] = counter
            counter+=1
        elif np.isnan(coals.iat[i,0]) == True:
            pass

    return coals

def gen_cargo(data=None):
    dict_cargo = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Keagenan':[],
        'Bendera':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Tanggal Tolak':[],
        'Muatan Tiba':[],
        'Asal Kapal':[],
        'Muatan Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if '(GEAR)' in data[i][j][19] or '(GEAR)' in data[i][j][28]:
                if ';' in data[i][j][18]:
                    arr_load = data[i][j][18].split('; ')

                    for p in range(len(arr_load)):
                        if p == 0:
                            dict_cargo['Nomor'].append(j)
                            dict_cargo['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_cargo['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_cargo['Keagenan'].append(data[i][j][13])
                            dict_cargo['Bendera'].append(data[i][j][12])
                            dict_cargo['GT'].append(data[i][j][5])
                            dict_cargo['Tanggal Tiba'].append(data[i][j][15])
                            dict_cargo['Tanggal Tolak'].append(data[i][j][24])
                            dict_cargo['Muatan Tiba'].append(arr_load[p])
                            dict_cargo['Asal Kapal'].append(data[i][j][14])
                            dict_cargo['Muatan Tolak'].append(data[i][j][27])
                            dict_cargo['Tujuan Kapal'].append(data[i][j][23])
                            dict_cargo['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_cargo['Nomor'].append(None)
                            dict_cargo['Kode Kapal'].append(None)
                            dict_cargo['Nama Kapal'].append(None)
                            dict_cargo['Keagenan'].append(None)
                            dict_cargo['Bendera'].append(None)
                            dict_cargo['GT'].append(None)
                            dict_cargo['Tanggal Tiba'].append(None)
                            dict_cargo['Tanggal Tolak'].append(data[i][j][24])
                            dict_cargo['Muatan Tiba'].append(arr_load[p])
                            dict_cargo['Asal Kapal'].append(None)
                            dict_cargo['Muatan Tolak'].append(None)
                            dict_cargo['Tujuan Kapal'].append(None)
                            dict_cargo['Jam Tolak'].append(data[i][j][25])

                elif ';' in data[i][j][27]:
                    depar_load = data[i][j][27].split('; ')

                    for p in range(len(depar_load)):
                        if p == 0:
                            dict_cargo['Nomor'].append(j)
                            dict_cargo['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_cargo['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_cargo['Keagenan'].append(data[i][j][13])
                            dict_cargo['Bendera'].append(data[i][j][12])
                            dict_cargo['GT'].append(data[i][j][5])
                            dict_cargo['Tanggal Tiba'].append(data[i][j][15])
                            dict_cargo['Tanggal Tolak'].append(data[i][j][24])
                            dict_cargo['Muatan Tiba'].append(data[i][j][18])
                            dict_cargo['Asal Kapal'].append(data[i][j][14])
                            dict_cargo['Muatan Tolak'].append(depar_load[p])
                            dict_cargo['Tujuan Kapal'].append(data[i][j][23])
                            dict_cargo['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_cargo['Nomor'].append(None)
                            dict_cargo['Kode Kapal'].append(None)
                            dict_cargo['Nama Kapal'].append(None)
                            dict_cargo['Keagenan'].append(None)
                            dict_cargo['Bendera'].append(None)
                            dict_cargo['GT'].append(None)
                            dict_cargo['Tanggal Tiba'].append(None)
                            dict_cargo['Tanggal Tolak'].append(data[i][j][24])
                            dict_cargo['Muatan Tiba'].append(None)
                            dict_cargo['Asal Kapal'].append(None)
                            dict_cargo['Muatan Tolak'].append(depar_load[p])
                            dict_cargo['Tujuan Kapal'].append(None)
                            dict_cargo['Jam Tolak'].append(data[i][j][25])

                elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                    dict_cargo['Nomor'].append(j)
                    dict_cargo['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                    dict_cargo['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                    dict_cargo['Keagenan'].append(data[i][j][13])
                    dict_cargo['Bendera'].append(data[i][j][12])
                    dict_cargo['GT'].append(data[i][j][5])
                    dict_cargo['Tanggal Tiba'].append(data[i][j][15])
                    dict_cargo['Tanggal Tolak'].append(data[i][j][24])
                    dict_cargo['Muatan Tiba'].append(data[i][j][18])
                    dict_cargo['Asal Kapal'].append(data[i][j][14])
                    dict_cargo['Muatan Tolak'].append(data[i][j][27])
                    dict_cargo['Tujuan Kapal'].append(data[i][j][23])
                    dict_cargo['Jam Tolak'].append(data[i][j][25])

    cargos = pd.DataFrame.from_dict(dict_cargo)
    cargos.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(cargos)):
        if np.isnan(cargos.iat[i,0]) == False:
            cargos.iat[i,0] = counter
            counter+=1
        elif np.isnan(cargos.iat[i,0]) == True:
            pass

    return cargos

def stones(data=None):
    dict_stone = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Keagenan':[],
        'Bendera':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Tanggal Tolak':[],
        'Muatan Tiba':[],
        'Asal Kapal':[],
        'Muatan Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if '(BAPE)' in data[i][j][19] or '(BAPE)' in data[i][j][28]:
                if ';' in data[i][j][18]:
                    arr_load = data[i][j][18].split('; ')

                    for p in range(len(arr_load)):
                        if p == 0:
                            dict_stone['Nomor'].append(j)
                            dict_stone['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_stone['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_stone['Keagenan'].append(data[i][j][13])
                            dict_stone['Bendera'].append(data[i][j][12])
                            dict_stone['GT'].append(data[i][j][5])
                            dict_stone['Tanggal Tiba'].append(data[i][j][15])
                            dict_stone['Tanggal Tolak'].append(data[i][j][24])
                            dict_stone['Muatan Tiba'].append(arr_load[p])
                            dict_stone['Asal Kapal'].append(data[i][j][14])
                            dict_stone['Muatan Tolak'].append(data[i][j][27])
                            dict_stone['Tujuan Kapal'].append(data[i][j][23])
                            dict_stone['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_stone['Nomor'].append(None)
                            dict_stone['Kode Kapal'].append(None)
                            dict_stone['Nama Kapal'].append(None)
                            dict_stone['Keagenan'].append(None)
                            dict_stone['Bendera'].append(None)
                            dict_stone['GT'].append(None)
                            dict_stone['Tanggal Tiba'].append(None)
                            dict_stone['Tanggal Tolak'].append(data[i][j][24])
                            dict_stone['Muatan Tiba'].append(arr_load[p])
                            dict_stone['Asal Kapal'].append(None)
                            dict_stone['Muatan Tolak'].append(None)
                            dict_stone['Tujuan Kapal'].append(None)
                            dict_stone['Jam Tolak'].append(data[i][j][25])

                elif ';' in data[i][j][27]:
                    depar_load = data[i][j][27].split('; ')

                    for p in range(len(depar_load)):
                        if p == 0:
                            dict_stone['Nomor'].append(j)
                            dict_stone['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_stone['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_stone['Keagenan'].append(data[i][j][13])
                            dict_stone['Bendera'].append(data[i][j][12])
                            dict_stone['GT'].append(data[i][j][5])
                            dict_stone['Tanggal Tiba'].append(data[i][j][15])
                            dict_stone['Tanggal Tolak'].append(data[i][j][24])
                            dict_stone['Muatan Tiba'].append(data[i][j][18])
                            dict_stone['Asal Kapal'].append(data[i][j][14])
                            dict_stone['Muatan Tolak'].append(depar_load[p])
                            dict_stone['Tujuan Kapal'].append(data[i][j][23])
                            dict_stone['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_stone['Nomor'].append(None)
                            dict_stone['Kode Kapal'].append(None)
                            dict_stone['Nama Kapal'].append(None)
                            dict_stone['Keagenan'].append(None)
                            dict_stone['Bendera'].append(None)
                            dict_stone['GT'].append(None)
                            dict_stone['Tanggal Tiba'].append(None)
                            dict_stone['Tanggal Tolak'].append(data[i][j][24])
                            dict_stone['Muatan Tiba'].append(None)
                            dict_stone['Asal Kapal'].append(None)
                            dict_stone['Muatan Tolak'].append(depar_load[p])
                            dict_stone['Tujuan Kapal'].append(None)
                            dict_stone['Jam Tolak'].append(data[i][j][25])

                elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                    dict_stone['Nomor'].append(j)
                    dict_stone['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                    dict_stone['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                    dict_stone['Keagenan'].append(data[i][j][13])
                    dict_stone['Bendera'].append(data[i][j][12])
                    dict_stone['GT'].append(data[i][j][5])
                    dict_stone['Tanggal Tiba'].append(data[i][j][15])
                    dict_stone['Tanggal Tolak'].append(data[i][j][24])
                    dict_stone['Muatan Tiba'].append(data[i][j][18])
                    dict_stone['Asal Kapal'].append(data[i][j][14])
                    dict_stone['Muatan Tolak'].append(data[i][j][27])
                    dict_stone['Tujuan Kapal'].append(data[i][j][23])
                    dict_stone['Jam Tolak'].append(data[i][j][25])
                
    stone = pd.DataFrame.from_dict(dict_stone)
    stone.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(stone)):
        if np.isnan(stone.iat[i,0]) == False:
            stone.iat[i,0] = counter
            counter+=1
        elif np.isnan(stone.iat[i,0]) == True:
            pass

    return stone

def crudes(data=None):
    dict_oils = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Keagenan':[],
        'Bendera':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Tanggal Tolak':[],
        'Muatan Tiba':[],
        'Asal Kapal':[],
        'Muatan Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if '(CRIL)' in data[i][j][19] or '(CRIL)' in data[i][j][28]:
                if ';' in data[i][j][18]:
                    arr_load = data[i][j][18].split('; ')

                    for p in range(len(arr_load)):
                        if p == 0:
                            dict_oils['Nomor'].append(j)
                            dict_oils['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_oils['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_oils['Keagenan'].append(data[i][j][13])
                            dict_oils['Bendera'].append(data[i][j][12])
                            dict_oils['GT'].append(data[i][j][5])
                            dict_oils['Tanggal Tiba'].append(data[i][j][15])
                            dict_oils['Tanggal Tolak'].append(data[i][j][24])
                            dict_oils['Muatan Tiba'].append(arr_load[p])
                            dict_oils['Asal Kapal'].append(data[i][j][14])
                            dict_oils['Muatan Tolak'].append(data[i][j][27])
                            dict_oils['Tujuan Kapal'].append(data[i][j][23])
                            dict_oils['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_oils['Nomor'].append(None)
                            dict_oils['Kode Kapal'].append(None)
                            dict_oils['Nama Kapal'].append(None)
                            dict_oils['Keagenan'].append(None)
                            dict_oils['Bendera'].append(None)
                            dict_oils['GT'].append(None)
                            dict_oils['Tanggal Tiba'].append(None)
                            dict_oils['Tanggal Tolak'].append(data[i][j][24])
                            dict_oils['Muatan Tiba'].append(arr_load[p])
                            dict_oils['Asal Kapal'].append(None)
                            dict_oils['Muatan Tolak'].append(None)
                            dict_oils['Tujuan Kapal'].append(None)
                            dict_oils['Jam Tolak'].append(data[i][j][25])

                elif ';' in data[i][j][27]:
                    depar_load = data[i][j][27].split('; ')

                    for p in range(len(depar_load)):
                        if p == 0:
                            dict_oils['Nomor'].append(j)
                            dict_oils['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_oils['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_oils['Keagenan'].append(data[i][j][13])
                            dict_oils['Bendera'].append(data[i][j][12])
                            dict_oils['GT'].append(data[i][j][5])
                            dict_oils['Tanggal Tiba'].append(data[i][j][15])
                            dict_oils['Tanggal Tolak'].append(data[i][j][24])
                            dict_oils['Muatan Tiba'].append(data[i][j][18])
                            dict_oils['Asal Kapal'].append(data[i][j][14])
                            dict_oils['Muatan Tolak'].append(depar_load[p])
                            dict_oils['Tujuan Kapal'].append(data[i][j][23])
                            dict_oils['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_oils['Nomor'].append(None)
                            dict_oils['Kode Kapal'].append(None)
                            dict_oils['Nama Kapal'].append(None)
                            dict_oils['Keagenan'].append(None)
                            dict_oils['Bendera'].append(None)
                            dict_oils['GT'].append(None)
                            dict_oils['Tanggal Tiba'].append(None)
                            dict_oils['Tanggal Tolak'].append(data[i][j][24])
                            dict_oils['Muatan Tiba'].append(None)
                            dict_oils['Asal Kapal'].append(None)
                            dict_oils['Muatan Tolak'].append(depar_load[p])
                            dict_oils['Tujuan Kapal'].append(None)
                            dict_oils['Jam Tolak'].append(data[i][j][25])

                elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                    dict_oils['Nomor'].append(j)
                    dict_oils['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                    dict_oils['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                    dict_oils['Keagenan'].append(data[i][j][13])
                    dict_oils['Bendera'].append(data[i][j][12])
                    dict_oils['GT'].append(data[i][j][5])
                    dict_oils['Tanggal Tiba'].append(data[i][j][15])
                    dict_oils['Tanggal Tolak'].append(data[i][j][24])
                    dict_oils['Muatan Tiba'].append(data[i][j][18])
                    dict_oils['Asal Kapal'].append(data[i][j][14])
                    dict_oils['Muatan Tolak'].append(data[i][j][27])
                    dict_oils['Tujuan Kapal'].append(data[i][j][23])
                    dict_oils['Jam Tolak'].append(data[i][j][25])

    crudeoil = pd.DataFrame.from_dict(dict_oils)
    crudeoil.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(crudeoil)):
        if np.isnan(crudeoil.iat[i,0]) == False:
            crudeoil.iat[i,0] = counter
            counter+=1
        elif np.isnan(crudeoil.iat[i,0]) == True:
            pass

    return crudeoil

def heavies(data=None):
    dict_heavies = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Keagenan':[],
        'Bendera':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Tanggal Tolak':[],
        'Muatan Tiba':[],
        'Asal Kapal':[],
        'Muatan Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if '(ALBE)' in data[i][j][19] or '(ALBE)' in data[i][j][28]:
                if ';' in data[i][j][18]:
                    arr_load = data[i][j][18].split('; ')

                    for p in range(len(arr_load)):
                        if p == 0:
                            dict_heavies['Nomor'].append(j)
                            dict_heavies['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_heavies['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_heavies['Keagenan'].append(data[i][j][13])
                            dict_heavies['Bendera'].append(data[i][j][12])
                            dict_heavies['GT'].append(data[i][j][5])
                            dict_heavies['Tanggal Tiba'].append(data[i][j][15])
                            dict_heavies['Tanggal Tolak'].append(data[i][j][24])
                            dict_heavies['Muatan Tiba'].append(arr_load[p])
                            dict_heavies['Asal Kapal'].append(data[i][j][14])
                            dict_heavies['Muatan Tolak'].append(data[i][j][27])
                            dict_heavies['Tujuan Kapal'].append(data[i][j][23])
                            dict_heavies['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_heavies['Nomor'].append(None)
                            dict_heavies['Kode Kapal'].append(None)
                            dict_heavies['Nama Kapal'].append(None)
                            dict_heavies['Keagenan'].append(None)
                            dict_heavies['Bendera'].append(None)
                            dict_heavies['GT'].append(None)
                            dict_heavies['Tanggal Tiba'].append(None)
                            dict_heavies['Tanggal Tolak'].append(data[i][j][24])
                            dict_heavies['Muatan Tiba'].append(arr_load[p])
                            dict_heavies['Asal Kapal'].append(None)
                            dict_heavies['Muatan Tolak'].append(None)
                            dict_heavies['Tujuan Kapal'].append(None)
                            dict_heavies['Jam Tolak'].append(data[i][j][25])

                elif ';' in data[i][j][27]:
                    depar_load = data[i][j][27].split('; ')

                    for p in range(len(depar_load)):
                        if p == 0:
                            dict_heavies['Nomor'].append(j)
                            dict_heavies['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_heavies['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_heavies['Keagenan'].append(data[i][j][13])
                            dict_heavies['Bendera'].append(data[i][j][12])
                            dict_heavies['GT'].append(data[i][j][5])
                            dict_heavies['Tanggal Tiba'].append(data[i][j][15])
                            dict_heavies['Tanggal Tolak'].append(data[i][j][24])
                            dict_heavies['Muatan Tiba'].append(data[i][j][18])
                            dict_heavies['Asal Kapal'].append(data[i][j][14])
                            dict_heavies['Muatan Tolak'].append(depar_load[p])
                            dict_heavies['Tujuan Kapal'].append(data[i][j][23])
                            dict_heavies['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_heavies['Nomor'].append(None)
                            dict_heavies['Kode Kapal'].append(None)
                            dict_heavies['Nama Kapal'].append(None)
                            dict_heavies['Keagenan'].append(None)
                            dict_heavies['Bendera'].append(None)
                            dict_heavies['GT'].append(None)
                            dict_heavies['Tanggal Tiba'].append(None)
                            dict_heavies['Tanggal Tolak'].append(data[i][j][24])
                            dict_heavies['Muatan Tiba'].append(None)
                            dict_heavies['Asal Kapal'].append(None)
                            dict_heavies['Muatan Tolak'].append(depar_load[p])
                            dict_heavies['Tujuan Kapal'].append(None)
                            dict_heavies['Jam Tolak'].append(data[i][j][25])

                elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                    dict_heavies['Nomor'].append(j)
                    dict_heavies['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                    dict_heavies['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                    dict_heavies['Keagenan'].append(data[i][j][13])
                    dict_heavies['Bendera'].append(data[i][j][12])
                    dict_heavies['GT'].append(data[i][j][5])
                    dict_heavies['Tanggal Tiba'].append(data[i][j][15])
                    dict_heavies['Tanggal Tolak'].append(data[i][j][24])
                    dict_heavies['Muatan Tiba'].append(data[i][j][18])
                    dict_heavies['Asal Kapal'].append(data[i][j][14])
                    dict_heavies['Muatan Tolak'].append(data[i][j][27])
                    dict_heavies['Tujuan Kapal'].append(data[i][j][23])
                    dict_heavies['Jam Tolak'].append(data[i][j][25])
                
    heavy = pd.DataFrame.from_dict(dict_heavies)
    heavy.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(heavy)):
        if np.isnan(heavy.iat[i,0]) == False:
            heavy.iat[i,0] = counter
            counter+=1
        elif np.isnan(heavy.iat[i,0]) == True:
            pass

    return heavy

def bbm(data=None):
    dict_fuels = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Keagenan':[],
        'Bendera':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Tanggal Tolak':[],
        'Muatan Tiba':[],
        'Asal Kapal':[],
        'Muatan Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if '(BBM)' in data[i][j][19] or '(BBM)' in data[i][j][28]:
                if ';' in data[i][j][18]:
                    arr_load = data[i][j][18].split('; ')

                    for p in range(len(arr_load)):
                        if p == 0:
                            dict_fuels['Nomor'].append(j)
                            dict_fuels['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_fuels['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_fuels['Keagenan'].append(data[i][j][13])
                            dict_fuels['Bendera'].append(data[i][j][12])
                            dict_fuels['GT'].append(data[i][j][5])
                            dict_fuels['Tanggal Tiba'].append(data[i][j][15])
                            dict_fuels['Tanggal Tolak'].append(data[i][j][24])
                            dict_fuels['Muatan Tiba'].append(arr_load[p])
                            dict_fuels['Asal Kapal'].append(data[i][j][14])
                            dict_fuels['Muatan Tolak'].append(data[i][j][27])
                            dict_fuels['Tujuan Kapal'].append(data[i][j][23])
                            dict_fuels['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_fuels['Nomor'].append(None)
                            dict_fuels['Kode Kapal'].append(None)
                            dict_fuels['Nama Kapal'].append(None)
                            dict_fuels['Keagenan'].append(None)
                            dict_fuels['Bendera'].append(None)
                            dict_fuels['GT'].append(None)
                            dict_fuels['Tanggal Tiba'].append(None)
                            dict_fuels['Tanggal Tolak'].append(data[i][j][24])
                            dict_fuels['Muatan Tiba'].append(arr_load[p])
                            dict_fuels['Asal Kapal'].append(None)
                            dict_fuels['Muatan Tolak'].append(None)
                            dict_fuels['Tujuan Kapal'].append(None)
                            dict_fuels['Jam Tolak'].append(data[i][j][25])

                elif ';' in data[i][j][27]:
                    depar_load = data[i][j][27].split('; ')

                    for p in range(len(depar_load)):
                        if p == 0:
                            dict_fuels['Nomor'].append(j)
                            dict_fuels['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_fuels['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_fuels['Keagenan'].append(data[i][j][13])
                            dict_fuels['Bendera'].append(data[i][j][12])
                            dict_fuels['GT'].append(data[i][j][5])
                            dict_fuels['Tanggal Tiba'].append(data[i][j][15])
                            dict_fuels['Tanggal Tolak'].append(data[i][j][24])
                            dict_fuels['Muatan Tiba'].append(data[i][j][18])
                            dict_fuels['Asal Kapal'].append(data[i][j][14])
                            dict_fuels['Muatan Tolak'].append(depar_load[p])
                            dict_fuels['Tujuan Kapal'].append(data[i][j][23])
                            dict_fuels['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_fuels['Nomor'].append(None)
                            dict_fuels['Kode Kapal'].append(None)
                            dict_fuels['Nama Kapal'].append(None)
                            dict_fuels['Keagenan'].append(None)
                            dict_fuels['Bendera'].append(None)
                            dict_fuels['GT'].append(None)
                            dict_fuels['Tanggal Tiba'].append(None)
                            dict_fuels['Tanggal Tolak'].append(data[i][j][24])
                            dict_fuels['Muatan Tiba'].append(None)
                            dict_fuels['Asal Kapal'].append(None)
                            dict_fuels['Muatan Tolak'].append(depar_load[p])
                            dict_fuels['Tujuan Kapal'].append(None)
                            dict_fuels['Jam Tolak'].append(data[i][j][25])

                elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                    dict_fuels['Nomor'].append(j)
                    dict_fuels['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                    dict_fuels['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                    dict_fuels['Keagenan'].append(data[i][j][13])
                    dict_fuels['Bendera'].append(data[i][j][12])
                    dict_fuels['GT'].append(data[i][j][5])
                    dict_fuels['Tanggal Tiba'].append(data[i][j][15])
                    dict_fuels['Tanggal Tolak'].append(data[i][j][24])
                    dict_fuels['Muatan Tiba'].append(data[i][j][18])
                    dict_fuels['Asal Kapal'].append(data[i][j][14])
                    dict_fuels['Muatan Tolak'].append(data[i][j][27])
                    dict_fuels['Tujuan Kapal'].append(data[i][j][23])
                    dict_fuels['Jam Tolak'].append(data[i][j][25])

    fuels = pd.DataFrame.from_dict(dict_fuels)
    fuels.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(fuels)):
        if np.isnan(fuels.iat[i,0]) == False:
            fuels.iat[i,0] = counter
            counter+=1
        elif np.isnan(fuels.iat[i,0]) == True:
            pass

    return fuels

def kndr(data=None):
    dict_carbike = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Keagenan':[],
        'Bendera':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Tanggal Tolak':[],
        'Muatan Tiba':[],
        'Asal Kapal':[],
        'Muatan Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if '(KNDR)' in data[i][j][19] or '(KNDR)' in data[i][j][28]:
                if ';' in data[i][j][18]:
                    arr_load = data[i][j][18].split('; ')
                    for p in range(len(arr_load)):
                        if p == 0:
                            dict_carbike['Nomor'].append(j)
                            dict_carbike['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_carbike['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_carbike['Keagenan'].append(data[i][j][13])
                            dict_carbike['Bendera'].append(data[i][j][12])
                            dict_carbike['GT'].append(data[i][j][5])
                            dict_carbike['Tanggal Tiba'].append(data[i][j][15])
                            dict_carbike['Tanggal Tolak'].append(data[i][j][24])
                            dict_carbike['Muatan Tiba'].append(arr_load[p])
                            dict_carbike['Asal Kapal'].append(data[i][j][14])
                            dict_carbike['Muatan Tolak'].append(data[i][j][27])
                            dict_carbike['Tujuan Kapal'].append(data[i][j][23])
                            dict_carbike['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_carbike['Nomor'].append(None)
                            dict_carbike['Kode Kapal'].append(None)
                            dict_carbike['Nama Kapal'].append(None)
                            dict_carbike['Keagenan'].append(None)
                            dict_carbike['Bendera'].append(None)
                            dict_carbike['GT'].append(None)
                            dict_carbike['Tanggal Tiba'].append(None)
                            dict_carbike['Tanggal Tolak'].append(data[i][j][24])
                            dict_carbike['Muatan Tiba'].append(arr_load[p])
                            dict_carbike['Asal Kapal'].append(None)
                            dict_carbike['Muatan Tolak'].append(None)
                            dict_carbike['Tujuan Kapal'].append(None)
                            dict_carbike['Jam Tolak'].append(data[i][j][25])

                elif ';' in data[i][j][27]:
                    depar_load = data[i][j][27].split('; ')
                    for p in range(len(depar_load)):
                        if p == 0:
                            dict_carbike['Nomor'].append(j)
                            dict_carbike['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_carbike['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_carbike['Keagenan'].append(data[i][j][13])
                            dict_carbike['Bendera'].append(data[i][j][12])
                            dict_carbike['GT'].append(data[i][j][5])
                            dict_carbike['Tanggal Tiba'].append(data[i][j][15])
                            dict_carbike['Tanggal Tolak'].append(data[i][j][24])
                            dict_carbike['Muatan Tiba'].append(data[i][j][18])
                            dict_carbike['Asal Kapal'].append(data[i][j][14])
                            dict_carbike['Muatan Tolak'].append(depar_load[p])
                            dict_carbike['Tujuan Kapal'].append(data[i][j][23])
                            dict_carbike['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_carbike['Nomor'].append(None)
                            dict_carbike['Kode Kapal'].append(None)
                            dict_carbike['Nama Kapal'].append(None)
                            dict_carbike['Keagenan'].append(None)
                            dict_carbike['Bendera'].append(None)
                            dict_carbike['GT'].append(None)
                            dict_carbike['Tanggal Tiba'].append(None)
                            dict_carbike['Tanggal Tolak'].append(data[i][j][24])
                            dict_carbike['Muatan Tiba'].append(None)
                            dict_carbike['Asal Kapal'].append(None)
                            dict_carbike['Muatan Tolak'].append(depar_load[p])
                            dict_carbike['Tujuan Kapal'].append(None)
                            dict_carbike['Jam Tolak'].append(data[i][j][25])

                elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                    dict_carbike['Nomor'].append(j)
                    dict_carbike['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                    dict_carbike['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                    dict_carbike['Keagenan'].append(data[i][j][13])
                    dict_carbike['Bendera'].append(data[i][j][12])
                    dict_carbike['GT'].append(data[i][j][5])
                    dict_carbike['Tanggal Tiba'].append(data[i][j][15])
                    dict_carbike['Tanggal Tolak'].append(data[i][j][24])
                    dict_carbike['Muatan Tiba'].append(data[i][j][18])
                    dict_carbike['Asal Kapal'].append(data[i][j][14])
                    dict_carbike['Muatan Tolak'].append(data[i][j][27])
                    dict_carbike['Tujuan Kapal'].append(data[i][j][23])
                    dict_carbike['Jam Tolak'].append(data[i][j][25])

    carbike = pd.DataFrame.from_dict(dict_carbike)
    carbike.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(carbike)):
        if np.isnan(carbike.iat[i,0]) == False:
            carbike.iat[i,0] = counter
            counter+=1
        elif np.isnan(carbike.iat[i,0]) == True:
            pass

    return carbike

def kayu(data=None):
    dict_wood = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Keagenan':[],
        'Bendera':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Tanggal Tolak':[],
        'Muatan Tiba':[],
        'Asal Kapal':[],
        'Muatan Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if '(KAYU)' in data[i][j][19] or '(KAYU)' in data[i][j][28]:
                if ';' in data[i][j][18]:
                    arr_load = data[i][j][18].split('; ')
                    for p in range(len(arr_load)):
                        if p == 0:
                            dict_wood['Nomor'].append(j)
                            dict_wood['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_wood['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_wood['Keagenan'].append(data[i][j][13])
                            dict_wood['Bendera'].append(data[i][j][12])
                            dict_wood['GT'].append(data[i][j][5])
                            dict_wood['Tanggal Tiba'].append(data[i][j][15])
                            dict_wood['Tanggal Tolak'].append(data[i][j][24])
                            dict_wood['Muatan Tiba'].append(arr_load[p])
                            dict_wood['Asal Kapal'].append(data[i][j][14])
                            dict_wood['Muatan Tolak'].append(data[i][j][27])
                            dict_wood['Tujuan Kapal'].append(data[i][j][23])
                            dict_wood['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_wood['Nomor'].append(None)
                            dict_wood['Kode Kapal'].append(None)
                            dict_wood['Nama Kapal'].append(None)
                            dict_wood['Keagenan'].append(None)
                            dict_wood['Bendera'].append(None)
                            dict_wood['GT'].append(None)
                            dict_wood['Tanggal Tiba'].append(None)
                            dict_wood['Tanggal Tolak'].append(data[i][j][24])
                            dict_wood['Muatan Tiba'].append(arr_load[p])
                            dict_wood['Asal Kapal'].append(None)
                            dict_wood['Muatan Tolak'].append(None)
                            dict_wood['Tujuan Kapal'].append(None)
                            dict_wood['Jam Tolak'].append(data[i][j][25])

                elif ';' in data[i][j][27]:
                    depar_load = data[i][j][27].split('; ')
                    for p in range(len(depar_load)):
                        if p == 0:
                            dict_wood['Nomor'].append(j)
                            dict_wood['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_wood['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_wood['Keagenan'].append(data[i][j][13])
                            dict_wood['Bendera'].append(data[i][j][12])
                            dict_wood['GT'].append(data[i][j][5])
                            dict_wood['Tanggal Tiba'].append(data[i][j][15])
                            dict_wood['Tanggal Tolak'].append(data[i][j][24])
                            dict_wood['Muatan Tiba'].append(data[i][j][18])
                            dict_wood['Asal Kapal'].append(data[i][j][14])
                            dict_wood['Muatan Tolak'].append(depar_load[p])
                            dict_wood['Tujuan Kapal'].append(data[i][j][23])
                            dict_wood['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_wood['Nomor'].append(None)
                            dict_wood['Kode Kapal'].append(None)
                            dict_wood['Nama Kapal'].append(None)
                            dict_wood['Keagenan'].append(None)
                            dict_wood['Bendera'].append(None)
                            dict_wood['GT'].append(None)
                            dict_wood['Tanggal Tiba'].append(None)
                            dict_wood['Tanggal Tolak'].append(data[i][j][24])
                            dict_wood['Muatan Tiba'].append(None)
                            dict_wood['Asal Kapal'].append(None)
                            dict_wood['Muatan Tolak'].append(depar_load[p])
                            dict_wood['Tujuan Kapal'].append(None)
                            dict_wood['Jam Tolak'].append(data[i][j][25])

                elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                    dict_wood['Nomor'].append(j)
                    dict_wood['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                    dict_wood['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                    dict_wood['Keagenan'].append(data[i][j][13])
                    dict_wood['Bendera'].append(data[i][j][12])
                    dict_wood['GT'].append(data[i][j][5])
                    dict_wood['Tanggal Tiba'].append(data[i][j][15])
                    dict_wood['Tanggal Tolak'].append(data[i][j][24])
                    dict_wood['Muatan Tiba'].append(data[i][j][18])
                    dict_wood['Asal Kapal'].append(data[i][j][14])
                    dict_wood['Muatan Tolak'].append(data[i][j][27])
                    dict_wood['Tujuan Kapal'].append(data[i][j][23])
                    dict_wood['Jam Tolak'].append(data[i][j][25])

    woody = pd.DataFrame.from_dict(dict_wood)
    woody.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(woody)):
        if np.isnan(woody.iat[i,0]) == False:
            woody.iat[i,0] = counter
            counter+=1
        elif np.isnan(woody.iat[i,0]) == True:
            pass

    return woody

def tanah(data=None):
    dict_sand = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Keagenan':[],
        'Bendera':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Tanggal Tolak':[],
        'Muatan Tiba':[],
        'Asal Kapal':[],
        'Muatan Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if '(TNH)' in data[i][j][19] or '(TNH)' in data[i][j][28]:
                if ';' in data[i][j][18]:
                    arr_load = data[i][j][18].split('; ')
                    for p in range(len(arr_load)):
                        if p == 0:
                            dict_sand['Nomor'].append(j)
                            dict_sand['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_sand['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_sand['Keagenan'].append(data[i][j][13])
                            dict_sand['Bendera'].append(data[i][j][12])
                            dict_sand['GT'].append(data[i][j][5])
                            dict_sand['Tanggal Tiba'].append(data[i][j][15])
                            dict_sand['Tanggal Tolak'].append(data[i][j][24])
                            dict_sand['Muatan Tiba'].append(arr_load[p])
                            dict_sand['Asal Kapal'].append(data[i][j][14])
                            dict_sand['Muatan Tolak'].append(data[i][j][27])
                            dict_sand['Tujuan Kapal'].append(data[i][j][23])
                            dict_sand['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_sand['Nomor'].append(None)
                            dict_sand['Kode Kapal'].append(None)
                            dict_sand['Nama Kapal'].append(None)
                            dict_sand['Keagenan'].append(None)
                            dict_sand['Bendera'].append(None)
                            dict_sand['GT'].append(None)
                            dict_sand['Tanggal Tiba'].append(None)
                            dict_sand['Tanggal Tolak'].append(data[i][j][24])
                            dict_sand['Muatan Tiba'].append(arr_load[p])
                            dict_sand['Asal Kapal'].append(None)
                            dict_sand['Muatan Tolak'].append(None)
                            dict_sand['Tujuan Kapal'].append(None)
                            dict_sand['Jam Tolak'].append(data[i][j][25])

                elif ';' in data[i][j][27]:
                    depar_load = data[i][j][27].split('; ')
                    for p in range(len(depar_load)):
                        if p == 0:
                            dict_sand['Nomor'].append(j)
                            dict_sand['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_sand['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_sand['Keagenan'].append(data[i][j][13])
                            dict_sand['Bendera'].append(data[i][j][12])
                            dict_sand['GT'].append(data[i][j][5])
                            dict_sand['Tanggal Tiba'].append(data[i][j][15])
                            dict_sand['Tanggal Tolak'].append(data[i][j][24])
                            dict_sand['Muatan Tiba'].append(data[i][j][18])
                            dict_sand['Asal Kapal'].append(data[i][j][14])
                            dict_sand['Muatan Tolak'].append(depar_load[p])
                            dict_sand['Tujuan Kapal'].append(data[i][j][23])
                            dict_sand['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_sand['Nomor'].append(None)
                            dict_sand['Kode Kapal'].append(None)
                            dict_sand['Nama Kapal'].append(None)
                            dict_sand['Keagenan'].append(None)
                            dict_sand['Bendera'].append(None)
                            dict_sand['GT'].append(None)
                            dict_sand['Tanggal Tiba'].append(None)
                            dict_sand['Tanggal Tolak'].append(data[i][j][24])
                            dict_sand['Muatan Tiba'].append(None)
                            dict_sand['Asal Kapal'].append(None)
                            dict_sand['Muatan Tolak'].append(depar_load[p])
                            dict_sand['Tujuan Kapal'].append(None)
                            dict_sand['Jam Tolak'].append(data[i][j][25])

                elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                    dict_sand['Nomor'].append(j)
                    dict_sand['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                    dict_sand['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                    dict_sand['Keagenan'].append(data[i][j][13])
                    dict_sand['Bendera'].append(data[i][j][12])
                    dict_sand['GT'].append(data[i][j][5])
                    dict_sand['Tanggal Tiba'].append(data[i][j][15])
                    dict_sand['Tanggal Tolak'].append(data[i][j][24])
                    dict_sand['Muatan Tiba'].append(data[i][j][18])
                    dict_sand['Asal Kapal'].append(data[i][j][14])
                    dict_sand['Muatan Tolak'].append(data[i][j][27])
                    dict_sand['Tujuan Kapal'].append(data[i][j][23])
                    dict_sand['Jam Tolak'].append(data[i][j][25])
                
    sandy = pd.DataFrame.from_dict(dict_sand)
    sandy.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(sandy)):
        if np.isnan(sandy.iat[i,0]) == False:
            sandy.iat[i,0] = counter
            counter+=1
        elif np.isnan(sandy.iat[i,0]) == True:
            pass

    return sandy

def campursari(data=None):
    dict_mix = {
        'Nomor':[],
        'Kode Kapal':[],
        'Nama Kapal':[],
        'Keagenan':[],
        'Bendera':[],
        'GT':[],
        'Tanggal Tiba':[],
        'Tanggal Tolak':[],
        'Muatan Tiba':[],
        'Asal Kapal':[],
        'Muatan Tolak':[],
        'Tujuan Kapal':[],
        'Jam Tolak':[],
    }

    for i in range(len(data)):
        for j in range(1,len(data[i])):
            if '(CMPR)' in data[i][j][19] or '(CMPR)' in data[i][j][28]:
                if ';' in data[i][j][18]:
                    arr_load = data[i][j][18].split('; ')
                    for p in range(len(arr_load)):
                        if p == 0:
                            dict_mix['Nomor'].append(j)
                            dict_mix['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_mix['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_mix['Keagenan'].append(data[i][j][13])
                            dict_mix['Bendera'].append(data[i][j][12])
                            dict_mix['GT'].append(data[i][j][5])
                            dict_mix['Tanggal Tiba'].append(data[i][j][15])
                            dict_mix['Tanggal Tolak'].append(data[i][j][24])
                            dict_mix['Muatan Tiba'].append(arr_load[p])
                            dict_mix['Asal Kapal'].append(data[i][j][14])
                            dict_mix['Muatan Tolak'].append(data[i][j][27])
                            dict_mix['Tujuan Kapal'].append(data[i][j][23])
                            dict_mix['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_mix['Nomor'].append(None)
                            dict_mix['Kode Kapal'].append(None)
                            dict_mix['Nama Kapal'].append(None)
                            dict_mix['Keagenan'].append(None)
                            dict_mix['Bendera'].append(None)
                            dict_mix['GT'].append(None)
                            dict_mix['Tanggal Tiba'].append(None)
                            dict_mix['Tanggal Tolak'].append(data[i][j][24])
                            dict_mix['Muatan Tiba'].append(arr_load[p])
                            dict_mix['Asal Kapal'].append(None)
                            dict_mix['Muatan Tolak'].append(None)
                            dict_mix['Tujuan Kapal'].append(None)
                            dict_mix['Jam Tolak'].append(data[i][j][25])

                elif ';' in data[i][j][27]:
                    depar_load = data[i][j][27].split('; ')
                    for p in range(len(depar_load)):
                        if p == 0:
                            dict_mix['Nomor'].append(j)
                            dict_mix['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                            dict_mix['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                            dict_mix['Keagenan'].append(data[i][j][13])
                            dict_mix['Bendera'].append(data[i][j][12])
                            dict_mix['GT'].append(data[i][j][5])
                            dict_mix['Tanggal Tiba'].append(data[i][j][15])
                            dict_mix['Tanggal Tolak'].append(data[i][j][24])
                            dict_mix['Muatan Tiba'].append(data[i][j][18])
                            dict_mix['Asal Kapal'].append(data[i][j][14])
                            dict_mix['Muatan Tolak'].append(depar_load[p])
                            dict_mix['Tujuan Kapal'].append(data[i][j][23])
                            dict_mix['Jam Tolak'].append(data[i][j][25])
                        elif p != 0:
                            dict_mix['Nomor'].append(None)
                            dict_mix['Kode Kapal'].append(None)
                            dict_mix['Nama Kapal'].append(None)
                            dict_mix['Keagenan'].append(None)
                            dict_mix['Bendera'].append(None)
                            dict_mix['GT'].append(None)
                            dict_mix['Tanggal Tiba'].append(None)
                            dict_mix['Tanggal Tolak'].append(data[i][j][24])
                            dict_mix['Muatan Tiba'].append(None)
                            dict_mix['Asal Kapal'].append(None)
                            dict_mix['Muatan Tolak'].append(depar_load[p])
                            dict_mix['Tujuan Kapal'].append(None)
                            dict_mix['Jam Tolak'].append(data[i][j][25])

                elif ';' not in data[i][j][18] and ';' not in data[i][j][27]:
                    dict_mix['Nomor'].append(j)
                    dict_mix['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                    dict_mix['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                    dict_mix['Keagenan'].append(data[i][j][13])
                    dict_mix['Bendera'].append(data[i][j][12])
                    dict_mix['GT'].append(data[i][j][5])
                    dict_mix['Tanggal Tiba'].append(data[i][j][15])
                    dict_mix['Tanggal Tolak'].append(data[i][j][24])
                    dict_mix['Muatan Tiba'].append(data[i][j][18])
                    dict_mix['Asal Kapal'].append(data[i][j][14])
                    dict_mix['Muatan Tolak'].append(data[i][j][27])
                    dict_mix['Tujuan Kapal'].append(data[i][j][23])
                    dict_mix['Jam Tolak'].append(data[i][j][25])

    mixy = pd.DataFrame.from_dict(dict_mix)
    mixy.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)
    counter = 1
    for i in range(len(mixy)):
        if np.isnan(mixy.iat[i,0]) == False:
            mixy.iat[i,0] = counter
            counter+=1
        elif np.isnan(mixy.iat[i,0]) == True:
            pass

    return mixy

# Fungsi utama untuk menjalankan program
def main():
    datadf = read_excel_file(filename='Tempat Input Kunjungan Kapal.xlsx')
    data_sibk,data_sibg = sib_based(data=datadf)
    data_gabungan = combine_data(data=datadf)
    data_domcateg = dom_categ(data=datadf)
    data_bunyu = bunyu(data=datadf)
    data_albunyu = bunyu_al(data=datadf)
    data_spb = port_clr(data=datadf)
    data_nihil = wo_loads(data=datadf)
    data_sawit = palm_oil(data=datadf)
    data_baba = coals(data=datadf)
    data_gencar = gen_cargo(data=datadf)
    data_batu = stones(data=datadf)
    data_crudeoil = crudes(data=datadf)
    data_alatberat = heavies(data=datadf)
    data_bbm = bbm(data=datadf)
    data_mobil = kndr(data=datadf)
    data_kayu = kayu(data=datadf)
    data_tanah = tanah(data=datadf)
    data_campuran = campursari(data=datadf)

    filewriter = pd.ExcelWriter('Untuk Hasil Olah Data.xlsx')

    data_sibk.to_excel(filewriter,'SIB Kecil',index=False) 
    data_sibg.to_excel(filewriter,'SIB Besar',index=False)
    data_gabungan.to_excel(filewriter,'TK.II UPT',index=False)
    data_domcateg.to_excel(filewriter,'Domestik',index=False)
    data_bunyu.to_excel(filewriter,'Bunyu',index=False)
    data_albunyu.to_excel(filewriter,'Bunyu AL',index=False)
    data_spb.to_excel(filewriter,'SPB',index=False)
    data_nihil.to_excel(filewriter,'Tanpa Muatan',index=False)
    data_sawit.to_excel(filewriter,'Sawit',index=False)
    data_baba.to_excel(filewriter,'Batubara',index=False)
    data_gencar.to_excel(filewriter,'General Cargo',index=False)
    data_batu.to_excel(filewriter,'Batu Pecah',index=False)
    data_crudeoil.to_excel(filewriter,'Crude Oil',index=False)
    data_alatberat.to_excel(filewriter,'Alat Berat',index=False)
    data_bbm.to_excel(filewriter,'BBM',index=False)
    data_mobil.to_excel(filewriter,'Mobil',index=False)
    data_kayu.to_excel(filewriter,'Kayu',index=False)
    data_tanah.to_excel(filewriter,'Tanah',index=False)
    data_campuran.to_excel(filewriter,'Campuran',index=False)

    filewriter.save()
    
if __name__ == '__main__':
    main()

''' ------------------------------End of The Code Writing---------------------------'''
