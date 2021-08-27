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

''' ----------------------Code Writing at 27th August 2021----------------------'''

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
            dict_combined['Tanggal Tolak'].append(data[i][j][22])
            dict_combined['Tujuan Kapal'].append(data[i][j][21])
            dict_combined['Jam Tolak'].append(data[i][j][23])
            dict_combined['Muatan Tiba'].append(data[i][j][18])
            dict_combined['Muatan Tolak'].append(data[i][j][25])
    
    combined = pd.DataFrame.from_dict(dict_combined)
    combined.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return combined

def dom_categ(data=None):
    dict_domcateg = {
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
            dict_domcateg['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
            dict_domcateg['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
            dict_domcateg['Keagenan'].append(data[i][j][13])
            dict_domcateg['Bendera'].append(data[i][j][12])
            dict_domcateg['GT'].append(data[i][j][5])
            dict_domcateg['Tanggal Tiba'].append(data[i][j][15])
            dict_domcateg['Tanggal Tolak'].append(data[i][j][22])
            dict_domcateg['Muatan Tiba'].append(data[i][j][18])
            dict_domcateg['Asal Kapal'].append(data[i][j][14])
            dict_domcateg['Muatan Tolak'].append(data[i][j][25])
            dict_domcateg['Tujuan Kapal'].append(data[i][j][21])
            dict_domcateg['Jam Tolak'].append(data[i][j][23])

    domcateg = pd.DataFrame.from_dict(dict_domcateg)
    domcateg.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return domcateg

def bunyu(data=None):
    dict_bunyu = {
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
            dict_bunyu['Tanggal Tolak'].append(data[i][j][22])
            dict_bunyu['Tujuan Kapal'].append(data[i][j][21])
            dict_bunyu['Jam Tolak'].append(data[i][j][23])
            dict_bunyu['Muatan Tiba'].append(data[i][j][18])
            dict_bunyu['Muatan Tolak'].append(data[i][j][25])
    
    bunyuisland = pd.DataFrame.from_dict(dict_bunyu)
    bunyuisland.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return bunyuisland

def bunyu_al(data=None):
    dict_albunyu = {
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
            if data[i][j][20] == 'BUNYU':
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
                dict_albunyu['Tanggal Tolak'].append(data[i][j][22])
                dict_albunyu['Tujuan Kapal'].append(data[i][j][21])
                dict_albunyu['Jam Tolak'].append(data[i][j][23])
                dict_albunyu['Muatan Tiba'].append(data[i][j][18])
                dict_albunyu['Muatan Tolak'].append(data[i][j][25])
    
    albunyu = pd.DataFrame.from_dict(dict_albunyu)
    albunyu.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return albunyu

def port_clr(data=None):
    dict_port = {
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
            dict_port['Tujuan Kapal'].append(data[i][j][21])
            dict_port['Tanggal Tolak'].append(data[i][j][22])
            dict_port['Muatan Tolak'].append(data[i][j][25])
            dict_port['Keagenan'].append(data[i][j][13])
            dict_port['Jam Tolak'].append(data[i][j][23])
    
    portclr = pd.DataFrame.from_dict(dict_port)
    portclr.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return portclr

def wo_loads(data=None):
    dict_loads = {
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
                dict_loads['Tanggal Tolak'].append(data[i][j][22])
                dict_loads['Tujuan Kapal'].append(data[i][j][21])
                dict_loads['Jam Tolak'].append(data[i][j][23])
            elif data[i][j][18] == 'NIHIL' and data[i][j][25] == 'NIHIL':
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
                dict_loads['Tanggal Tolak'].append(data[i][j][22])
                dict_loads['Tujuan Kapal'].append(data[i][j][21])
                dict_loads['Jam Tolak'].append(data[i][j][23])
    
    loads= pd.DataFrame.from_dict(dict_loads)
    loads.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return loads

def palm_oil(data=None):
    dict_palm = {
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
            if '(SAWIT)' in data[i][j][18] or '(SAWIT)' in data[i][j][25]:
                dict_palm['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_palm['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_palm['Keagenan'].append(data[i][j][13])
                dict_palm['Bendera'].append(data[i][j][12])
                dict_palm['GT'].append(data[i][j][5])
                dict_palm['Tanggal Tiba'].append(data[i][j][15])
                dict_palm['Tanggal Tolak'].append(data[i][j][22])
                dict_palm['Muatan Tiba'].append(data[i][j][18])
                dict_palm['Asal Kapal'].append(data[i][j][14])
                dict_palm['Muatan Tolak'].append(data[i][j][25])
                dict_palm['Tujuan Kapal'].append(data[i][j][21])
                dict_palm['Jam Tolak'].append(data[i][j][23])

    palm = pd.DataFrame.from_dict(dict_palm)
    palm.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return palm

def coals(data=None):
    dict_coals = {
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
            if '(BABA)' in data[i][j][18] or '(BABA)' in data[i][j][25]:
                dict_coals['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_coals['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_coals['Keagenan'].append(data[i][j][13])
                dict_coals['Bendera'].append(data[i][j][12])
                dict_coals['GT'].append(data[i][j][5])
                dict_coals['Tanggal Tiba'].append(data[i][j][15])
                dict_coals['Tanggal Tolak'].append(data[i][j][22])
                dict_coals['Muatan Tiba'].append(data[i][j][18])
                dict_coals['Asal Kapal'].append(data[i][j][14])
                dict_coals['Muatan Tolak'].append(data[i][j][25])
                dict_coals['Tujuan Kapal'].append(data[i][j][21])
                dict_coals['Jam Tolak'].append(data[i][j][23])
            elif data[i][j][4][:data[i][j][4].find('. ')] == 'TB':
                dict_coals['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_coals['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_coals['Keagenan'].append(data[i][j][13])
                dict_coals['Bendera'].append(data[i][j][12])
                dict_coals['GT'].append(data[i][j][5])
                dict_coals['Tanggal Tiba'].append(data[i][j][15])
                dict_coals['Tanggal Tolak'].append(data[i][j][22])
                dict_coals['Muatan Tiba'].append(data[i][j][18])
                dict_coals['Asal Kapal'].append(data[i][j][14])
                dict_coals['Muatan Tolak'].append(data[i][j][25])
                dict_coals['Tujuan Kapal'].append(data[i][j][21])
                dict_coals['Jam Tolak'].append(data[i][j][23])

    coals = pd.DataFrame.from_dict(dict_coals)
    coals.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return coals

def gen_cargo(data=None):
    dict_cargo = {
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
            if '(GC)' in data[i][j][18] or '(GC)' in data[i][j][25]:
                dict_cargo['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_cargo['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_cargo['Keagenan'].append(data[i][j][13])
                dict_cargo['Bendera'].append(data[i][j][12])
                dict_cargo['GT'].append(data[i][j][5])
                dict_cargo['Tanggal Tiba'].append(data[i][j][15])
                dict_cargo['Tanggal Tolak'].append(data[i][j][22])
                dict_cargo['Muatan Tiba'].append(data[i][j][18])
                dict_cargo['Asal Kapal'].append(data[i][j][14])
                dict_cargo['Muatan Tolak'].append(data[i][j][25])
                dict_cargo['Tujuan Kapal'].append(data[i][j][21])
                dict_cargo['Jam Tolak'].append(data[i][j][23])
            elif data[i][j][4][:data[i][j][4].find('. ')] == 'TB':
                dict_cargo['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_cargo['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_cargo['Keagenan'].append(data[i][j][13])
                dict_cargo['Bendera'].append(data[i][j][12])
                dict_cargo['GT'].append(data[i][j][5])
                dict_cargo['Tanggal Tiba'].append(data[i][j][15])
                dict_cargo['Tanggal Tolak'].append(data[i][j][22])
                dict_cargo['Muatan Tiba'].append(data[i][j][18])
                dict_cargo['Asal Kapal'].append(data[i][j][14])
                dict_cargo['Muatan Tolak'].append(data[i][j][25])
                dict_cargo['Tujuan Kapal'].append(data[i][j][21])
                dict_cargo['Jam Tolak'].append(data[i][j][23])

    cargos = pd.DataFrame.from_dict(dict_cargo)
    cargos.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return cargos

def stones(data=None):
    dict_stone = {
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
            if '(BAPE)' in data[i][j][18] or '(BAPE)' in data[i][j][25]:
                dict_stone['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_stone['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_stone['Keagenan'].append(data[i][j][13])
                dict_stone['Bendera'].append(data[i][j][12])
                dict_stone['GT'].append(data[i][j][5])
                dict_stone['Tanggal Tiba'].append(data[i][j][15])
                dict_stone['Tanggal Tolak'].append(data[i][j][22])
                dict_stone['Muatan Tiba'].append(data[i][j][18])
                dict_stone['Asal Kapal'].append(data[i][j][14])
                dict_stone['Muatan Tolak'].append(data[i][j][25])
                dict_stone['Tujuan Kapal'].append(data[i][j][21])
                dict_stone['Jam Tolak'].append(data[i][j][23])
            elif data[i][j][4][:data[i][j][4].find('. ')] == 'TB':
                dict_stone['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_stone['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_stone['Keagenan'].append(data[i][j][13])
                dict_stone['Bendera'].append(data[i][j][12])
                dict_stone['GT'].append(data[i][j][5])
                dict_stone['Tanggal Tiba'].append(data[i][j][15])
                dict_stone['Tanggal Tolak'].append(data[i][j][22])
                dict_stone['Muatan Tiba'].append(data[i][j][18])
                dict_stone['Asal Kapal'].append(data[i][j][14])
                dict_stone['Muatan Tolak'].append(data[i][j][25])
                dict_stone['Tujuan Kapal'].append(data[i][j][21])
                dict_stone['Jam Tolak'].append(data[i][j][23])
                
    stone = pd.DataFrame.from_dict(dict_stone)
    stone.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return stone

def crudes(data=None):
    dict_oils = {
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
            if '(CO)' in data[i][j][18] or '(CO)' in data[i][j][25]:
                dict_oils['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_oils['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_oils['Keagenan'].append(data[i][j][13])
                dict_oils['Bendera'].append(data[i][j][12])
                dict_oils['GT'].append(data[i][j][5])
                dict_oils['Tanggal Tiba'].append(data[i][j][15])
                dict_oils['Tanggal Tolak'].append(data[i][j][22])
                dict_oils['Muatan Tiba'].append(data[i][j][18])
                dict_oils['Asal Kapal'].append(data[i][j][14])
                dict_oils['Muatan Tolak'].append(data[i][j][25])
                dict_oils['Tujuan Kapal'].append(data[i][j][21])
                dict_oils['Jam Tolak'].append(data[i][j][23])
            elif data[i][j][4][:data[i][j][4].find('. ')] == 'TB':
                dict_oils['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_oils['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_oils['Keagenan'].append(data[i][j][13])
                dict_oils['Bendera'].append(data[i][j][12])
                dict_oils['GT'].append(data[i][j][5])
                dict_oils['Tanggal Tiba'].append(data[i][j][15])
                dict_oils['Tanggal Tolak'].append(data[i][j][22])
                dict_oils['Muatan Tiba'].append(data[i][j][18])
                dict_oils['Asal Kapal'].append(data[i][j][14])
                dict_oils['Muatan Tolak'].append(data[i][j][25])
                dict_oils['Tujuan Kapal'].append(data[i][j][21])
                dict_oils['Jam Tolak'].append(data[i][j][23])

    crudeoil = pd.DataFrame.from_dict(dict_oils)
    crudeoil.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return crudeoil

def heavies(data=None):
    dict_heavies = {
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
            if '(ALBE)' in data[i][j][18] or '(ALBE)' in data[i][j][25]:
                dict_heavies['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_heavies['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_heavies['Keagenan'].append(data[i][j][13])
                dict_heavies['Bendera'].append(data[i][j][12])
                dict_heavies['GT'].append(data[i][j][5])
                dict_heavies['Tanggal Tiba'].append(data[i][j][15])
                dict_heavies['Tanggal Tolak'].append(data[i][j][22])
                dict_heavies['Muatan Tiba'].append(data[i][j][18])
                dict_heavies['Asal Kapal'].append(data[i][j][14])
                dict_heavies['Muatan Tolak'].append(data[i][j][25])
                dict_heavies['Tujuan Kapal'].append(data[i][j][21])
                dict_heavies['Jam Tolak'].append(data[i][j][23])
            elif data[i][j][4][:data[i][j][4].find('. ')] == 'TB':
                dict_heavies['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_heavies['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_heavies['Keagenan'].append(data[i][j][13])
                dict_heavies['Bendera'].append(data[i][j][12])
                dict_heavies['GT'].append(data[i][j][5])
                dict_heavies['Tanggal Tiba'].append(data[i][j][15])
                dict_heavies['Tanggal Tolak'].append(data[i][j][22])
                dict_heavies['Muatan Tiba'].append(data[i][j][18])
                dict_heavies['Asal Kapal'].append(data[i][j][14])
                dict_heavies['Muatan Tolak'].append(data[i][j][25])
                dict_heavies['Tujuan Kapal'].append(data[i][j][21])
                dict_heavies['Jam Tolak'].append(data[i][j][23])
                
    heavy = pd.DataFrame.from_dict(dict_heavies)
    heavy.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return heavy

def bbm(data=None):
    dict_fuels = {
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
            if '(BBM)' in data[i][j][18] or '(BBM)' in data[i][j][25]:
                dict_fuels['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_fuels['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_fuels['Keagenan'].append(data[i][j][13])
                dict_fuels['Bendera'].append(data[i][j][12])
                dict_fuels['GT'].append(data[i][j][5])
                dict_fuels['Tanggal Tiba'].append(data[i][j][15])
                dict_fuels['Tanggal Tolak'].append(data[i][j][22])
                dict_fuels['Muatan Tiba'].append(data[i][j][18])
                dict_fuels['Asal Kapal'].append(data[i][j][14])
                dict_fuels['Muatan Tolak'].append(data[i][j][25])
                dict_fuels['Tujuan Kapal'].append(data[i][j][21])
                dict_fuels['Jam Tolak'].append(data[i][j][23])
            elif data[i][j][4][:data[i][j][4].find('. ')] == 'TB':
                dict_fuels['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_fuels['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_fuels['Keagenan'].append(data[i][j][13])
                dict_fuels['Bendera'].append(data[i][j][12])
                dict_fuels['GT'].append(data[i][j][5])
                dict_fuels['Tanggal Tiba'].append(data[i][j][15])
                dict_fuels['Tanggal Tolak'].append(data[i][j][22])
                dict_fuels['Muatan Tiba'].append(data[i][j][18])
                dict_fuels['Asal Kapal'].append(data[i][j][14])
                dict_fuels['Muatan Tolak'].append(data[i][j][25])
                dict_fuels['Tujuan Kapal'].append(data[i][j][21])
                dict_fuels['Jam Tolak'].append(data[i][j][23])

    fuels = pd.DataFrame.from_dict(dict_fuels)
    fuels.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return fuels

def kndr(data=None):
    dict_carbike = {
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
            if '(KNDR)' in data[i][j][18] or '(KNDR)' in data[i][j][25]:
                dict_carbike['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_carbike['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_carbike['Keagenan'].append(data[i][j][13])
                dict_carbike['Bendera'].append(data[i][j][12])
                dict_carbike['GT'].append(data[i][j][5])
                dict_carbike['Tanggal Tiba'].append(data[i][j][15])
                dict_carbike['Tanggal Tolak'].append(data[i][j][22])
                dict_carbike['Muatan Tiba'].append(data[i][j][18])
                dict_carbike['Asal Kapal'].append(data[i][j][14])
                dict_carbike['Muatan Tolak'].append(data[i][j][25])
                dict_carbike['Tujuan Kapal'].append(data[i][j][21])
                dict_carbike['Jam Tolak'].append(data[i][j][23])
            elif data[i][j][4][:data[i][j][4].find('. ')] == 'TB':
                dict_carbike['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_carbike['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_carbike['Keagenan'].append(data[i][j][13])
                dict_carbike['Bendera'].append(data[i][j][12])
                dict_carbike['GT'].append(data[i][j][5])
                dict_carbike['Tanggal Tiba'].append(data[i][j][15])
                dict_carbike['Tanggal Tolak'].append(data[i][j][22])
                dict_carbike['Muatan Tiba'].append(data[i][j][18])
                dict_carbike['Asal Kapal'].append(data[i][j][14])
                dict_carbike['Muatan Tolak'].append(data[i][j][25])
                dict_carbike['Tujuan Kapal'].append(data[i][j][21])
                dict_carbike['Jam Tolak'].append(data[i][j][23])

    carbike = pd.DataFrame.from_dict(dict_carbike)
    carbike.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return carbike

def kayu(data=None):
    dict_wood = {
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
            if '(KY)' in data[i][j][18] or '(KY)' in data[i][j][25]:
                dict_wood['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_wood['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_wood['Keagenan'].append(data[i][j][13])
                dict_wood['Bendera'].append(data[i][j][12])
                dict_wood['GT'].append(data[i][j][5])
                dict_wood['Tanggal Tiba'].append(data[i][j][15])
                dict_wood['Tanggal Tolak'].append(data[i][j][22])
                dict_wood['Muatan Tiba'].append(data[i][j][18])
                dict_wood['Asal Kapal'].append(data[i][j][14])
                dict_wood['Muatan Tolak'].append(data[i][j][25])
                dict_wood['Tujuan Kapal'].append(data[i][j][21])
                dict_wood['Jam Tolak'].append(data[i][j][23])
            elif data[i][j][4][:data[i][j][4].find('. ')] == 'TB':
                dict_wood['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_wood['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_wood['Keagenan'].append(data[i][j][13])
                dict_wood['Bendera'].append(data[i][j][12])
                dict_wood['GT'].append(data[i][j][5])
                dict_wood['Tanggal Tiba'].append(data[i][j][15])
                dict_wood['Tanggal Tolak'].append(data[i][j][22])
                dict_wood['Muatan Tiba'].append(data[i][j][18])
                dict_wood['Asal Kapal'].append(data[i][j][14])
                dict_wood['Muatan Tolak'].append(data[i][j][25])
                dict_wood['Tujuan Kapal'].append(data[i][j][21])
                dict_wood['Jam Tolak'].append(data[i][j][23])

    woody = pd.DataFrame.from_dict(dict_wood)
    woody.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return woody

def tanah(data=None):
    dict_sand = {
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
            if '(TNH)' in data[i][j][18] or '(TNH)' in data[i][j][25]:
                dict_sand['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_sand['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_sand['Keagenan'].append(data[i][j][13])
                dict_sand['Bendera'].append(data[i][j][12])
                dict_sand['GT'].append(data[i][j][5])
                dict_sand['Tanggal Tiba'].append(data[i][j][15])
                dict_sand['Tanggal Tolak'].append(data[i][j][22])
                dict_sand['Muatan Tiba'].append(data[i][j][18])
                dict_sand['Asal Kapal'].append(data[i][j][14])
                dict_sand['Muatan Tolak'].append(data[i][j][25])
                dict_sand['Tujuan Kapal'].append(data[i][j][21])
                dict_sand['Jam Tolak'].append(data[i][j][23])
            elif data[i][j][4][:data[i][j][4].find('. ')] == 'TB':
                dict_sand['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_sand['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_sand['Keagenan'].append(data[i][j][13])
                dict_sand['Bendera'].append(data[i][j][12])
                dict_sand['GT'].append(data[i][j][5])
                dict_sand['Tanggal Tiba'].append(data[i][j][15])
                dict_sand['Tanggal Tolak'].append(data[i][j][22])
                dict_sand['Muatan Tiba'].append(data[i][j][18])
                dict_sand['Asal Kapal'].append(data[i][j][14])
                dict_sand['Muatan Tolak'].append(data[i][j][25])
                dict_sand['Tujuan Kapal'].append(data[i][j][21])
                dict_sand['Jam Tolak'].append(data[i][j][23])
                
    sandy = pd.DataFrame.from_dict(dict_sand)
    sandy.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

    return sandy

def campursari(data=None):
    dict_mix = {
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
            if '(CMPR)' in data[i][j][18] or '(CMPR)' in data[i][j][25]:
                dict_mix['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_mix['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_mix['Keagenan'].append(data[i][j][13])
                dict_mix['Bendera'].append(data[i][j][12])
                dict_mix['GT'].append(data[i][j][5])
                dict_mix['Tanggal Tiba'].append(data[i][j][15])
                dict_mix['Tanggal Tolak'].append(data[i][j][22])
                dict_mix['Muatan Tiba'].append(data[i][j][18])
                dict_mix['Asal Kapal'].append(data[i][j][14])
                dict_mix['Muatan Tolak'].append(data[i][j][25])
                dict_mix['Tujuan Kapal'].append(data[i][j][21])
                dict_mix['Jam Tolak'].append(data[i][j][23])
            elif data[i][j][4][:data[i][j][4].find('. ')] == 'TB':
                dict_mix['Kode Kapal'].append(data[i][j][4][:data[i][j][4].find('. ')])
                dict_mix['Nama Kapal'].append(data[i][j][4][data[i][j][4].find('. ')+2:])
                dict_mix['Keagenan'].append(data[i][j][13])
                dict_mix['Bendera'].append(data[i][j][12])
                dict_mix['GT'].append(data[i][j][5])
                dict_mix['Tanggal Tiba'].append(data[i][j][15])
                dict_mix['Tanggal Tolak'].append(data[i][j][22])
                dict_mix['Muatan Tiba'].append(data[i][j][18])
                dict_mix['Asal Kapal'].append(data[i][j][14])
                dict_mix['Muatan Tolak'].append(data[i][j][25])
                dict_mix['Tujuan Kapal'].append(data[i][j][21])
                dict_mix['Jam Tolak'].append(data[i][j][23])

    mixy = pd.DataFrame.from_dict(dict_mix)
    mixy.sort_values(by=['Tanggal Tolak','Jam Tolak'],inplace=True)

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
