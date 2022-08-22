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
import numpy as np
import pandas as pd
from systemtools.number import *
from datetime import datetime as dt, timedelta as td
from translate import Translator
trnslt = Translator(to_lang='id')

# functions for general purpose

#1 read data
def read_excel_file(filename='None',col_title=4):
    dfname=pd.read_excel(filename,sheet_name='sheet_target',header=col_title)
    dfname.sort_values(by=['DATE','TIME'],inplace=True)
    dfname['NO'] = range(1,len(dfname)+1)
    dfname = dfname.reset_index(drop=True)
    #enable this code below, when you want to check the right amount of row
    #print('it has',len(dfname),'row(s)')

    return dfname

#2 find root index in a nested list
def nestedlist_rootindex(thelist, char1, char2):
    for i in range(len(thelist)):
        if char1 == thelist[i][0] and char2 == thelist[i][1]:
            return i
        
#3 categorizing data based on specific keys
def categorizing(dfreadfile):

    #categorizing by gross tonnage
    dfun500 = dfreadfile.loc[dfreadfile['GT'] <= 500]
    dfun500['NO'] = range(1,len(dfun500)+1)
    dfun500 = dfun500.reset_index(drop=True)
    dfab500 = dfreadfile.loc[dfreadfile['GT'] > 500]
    dfab500['NO'] = range(1,len(dfab500)+1)
    dfab500 = dfab500.reset_index(drop=True)

    #categorizing by destination type
    dfdom = dfreadfile.loc[dfreadfile['CATEGORY'] == 'DOMESTIC']
    dfdom['NO'] = range(1,len(dfdom)+1)
    dfdom = dfdom.reset_index(drop=True)
    dfexp = dfreadfile.loc[dfreadfile['CATEGORY'] == 'EXPORT']
    dfexp['NO'] = range(1,len(dfexp)+1)
    dfexp = dfexp.reset_index(drop=True)

    #categorizing by place registered
    dfpla = dfreadfile.loc[dfreadfile['NOTES'] == 'REGISTER LOCATION']
    dfpla['NO'] = range(1,len(dfpla)+1)
    dfpla = dfpla.reset_index(drop=True)
    
    return dfun500,dfab500,dfdom,dfexp,dfpla

#4 create load's summary, from all registered cargo
def goodstkii(data,mode):
    garname, garnum, garmea, depname, depnum, depmea = [],[],[],[],[],[]

    if mode == 'dom':
        for i in range(len(data['No'])):
            garname.append(data['Dom Load Out Nam'][i])
            garnum.append(data['Dom Load Out Num'][i])
            garmea.append(data['Dom Load Out Mea'][i])

            depname.append(data['Dom Load In Nam'][i])
            depnum.append(data['Dom Load In Num'][i])
            depmea.append(data['Dom Load In Mea'][i])
    
    elif mode == 'exp':
        for i in range(len(data['No'])):
            garname.append(data['Exp Load Out Nam'][i])
            garnum.append(data['Exp Load Out Num'][i])
            garmea.append(data['Exp Load Out Mea'][i])

            depname.append(data['Exp Load In Nam'][i])
            depnum.append(data['Exp Load In Num'][i])
            depmea.append(data['Exp Load In Mea'][i])

    #A little bit cleaning
    garname = [x for x in garname if x != 'NONE' and x != None]
    garnum = list(map(str,[x for x in garnum]))
    garnum = list(map(float,[parseNumber(x) for x in garnum if x != '--' and x != 'None']))
    garmea = [x for x in garmea if x != '--' and x != None]

    depname = [x for x in depname if x != 'NONE' and x != None]
    depnum = list(map(str,[x for x in depnum]))
    depnum = list(map(float,[parseNumber(x) for x in depnum if x != '--' and x != 'None']))
    depmea = [x for x in depmea if x != '--' and x != None]

    #calculating load's summary
    calarr = pd.DataFrame(list(zip(garname,garnum,garmea)), columns = ['Name', 'Number','Measure'])
    caldep = pd.DataFrame(list(zip(depname,depnum,depmea)), columns = ['Name', 'Number','Measure'])

    goar = calarr.groupby(['Name','Measure'],as_index=False).sum('Number')
    gode = caldep.groupby(['Name','Measure'],as_index=False).sum('Number')
    calarr,caldep = goar.values.tolist(),gode.values.tolist()

    for glist in [calarr,caldep]:
        for i in range(len(glist)):
            peek = str(glist[i][2])
            if len(peek[peek.find('.')+1:]) > 3:
                glist[i][2] = float(round(glist[i][2],3))
            elif peek[peek.find('.')+1:] == '0':
                glist[i][2] = int(glist[i][2])
            elif len(peek[peek.find('.')+1:]) <= 3:
                pass
            else:
                pass

    sumcag = [glist[i][0]+'-'+glist[i][1] for glist in [calarr,caldep] for i in range(len(glist))]
    sumcag = np.unique(sumcag).tolist()

    sumnam,sumar,sumde = [],[],[]
    for i in range(len(sumcag)):
        gname,gmea = sumcag[i].split('-')
        sumnam.append(gname)
        try:
            j = nestedlist_rootindex(calarr, gname, gmea)
            sumar.append(str(calarr[j][2])+' '+calarr[j][1])
        except TypeError:
            sumar.append('--')

        try:
            j = nestedlist_rootindex(caldep, gname, gmea)
            sumde.append(str(caldep[j][2])+' '+caldep[j][1])
        except TypeError:
            sumde.append('--')

    return [sumnam,sumar,sumde]

# functions for creating blank space in a certain format

# for TKII format
# Format number 01
def blankrows_tkii01(datadom):
    data = [datadom.columns.values.tolist()] + datadom.values.tolist()

    blankfmt_dom = {'No':[],'Kode Kapal':[],'Nama Kapal':[],'Bendera':[],'Keagenan':[],'GT':[],
                    'Tgl Tiba':[],'Jam Tiba':[],'Asal':[],'Tgl Tambat':[],'Jam Tambat':[],'Tgl Tolak':[],
                    'Tujuan':[],'Brg Bongkar D':[],'Jml Bongkar D':[],'1an Bongkar D':[],'Brg Muat D':[],
                    'Jml Muat D':[],'1an Muat D':[],'Brg Bongkar E':[],'Jml Bongkar E':[],'1an Bongkar E':[],
                    'Brg Muat E':[],'Jml Muat E':[],'1an Muat E':[],'KET':[]}

    for i in range(1,len(data)):
        if ';' in data[i][18] and ';' in data[i][29]:
            arr_load = data[i][18].split('; ')
            arr_num,arr_mu = data[i][19].split('; '),data[i][20].split('; ')

            depar_load = data[i][29].split('; ')
            depar_num,depar_mu = data[i][30].split('; '),data[i][31].split('; ')
            
            if len(arr_load) == len(depar_load):
                for p in range(len(arr_load)):
                    if isinstance(arr_load[p],str) and p == 0:
                        blankfmt_dom['No'].append(data[i][0])
                        blankfmt_dom['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                        blankfmt_dom['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                        blankfmt_dom['Bendera'].append(data[i][12])
                        blankfmt_dom['Keagenan'].append(data[i][13])
                        blankfmt_dom['GT'].append(data[i][5])
                        blankfmt_dom['Tgl Tiba'].append(data[i][15])
                        blankfmt_dom['Jam Tiba'].append(data[i][16])
                        blankfmt_dom['Asal'].append(data[i][14])
                        blankfmt_dom['Tgl Tambat'].append(data[i][15])
                        blankfmt_dom['Jam Tambat'].append(data[i][16])
                        blankfmt_dom['Tgl Tolak'].append(data[i][26])
                        blankfmt_dom['Tujuan'].append(data[i][25])
                        blankfmt_dom['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dom['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dom['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dom['Brg Muat D'].append(depar_load[p])
                        blankfmt_dom['Jml Muat D'].append(depar_num[p])
                        blankfmt_dom['1an Muat D'].append(depar_mu[p])
                        blankfmt_dom['Brg Bongkar E'].append(None)
                        blankfmt_dom['Jml Bongkar E'].append(None)
                        blankfmt_dom['1an Bongkar E'].append(None)
                        blankfmt_dom['Brg Muat E'].append(None)
                        blankfmt_dom['Jml Muat E'].append(None)
                        blankfmt_dom['1an Muat E'].append(None)
                        blankfmt_dom['KET'].append(None)
                    else:
                        blankfmt_dom['No'].append(None)
                        blankfmt_dom['Kode Kapal'].append(None)
                        blankfmt_dom['Nama Kapal'].append(None)
                        blankfmt_dom['Bendera'].append(None)
                        blankfmt_dom['Keagenan'].append(None)
                        blankfmt_dom['GT'].append(None)
                        blankfmt_dom['Tgl Tiba'].append(None)
                        blankfmt_dom['Jam Tiba'].append(None)
                        blankfmt_dom['Asal'].append(None)
                        blankfmt_dom['Tgl Tambat'].append(None)
                        blankfmt_dom['Jam Tambat'].append(None)
                        blankfmt_dom['Tgl Tolak'].append(None)
                        blankfmt_dom['Tujuan'].append(None)
                        blankfmt_dom['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dom['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dom['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dom['Brg Muat D'].append(depar_load[p])
                        blankfmt_dom['Jml Muat D'].append(depar_num[p])
                        blankfmt_dom['1an Muat D'].append(depar_mu[p])
                        blankfmt_dom['Brg Bongkar E'].append(None)
                        blankfmt_dom['Jml Bongkar E'].append(None)
                        blankfmt_dom['1an Bongkar E'].append(None)
                        blankfmt_dom['Brg Muat E'].append(None)
                        blankfmt_dom['Jml Muat E'].append(None)
                        blankfmt_dom['1an Muat E'].append(None)
                        blankfmt_dom['KET'].append(None)

            elif len(arr_load) < len(depar_load):
                arr_load.extend(np.full([len(depar_load)-len(arr_load),1],None))
                arr_num.extend(np.full([len(depar_num)-len(arr_num),1],None))
                arr_mu.extend(np.full([len(depar_mu)-len(arr_mu),1],None))

                for p in range(len(depar_load)):
                    if isinstance(arr_load[p],str) and p == 0:
                        blankfmt_dom['No'].append(data[i][0])
                        blankfmt_dom['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                        blankfmt_dom['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                        blankfmt_dom['Bendera'].append(data[i][12])
                        blankfmt_dom['Keagenan'].append(data[i][13])
                        blankfmt_dom['GT'].append(data[i][5])
                        blankfmt_dom['Tgl Tiba'].append(data[i][15])
                        blankfmt_dom['Jam Tiba'].append(data[i][16])
                        blankfmt_dom['Asal'].append(data[i][14])
                        blankfmt_dom['Tgl Tambat'].append(data[i][15])
                        blankfmt_dom['Jam Tambat'].append(data[i][16])
                        blankfmt_dom['Tgl Tolak'].append(data[i][26])
                        blankfmt_dom['Tujuan'].append(data[i][25])
                        blankfmt_dom['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dom['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dom['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dom['Brg Muat D'].append(depar_load[p])
                        blankfmt_dom['Jml Muat D'].append(depar_num[p])
                        blankfmt_dom['1an Muat D'].append(depar_mu[p])
                        blankfmt_dom['Brg Bongkar E'].append(None)
                        blankfmt_dom['Jml Bongkar E'].append(None)
                        blankfmt_dom['1an Bongkar E'].append(None)
                        blankfmt_dom['Brg Muat E'].append(None)
                        blankfmt_dom['Jml Muat E'].append(None)
                        blankfmt_dom['1an Muat E'].append(None)
                        blankfmt_dom['KET'].append(None)
                    elif isinstance(arr_load[p],str) and p != 0:
                        blankfmt_dom['No'].append(None)
                        blankfmt_dom['Kode Kapal'].append(None)
                        blankfmt_dom['Nama Kapal'].append(None)
                        blankfmt_dom['Bendera'].append(None)
                        blankfmt_dom['Keagenan'].append(None)
                        blankfmt_dom['GT'].append(None)
                        blankfmt_dom['Tgl Tiba'].append(None)
                        blankfmt_dom['Jam Tiba'].append(None)
                        blankfmt_dom['Asal'].append(None)
                        blankfmt_dom['Tgl Tambat'].append(None)
                        blankfmt_dom['Jam Tambat'].append(None)
                        blankfmt_dom['Tgl Tolak'].append(None)
                        blankfmt_dom['Tujuan'].append(None)
                        blankfmt_dom['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dom['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dom['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dom['Brg Muat D'].append(depar_load[p])
                        blankfmt_dom['Jml Muat D'].append(depar_num[p])
                        blankfmt_dom['1an Muat D'].append(depar_mu[p])
                        blankfmt_dom['Brg Bongkar E'].append(None)
                        blankfmt_dom['Jml Bongkar E'].append(None)
                        blankfmt_dom['1an Bongkar E'].append(None)
                        blankfmt_dom['Brg Muat E'].append(None)
                        blankfmt_dom['Jml Muat E'].append(None)
                        blankfmt_dom['1an Muat E'].append(None)
                        blankfmt_dom['KET'].append(None)
                    elif not isinstance(arr_load[p],str) and p != 0:
                        blankfmt_dom['No'].append(None)
                        blankfmt_dom['Kode Kapal'].append(None)
                        blankfmt_dom['Nama Kapal'].append(None)
                        blankfmt_dom['Bendera'].append(None)
                        blankfmt_dom['Keagenan'].append(None)
                        blankfmt_dom['GT'].append(None)
                        blankfmt_dom['Tgl Tiba'].append(None)
                        blankfmt_dom['Jam Tiba'].append(None)
                        blankfmt_dom['Asal'].append(None)
                        blankfmt_dom['Tgl Tambat'].append(None)
                        blankfmt_dom['Jam Tambat'].append(None)
                        blankfmt_dom['Tgl Tolak'].append(None)
                        blankfmt_dom['Tujuan'].append(None)
                        blankfmt_dom['Brg Bongkar D'].append(None)
                        blankfmt_dom['Jml Bongkar D'].append(None)
                        blankfmt_dom['1an Bongkar D'].append(None)
                        blankfmt_dom['Brg Muat D'].append(depar_load[p])
                        blankfmt_dom['Jml Muat D'].append(depar_num[p])
                        blankfmt_dom['1an Muat D'].append(depar_mu[p])
                        blankfmt_dom['Brg Bongkar E'].append(None)
                        blankfmt_dom['Jml Bongkar E'].append(None)
                        blankfmt_dom['1an Bongkar E'].append(None)
                        blankfmt_dom['Brg Muat E'].append(None)
                        blankfmt_dom['Jml Muat E'].append(None)
                        blankfmt_dom['1an Muat E'].append(None)
                        blankfmt_dom['KET'].append(None)
            
            elif len(arr_load) > len(depar_load):
                depar_load.extend(np.full([1,len(arr_load)-len(depar_load)],None))
                depar_num.extend(np.full([1,len(arr_num)-len(depar_num)],None))
                depar_mu.extend(np.full([1,len(arr_mu)-len(depar_mu)],None))

                for p in range(len(arr_load)):
                    if isinstance(depar_load[p],str) and p == 0:
                        blankfmt_dom['No'].append(data[i][0])
                        blankfmt_dom['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                        blankfmt_dom['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                        blankfmt_dom['Bendera'].append(data[i][12])
                        blankfmt_dom['Keagenan'].append(data[i][13])
                        blankfmt_dom['GT'].append(data[i][5])
                        blankfmt_dom['Tgl Tiba'].append(data[i][15])
                        blankfmt_dom['Jam Tiba'].append(data[i][16])
                        blankfmt_dom['Asal'].append(data[i][14])
                        blankfmt_dom['Tgl Tambat'].append(data[i][15])
                        blankfmt_dom['Jam Tambat'].append(data[i][16])
                        blankfmt_dom['Tgl Tolak'].append(data[i][26])
                        blankfmt_dom['Tujuan'].append(data[i][25])
                        blankfmt_dom['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dom['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dom['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dom['Brg Muat D'].append(depar_load[p])
                        blankfmt_dom['Jml Muat D'].append(depar_num[p])
                        blankfmt_dom['1an Muat D'].append(depar_mu[p])
                        blankfmt_dom['Brg Bongkar E'].append(None)
                        blankfmt_dom['Jml Bongkar E'].append(None)
                        blankfmt_dom['1an Bongkar E'].append(None)
                        blankfmt_dom['Brg Muat E'].append(None)
                        blankfmt_dom['Jml Muat E'].append(None)
                        blankfmt_dom['1an Muat E'].append(None)
                        blankfmt_dom['KET'].append(None)
                    elif isinstance(depar_load[p],str) and p != 0:
                        blankfmt_dom['No'].append(None)
                        blankfmt_dom['Kode Kapal'].append(None)
                        blankfmt_dom['Nama Kapal'].append(None)
                        blankfmt_dom['Bendera'].append(None)
                        blankfmt_dom['Keagenan'].append(None)
                        blankfmt_dom['GT'].append(None)
                        blankfmt_dom['Tgl Tiba'].append(None)
                        blankfmt_dom['Jam Tiba'].append(None)
                        blankfmt_dom['Asal'].append(None)
                        blankfmt_dom['Tgl Tambat'].append(None)
                        blankfmt_dom['Jam Tambat'].append(None)
                        blankfmt_dom['Tgl Tolak'].append(None)
                        blankfmt_dom['Tujuan'].append(None)
                        blankfmt_dom['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dom['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dom['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dom['Brg Muat D'].append(depar_load[p])
                        blankfmt_dom['Jml Muat D'].append(depar_num[p])
                        blankfmt_dom['1an Muat D'].append(depar_mu[p])
                        blankfmt_dom['Brg Bongkar E'].append(None)
                        blankfmt_dom['Jml Bongkar E'].append(None)
                        blankfmt_dom['1an Bongkar E'].append(None)
                        blankfmt_dom['Brg Muat E'].append(None)
                        blankfmt_dom['Jml Muat E'].append(None)
                        blankfmt_dom['1an Muat E'].append(None)
                        blankfmt_dom['KET'].append(None)
                    elif not isinstance(depar_load[p],str) and p != 0:
                        blankfmt_dom['No'].append(None)
                        blankfmt_dom['Kode Kapal'].append(None)
                        blankfmt_dom['Nama Kapal'].append(None)
                        blankfmt_dom['Bendera'].append(None)
                        blankfmt_dom['Keagenan'].append(None)
                        blankfmt_dom['GT'].append(None)
                        blankfmt_dom['Tgl Tiba'].append(None)
                        blankfmt_dom['Jam Tiba'].append(None)
                        blankfmt_dom['Asal'].append(None)
                        blankfmt_dom['Tgl Tambat'].append(None)
                        blankfmt_dom['Jam Tambat'].append(None)
                        blankfmt_dom['Tgl Tolak'].append(None)
                        blankfmt_dom['Tujuan'].append(None)
                        blankfmt_dom['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dom['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dom['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dom['Brg Muat D'].append(None)
                        blankfmt_dom['Jml Muat D'].append(None)
                        blankfmt_dom['1an Muat D'].append(None)
                        blankfmt_dom['Brg Bongkar E'].append(None)
                        blankfmt_dom['Jml Bongkar E'].append(None)
                        blankfmt_dom['1an Bongkar E'].append(None)
                        blankfmt_dom['Brg Muat E'].append(None)
                        blankfmt_dom['Jml Muat E'].append(None)
                        blankfmt_dom['1an Muat E'].append(None)
                        blankfmt_dom['KET'].append(None)
                        
        elif ';' in data[i][18] and ';' not in data[i][29]:
            arr_load = data[i][18].split('; ')
            arr_num = data[i][19].split('; ')
            arr_mu = data[i][20].split('; ')

            for p in range(len(arr_load)):
                if p == 0:
                    blankfmt_dom['No'].append(data[i][0])
                    blankfmt_dom['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                    blankfmt_dom['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                    blankfmt_dom['Bendera'].append(data[i][12])
                    blankfmt_dom['Keagenan'].append(data[i][13])
                    blankfmt_dom['GT'].append(data[i][5])
                    blankfmt_dom['Tgl Tiba'].append(data[i][15])
                    blankfmt_dom['Jam Tiba'].append(data[i][16])
                    blankfmt_dom['Asal'].append(data[i][14])
                    blankfmt_dom['Tgl Tambat'].append(data[i][15])
                    blankfmt_dom['Jam Tambat'].append(data[i][16])
                    blankfmt_dom['Tgl Tolak'].append(data[i][26])
                    blankfmt_dom['Tujuan'].append(data[i][25])
                    blankfmt_dom['Brg Bongkar D'].append(arr_load[p])
                    blankfmt_dom['Jml Bongkar D'].append(arr_num[p])
                    blankfmt_dom['1an Bongkar D'].append(arr_mu[p])
                    blankfmt_dom['Brg Muat D'].append(data[i][29])
                    blankfmt_dom['Jml Muat D'].append(data[i][30])
                    blankfmt_dom['1an Muat D'].append(data[i][31])
                    blankfmt_dom['Brg Bongkar E'].append(None)
                    blankfmt_dom['Jml Bongkar E'].append(None)
                    blankfmt_dom['1an Bongkar E'].append(None)
                    blankfmt_dom['Brg Muat E'].append(None)
                    blankfmt_dom['Jml Muat E'].append(None)
                    blankfmt_dom['1an Muat E'].append(None)
                    blankfmt_dom['KET'].append(None)
                elif p != 0:
                    blankfmt_dom['No'].append(None)
                    blankfmt_dom['Kode Kapal'].append(None)
                    blankfmt_dom['Nama Kapal'].append(None)
                    blankfmt_dom['Bendera'].append(None)
                    blankfmt_dom['Keagenan'].append(None)
                    blankfmt_dom['GT'].append(None)
                    blankfmt_dom['Tgl Tiba'].append(None)
                    blankfmt_dom['Jam Tiba'].append(None)
                    blankfmt_dom['Asal'].append(None)
                    blankfmt_dom['Tgl Tambat'].append(None)
                    blankfmt_dom['Jam Tambat'].append(None)
                    blankfmt_dom['Tgl Tolak'].append(None)
                    blankfmt_dom['Tujuan'].append(None)
                    blankfmt_dom['Brg Bongkar D'].append(arr_load[p])
                    blankfmt_dom['Jml Bongkar D'].append(arr_num[p])
                    blankfmt_dom['1an Bongkar D'].append(arr_mu[p])
                    blankfmt_dom['Brg Muat D'].append(None)
                    blankfmt_dom['Jml Muat D'].append(None)
                    blankfmt_dom['1an Muat D'].append(None)
                    blankfmt_dom['Brg Bongkar E'].append(None)
                    blankfmt_dom['Jml Bongkar E'].append(None)
                    blankfmt_dom['1an Bongkar E'].append(None)
                    blankfmt_dom['Brg Muat E'].append(None)
                    blankfmt_dom['Jml Muat E'].append(None)
                    blankfmt_dom['1an Muat E'].append(None)
                    blankfmt_dom['KET'].append(None)

        elif ';' not in data[i][18] and ';' in data[i][29]:
            depar_load = data[i][29].split('; ')
            depar_num = data[i][30].split('; ')
            depar_mu = data[i][31].split('; ')

            for p in range(len(depar_load)):
                if p == 0:
                    blankfmt_dom['No'].append(data[i][0])
                    blankfmt_dom['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                    blankfmt_dom['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                    blankfmt_dom['Bendera'].append(data[i][12])
                    blankfmt_dom['Keagenan'].append(data[i][13])
                    blankfmt_dom['GT'].append(data[i][5])
                    blankfmt_dom['Tgl Tiba'].append(data[i][15])
                    blankfmt_dom['Jam Tiba'].append(data[i][16])
                    blankfmt_dom['Asal'].append(data[i][14])
                    blankfmt_dom['Tgl Tambat'].append(data[i][15])
                    blankfmt_dom['Jam Tambat'].append(data[i][16])
                    blankfmt_dom['Tgl Tolak'].append(data[i][26])
                    blankfmt_dom['Tujuan'].append(data[i][25])
                    blankfmt_dom['Brg Bongkar D'].append(data[i][18])
                    blankfmt_dom['Jml Bongkar D'].append(data[i][19])
                    blankfmt_dom['1an Bongkar D'].append(data[i][20])
                    blankfmt_dom['Brg Muat D'].append(depar_load[p])
                    blankfmt_dom['Jml Muat D'].append(depar_num[p])
                    blankfmt_dom['1an Muat D'].append(depar_mu[p])
                    blankfmt_dom['Brg Bongkar E'].append(None)
                    blankfmt_dom['Jml Bongkar E'].append(None)
                    blankfmt_dom['1an Bongkar E'].append(None)
                    blankfmt_dom['Brg Muat E'].append(None)
                    blankfmt_dom['Jml Muat E'].append(None)
                    blankfmt_dom['1an Muat E'].append(None)
                    blankfmt_dom['KET'].append(None)
                elif p != 0:
                    blankfmt_dom['No'].append(None)
                    blankfmt_dom['Kode Kapal'].append(None)
                    blankfmt_dom['Nama Kapal'].append(None)
                    blankfmt_dom['Bendera'].append(None)
                    blankfmt_dom['Keagenan'].append(None)
                    blankfmt_dom['GT'].append(None)
                    blankfmt_dom['Tgl Tiba'].append(None)
                    blankfmt_dom['Jam Tiba'].append(None)
                    blankfmt_dom['Asal'].append(None)
                    blankfmt_dom['Tgl Tambat'].append(None)
                    blankfmt_dom['Jam Tambat'].append(None)
                    blankfmt_dom['Tgl Tolak'].append(None)
                    blankfmt_dom['Tujuan'].append(None)
                    blankfmt_dom['Brg Bongkar D'].append(None)
                    blankfmt_dom['Jml Bongkar D'].append(None)
                    blankfmt_dom['1an Bongkar D'].append(None)
                    blankfmt_dom['Brg Muat D'].append(depar_load[p])
                    blankfmt_dom['Jml Muat D'].append(depar_num[p])
                    blankfmt_dom['1an Muat D'].append(depar_mu[p])
                    blankfmt_dom['Brg Bongkar E'].append(None)
                    blankfmt_dom['Jml Bongkar E'].append(None)
                    blankfmt_dom['1an Bongkar E'].append(None)
                    blankfmt_dom['Brg Muat E'].append(None)
                    blankfmt_dom['Jml Muat E'].append(None)
                    blankfmt_dom['1an Muat E'].append(None)
                    blankfmt_dom['KET'].append(None)

        elif ';' not in data[i][18] and ';' not in data[i][29]:
            blankfmt_dom['No'].append(data[i][0])
            blankfmt_dom['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
            blankfmt_dom['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
            blankfmt_dom['Bendera'].append(data[i][12])
            blankfmt_dom['Keagenan'].append(data[i][13])
            blankfmt_dom['GT'].append(data[i][5])
            blankfmt_dom['Tgl Tiba'].append(data[i][15])
            blankfmt_dom['Jam Tiba'].append(data[i][16])
            blankfmt_dom['Asal'].append(data[i][14])
            blankfmt_dom['Tgl Tambat'].append(data[i][15])
            blankfmt_dom['Jam Tambat'].append(data[i][16])
            blankfmt_dom['Tgl Tolak'].append(data[i][26])
            blankfmt_dom['Tujuan'].append(data[i][25])
            blankfmt_dom['Brg Bongkar D'].append(data[i][18])
            blankfmt_dom['Jml Bongkar D'].append(data[i][19])
            blankfmt_dom['1an Bongkar D'].append(data[i][20])
            blankfmt_dom['Brg Muat D'].append(data[i][29])
            blankfmt_dom['Jml Muat D'].append(data[i][30])
            blankfmt_dom['1an Muat D'].append(data[i][31])
            blankfmt_dom['Brg Bongkar E'].append(None)
            blankfmt_dom['Jml Bongkar E'].append(None)
            blankfmt_dom['1an Bongkar E'].append(None)
            blankfmt_dom['Brg Muat E'].append(None)
            blankfmt_dom['Jml Muat E'].append(None)
            blankfmt_dom['1an Muat E'].append(None)
            blankfmt_dom['KET'].append(None)

    return blankfmt_dom

# Format Number 02
def blankrows_tkii02(dataexp):
    data = [dataexp.columns.values.tolist()] + dataexp.values.tolist()

    blankfmt_exp = {'No':[],'Kode Kapal':[],'Nama Kapal':[],'Bendera':[],'Keagenan':[],'GT':[],
                    'Tgl Tiba':[],'Jam Tiba':[],'Asal':[],'Tgl Tambat':[],'Jam Tambat':[],'Tgl Tolak':[],
                    'Tujuan':[],'Brg Bongkar D':[],'Jml Bongkar D':[],'1an Bongkar D':[],'Brg Muat D':[],
                    'Jml Muat D':[],'1an Muat D':[],'Brg Bongkar E':[],'Jml Bongkar E':[],'1an Bongkar E':[],
                    'Brg Muat E':[],'Jml Muat E':[],'1an Muat E':[],'KET':[]}

    for i in range(1,len(data)):
        if ';' in data[i][18] and ';' in data[i][29]:
            arr_load = data[i][18].split('; ')
            arr_num,arr_mu = data[i][19].split('; '),data[i][20].split('; ')

            depar_load = data[i][29].split('; ')
            depar_num,depar_mu = data[i][30].split('; '),data[i][31].split('; ')
            
            if len(arr_load) == len(depar_load):
                for p in range(len(arr_load)):
                    if isinstance(arr_load[p],str) and p == 0:
                        blankfmt_exp['No'].append(data[i][0])
                        blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                        blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                        blankfmt_exp['Bendera'].append(data[i][12])
                        blankfmt_exp['Keagenan'].append(data[i][13])
                        blankfmt_exp['GT'].append(data[i][5])
                        blankfmt_exp['Tgl Tiba'].append(data[i][15])
                        blankfmt_exp['Jam Tiba'].append(data[i][16])
                        blankfmt_exp['Asal'].append(data[i][14])
                        blankfmt_exp['Tgl Tambat'].append(data[i][15])
                        blankfmt_exp['Jam Tambat'].append(data[i][16])
                        blankfmt_exp['Tgl Tolak'].append(data[i][26])
                        blankfmt_exp['Tujuan'].append(data[i][25])
                        blankfmt_exp['Brg Bongkar D'].append(None)
                        blankfmt_exp['Jml Bongkar D'].append(None)
                        blankfmt_exp['1an Bongkar D'].append(None)
                        blankfmt_exp['Brg Muat D'].append(None)
                        blankfmt_exp['Jml Muat D'].append(None)
                        blankfmt_exp['1an Muat D'].append(None)
                        blankfmt_exp['Brg Bongkar E'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar E'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar E'].append(arr_mu[p])
                        blankfmt_exp['Brg Muat E'].append(depar_load[p])
                        blankfmt_exp['Jml Muat E'].append(depar_num[p])
                        blankfmt_exp['1an Muat E'].append(depar_mu[p])
                        blankfmt_exp['KET'].append(None)
                    else:
                        blankfmt_exp['No'].append(None)
                        blankfmt_exp['Kode Kapal'].append(None)
                        blankfmt_exp['Nama Kapal'].append(None)
                        blankfmt_exp['Bendera'].append(None)
                        blankfmt_exp['Keagenan'].append(None)
                        blankfmt_exp['GT'].append(None)
                        blankfmt_exp['Tgl Tiba'].append(None)
                        blankfmt_exp['Jam Tiba'].append(None)
                        blankfmt_exp['Asal'].append(None)
                        blankfmt_exp['Tgl Tambat'].append(None)
                        blankfmt_exp['Jam Tambat'].append(None)
                        blankfmt_exp['Tgl Tolak'].append(None)
                        blankfmt_exp['Tujuan'].append(None)
                        blankfmt_exp['Brg Bongkar D'].append(None)
                        blankfmt_exp['Jml Bongkar D'].append(None)
                        blankfmt_exp['1an Bongkar D'].append(None)
                        blankfmt_exp['Brg Muat D'].append(None)
                        blankfmt_exp['Jml Muat D'].append(None)
                        blankfmt_exp['1an Muat D'].append(None)
                        blankfmt_exp['Brg Bongkar E'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar E'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar E'].append(arr_mu[p])
                        blankfmt_exp['Brg Muat E'].append(depar_load[p])
                        blankfmt_exp['Jml Muat E'].append(depar_num[p])
                        blankfmt_exp['1an Muat E'].append(depar_mu[p])
                        blankfmt_exp['KET'].append(None)

            elif len(arr_load) < len(depar_load):
                arr_load.extend(np.full([len(depar_load)-len(arr_load),1],None))
                arr_num.extend(np.full([len(depar_num)-len(arr_num),1],None))
                arr_mu.extend(np.full([len(depar_mu)-len(arr_mu),1],None))

                for p in range(len(depar_load)):
                    if isinstance(arr_load[p],str) and p == 0:
                        blankfmt_exp['No'].append(data[i][0])
                        blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                        blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                        blankfmt_exp['Bendera'].append(data[i][12])
                        blankfmt_exp['Keagenan'].append(data[i][13])
                        blankfmt_exp['GT'].append(data[i][5])
                        blankfmt_exp['Tgl Tiba'].append(data[i][15])
                        blankfmt_exp['Jam Tiba'].append(data[i][16])
                        blankfmt_exp['Asal'].append(data[i][14])
                        blankfmt_exp['Tgl Tambat'].append(data[i][15])
                        blankfmt_exp['Jam Tambat'].append(data[i][16])
                        blankfmt_exp['Tgl Tolak'].append(data[i][26])
                        blankfmt_exp['Tujuan'].append(data[i][25])
                        blankfmt_exp['Brg Bongkar D'].append(None)
                        blankfmt_exp['Jml Bongkar D'].append(None)
                        blankfmt_exp['1an Bongkar D'].append(None)
                        blankfmt_exp['Brg Muat D'].append(None)
                        blankfmt_exp['Jml Muat D'].append(None)
                        blankfmt_exp['1an Muat D'].append(None)
                        blankfmt_exp['Brg Bongkar E'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar E'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar E'].append(arr_mu[p])
                        blankfmt_exp['Brg Muat E'].append(depar_load[p])
                        blankfmt_exp['Jml Muat E'].append(depar_num[p])
                        blankfmt_exp['1an Muat E'].append(depar_mu[p])
                        blankfmt_exp['KET'].append(None)
                    elif isinstance(arr_load[p],str) and p != 0:
                        blankfmt_exp['No'].append(None)
                        blankfmt_exp['Kode Kapal'].append(None)
                        blankfmt_exp['Nama Kapal'].append(None)
                        blankfmt_exp['Bendera'].append(None)
                        blankfmt_exp['Keagenan'].append(None)
                        blankfmt_exp['GT'].append(None)
                        blankfmt_exp['Tgl Tiba'].append(None)
                        blankfmt_exp['Jam Tiba'].append(None)
                        blankfmt_exp['Asal'].append(None)
                        blankfmt_exp['Tgl Tambat'].append(None)
                        blankfmt_exp['Jam Tambat'].append(None)
                        blankfmt_exp['Tgl Tolak'].append(None)
                        blankfmt_exp['Tujuan'].append(None)
                        blankfmt_exp['Brg Bongkar D'].append(None)
                        blankfmt_exp['Jml Bongkar D'].append(None)
                        blankfmt_exp['1an Bongkar D'].append(None)
                        blankfmt_exp['Brg Muat D'].append(None)
                        blankfmt_exp['Jml Muat D'].append(None)
                        blankfmt_exp['1an Muat D'].append(None)
                        blankfmt_exp['Brg Bongkar E'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar E'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar E'].append(arr_mu[p])
                        blankfmt_exp['Brg Muat E'].append(depar_load[p])
                        blankfmt_exp['Jml Muat E'].append(depar_num[p])
                        blankfmt_exp['1an Muat E'].append(depar_mu[p])
                        blankfmt_exp['KET'].append(None)
                    elif not isinstance(arr_load[p],str) and p != 0:
                        blankfmt_exp['No'].append(None)
                        blankfmt_exp['Kode Kapal'].append(None)
                        blankfmt_exp['Nama Kapal'].append(None)
                        blankfmt_exp['Bendera'].append(None)
                        blankfmt_exp['Keagenan'].append(None)
                        blankfmt_exp['GT'].append(None)
                        blankfmt_exp['Tgl Tiba'].append(None)
                        blankfmt_exp['Jam Tiba'].append(None)
                        blankfmt_exp['Asal'].append(None)
                        blankfmt_exp['Tgl Tambat'].append(None)
                        blankfmt_exp['Jam Tambat'].append(None)
                        blankfmt_exp['Tgl Tolak'].append(None)
                        blankfmt_exp['Tujuan'].append(None)
                        blankfmt_exp['Brg Bongkar D'].append(None)
                        blankfmt_exp['Jml Bongkar D'].append(None)
                        blankfmt_exp['1an Bongkar D'].append(None)
                        blankfmt_exp['Brg Muat D'].append(None)
                        blankfmt_exp['Jml Muat D'].append(None)
                        blankfmt_exp['1an Muat D'].append(None)
                        blankfmt_exp['Brg Bongkar E'].append(None)
                        blankfmt_exp['Jml Bongkar E'].append(None)
                        blankfmt_exp['1an Bongkar E'].append(None)
                        blankfmt_exp['Brg Muat E'].append(depar_load[p])
                        blankfmt_exp['Jml Muat E'].append(depar_num[p])
                        blankfmt_exp['1an Muat E'].append(depar_mu[p])
                        blankfmt_exp['KET'].append(None)
            
            elif len(arr_load) > len(depar_load):
                depar_load.extend(np.full([1,len(arr_load)-len(depar_load)],None))
                depar_num.extend(np.full([1,len(arr_num)-len(depar_num)],None))
                depar_mu.extend(np.full([1,len(arr_mu)-len(depar_mu)],None))

                for p in range(len(arr_load)):
                    if isinstance(depar_load[p],str) and p == 0:
                        blankfmt_exp['No'].append(data[i][0])
                        blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                        blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                        blankfmt_exp['Bendera'].append(data[i][12])
                        blankfmt_exp['Keagenan'].append(data[i][13])
                        blankfmt_exp['GT'].append(data[i][5])
                        blankfmt_exp['Tgl Tiba'].append(data[i][15])
                        blankfmt_exp['Jam Tiba'].append(data[i][16])
                        blankfmt_exp['Asal'].append(data[i][14])
                        blankfmt_exp['Tgl Tambat'].append(data[i][15])
                        blankfmt_exp['Jam Tambat'].append(data[i][16])
                        blankfmt_exp['Tgl Tolak'].append(data[i][26])
                        blankfmt_exp['Tujuan'].append(data[i][25])
                        blankfmt_exp['Brg Bongkar D'].append(None)
                        blankfmt_exp['Jml Bongkar D'].append(None)
                        blankfmt_exp['1an Bongkar D'].append(None)
                        blankfmt_exp['Brg Muat D'].append(None)
                        blankfmt_exp['Jml Muat D'].append(None)
                        blankfmt_exp['1an Muat D'].append(None)
                        blankfmt_exp['Brg Bongkar E'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar E'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar E'].append(arr_mu[p])
                        blankfmt_exp['Brg Muat E'].append(depar_load[p])
                        blankfmt_exp['Jml Muat E'].append(depar_num[p])
                        blankfmt_exp['1an Muat E'].append(depar_mu[p])
                        blankfmt_exp['KET'].append(None)
                    elif isinstance(depar_load[p],str) and p != 0:
                        blankfmt_exp['No'].append(None)
                        blankfmt_exp['Kode Kapal'].append(None)
                        blankfmt_exp['Nama Kapal'].append(None)
                        blankfmt_exp['Bendera'].append(None)
                        blankfmt_exp['Keagenan'].append(None)
                        blankfmt_exp['GT'].append(None)
                        blankfmt_exp['Tgl Tiba'].append(None)
                        blankfmt_exp['Jam Tiba'].append(None)
                        blankfmt_exp['Asal'].append(None)
                        blankfmt_exp['Tgl Tambat'].append(None)
                        blankfmt_exp['Jam Tambat'].append(None)
                        blankfmt_exp['Tgl Tolak'].append(None)
                        blankfmt_exp['Tujuan'].append(None)
                        blankfmt_exp['Brg Bongkar D'].append(None)
                        blankfmt_exp['Jml Bongkar D'].append(None)
                        blankfmt_exp['1an Bongkar D'].append(None)
                        blankfmt_exp['Brg Muat D'].append(None)
                        blankfmt_exp['Jml Muat D'].append(None)
                        blankfmt_exp['1an Muat D'].append(None)
                        blankfmt_exp['Brg Bongkar E'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar E'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar E'].append(arr_mu[p])
                        blankfmt_exp['Brg Muat E'].append(depar_load[p])
                        blankfmt_exp['Jml Muat E'].append(depar_num[p])
                        blankfmt_exp['1an Muat E'].append(depar_mu[p])
                        blankfmt_exp['KET'].append(None)
                    elif not isinstance(depar_load[p],str) and p != 0:
                        blankfmt_exp['No'].append(None)
                        blankfmt_exp['Kode Kapal'].append(None)
                        blankfmt_exp['Nama Kapal'].append(None)
                        blankfmt_exp['Bendera'].append(None)
                        blankfmt_exp['Keagenan'].append(None)
                        blankfmt_exp['GT'].append(None)
                        blankfmt_exp['Tgl Tiba'].append(None)
                        blankfmt_exp['Jam Tiba'].append(None)
                        blankfmt_exp['Asal'].append(None)
                        blankfmt_exp['Tgl Tambat'].append(None)
                        blankfmt_exp['Jam Tambat'].append(None)
                        blankfmt_exp['Tgl Tolak'].append(None)
                        blankfmt_exp['Tujuan'].append(None)
                        blankfmt_exp['Brg Bongkar D'].append(None)
                        blankfmt_exp['Jml Bongkar D'].append(None)
                        blankfmt_exp['1an Bongkar D'].append(None)
                        blankfmt_exp['Brg Muat D'].append(None)
                        blankfmt_exp['Jml Muat D'].append(None)
                        blankfmt_exp['1an Muat D'].append(None)
                        blankfmt_exp['Brg Bongkar E'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar E'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar E'].append(arr_mu[p])
                        blankfmt_exp['Brg Muat E'].append(None)
                        blankfmt_exp['Jml Muat E'].append(None)
                        blankfmt_exp['1an Muat E'].append(None)
                        blankfmt_exp['KET'].append(None)
                        
        elif ';' in data[i][18] and ';' not in data[i][29]:
            arr_load = data[i][18].split('; ')
            arr_num = data[i][19].split('; ')
            arr_mu = data[i][20].split('; ')

            for p in range(len(arr_load)):
                if p == 0:
                    blankfmt_exp['No'].append(data[i][0])
                    blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                    blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                    blankfmt_exp['Bendera'].append(data[i][12])
                    blankfmt_exp['Keagenan'].append(data[i][13])
                    blankfmt_exp['GT'].append(data[i][5])
                    blankfmt_exp['Tgl Tiba'].append(data[i][15])
                    blankfmt_exp['Jam Tiba'].append(data[i][16])
                    blankfmt_exp['Asal'].append(data[i][14])
                    blankfmt_exp['Tgl Tambat'].append(data[i][15])
                    blankfmt_exp['Jam Tambat'].append(data[i][16])
                    blankfmt_exp['Tgl Tolak'].append(data[i][26])
                    blankfmt_exp['Tujuan'].append(data[i][25])
                    blankfmt_exp['Brg Bongkar D'].append(None)
                    blankfmt_exp['Jml Bongkar D'].append(None)
                    blankfmt_exp['1an Bongkar D'].append(None)
                    blankfmt_exp['Brg Muat D'].append(None)
                    blankfmt_exp['Jml Muat D'].append(None)
                    blankfmt_exp['1an Muat D'].append(None)
                    blankfmt_exp['Brg Bongkar E'].append(arr_load[p])
                    blankfmt_exp['Jml Bongkar E'].append(arr_num[p])
                    blankfmt_exp['1an Bongkar E'].append(arr_mu[p])
                    blankfmt_exp['Brg Muat E'].append(data[i][29])
                    blankfmt_exp['Jml Muat E'].append(data[i][30])
                    blankfmt_exp['1an Muat E'].append(data[i][31])
                    blankfmt_exp['KET'].append(None)
                elif p != 0:
                    blankfmt_exp['No'].append(None)
                    blankfmt_exp['Kode Kapal'].append(None)
                    blankfmt_exp['Nama Kapal'].append(None)
                    blankfmt_exp['Bendera'].append(None)
                    blankfmt_exp['Keagenan'].append(None)
                    blankfmt_exp['GT'].append(None)
                    blankfmt_exp['Tgl Tiba'].append(None)
                    blankfmt_exp['Jam Tiba'].append(None)
                    blankfmt_exp['Asal'].append(None)
                    blankfmt_exp['Tgl Tambat'].append(None)
                    blankfmt_exp['Jam Tambat'].append(None)
                    blankfmt_exp['Tgl Tolak'].append(None)
                    blankfmt_exp['Tujuan'].append(None)
                    blankfmt_exp['Brg Bongkar D'].append(None)
                    blankfmt_exp['Jml Bongkar D'].append(None)
                    blankfmt_exp['1an Bongkar D'].append(None)
                    blankfmt_exp['Brg Muat D'].append(None)
                    blankfmt_exp['Jml Muat D'].append(None)
                    blankfmt_exp['1an Muat D'].append(None)
                    blankfmt_exp['Brg Bongkar E'].append(arr_load[p])
                    blankfmt_exp['Jml Bongkar E'].append(arr_num[p])
                    blankfmt_exp['1an Bongkar E'].append(arr_mu[p])
                    blankfmt_exp['Brg Muat E'].append(None)
                    blankfmt_exp['Jml Muat E'].append(None)
                    blankfmt_exp['1an Muat E'].append(None)
                    blankfmt_exp['KET'].append(None)

        elif ';' not in data[i][18] and ';' in data[i][29]:
            depar_load = data[i][29].split('; ')
            depar_num = data[i][30].split('; ')
            depar_mu = data[i][31].split('; ')

            for p in range(len(depar_load)):
                if p == 0:
                    blankfmt_exp['No'].append(data[i][0])
                    blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                    blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                    blankfmt_exp['Bendera'].append(data[i][12])
                    blankfmt_exp['Keagenan'].append(data[i][13])
                    blankfmt_exp['GT'].append(data[i][5])
                    blankfmt_exp['Tgl Tiba'].append(data[i][15])
                    blankfmt_exp['Jam Tiba'].append(data[i][16])
                    blankfmt_exp['Asal'].append(data[i][14])
                    blankfmt_exp['Tgl Tambat'].append(data[i][15])
                    blankfmt_exp['Jam Tambat'].append(data[i][16])
                    blankfmt_exp['Tgl Tolak'].append(data[i][26])
                    blankfmt_exp['Tujuan'].append(data[i][25])
                    blankfmt_exp['Brg Bongkar D'].append(None)
                    blankfmt_exp['Jml Bongkar D'].append(None)
                    blankfmt_exp['1an Bongkar D'].append(None)
                    blankfmt_exp['Brg Muat D'].append(None)
                    blankfmt_exp['Jml Muat D'].append(None)
                    blankfmt_exp['1an Muat D'].append(None)
                    blankfmt_exp['Brg Bongkar E'].append(data[i][18])
                    blankfmt_exp['Jml Bongkar E'].append(data[i][19])
                    blankfmt_exp['1an Bongkar E'].append(data[i][20])
                    blankfmt_exp['Brg Muat E'].append(depar_load[p])
                    blankfmt_exp['Jml Muat E'].append(depar_num[p])
                    blankfmt_exp['1an Muat E'].append(depar_mu[p])
                    blankfmt_exp['KET'].append(None)
                elif p != 0:
                    blankfmt_exp['No'].append(None)
                    blankfmt_exp['Kode Kapal'].append(None)
                    blankfmt_exp['Nama Kapal'].append(None)
                    blankfmt_exp['Bendera'].append(None)
                    blankfmt_exp['Keagenan'].append(None)
                    blankfmt_exp['GT'].append(None)
                    blankfmt_exp['Tgl Tiba'].append(None)
                    blankfmt_exp['Jam Tiba'].append(None)
                    blankfmt_exp['Asal'].append(None)
                    blankfmt_exp['Tgl Tambat'].append(None)
                    blankfmt_exp['Jam Tambat'].append(None)
                    blankfmt_exp['Tgl Tolak'].append(None)
                    blankfmt_exp['Tujuan'].append(None)
                    blankfmt_exp['Brg Bongkar D'].append(None)
                    blankfmt_exp['Jml Bongkar D'].append(None)
                    blankfmt_exp['1an Bongkar D'].append(None)
                    blankfmt_exp['Brg Muat D'].append(None)
                    blankfmt_exp['Jml Muat D'].append(None)
                    blankfmt_exp['1an Muat D'].append(None)
                    blankfmt_exp['Brg Bongkar E'].append(None)
                    blankfmt_exp['Jml Bongkar E'].append(None)
                    blankfmt_exp['1an Bongkar E'].append(None)
                    blankfmt_exp['Brg Muat E'].append(depar_load[p])
                    blankfmt_exp['Jml Muat E'].append(depar_num[p])
                    blankfmt_exp['1an Muat E'].append(depar_mu[p])
                    blankfmt_exp['KET'].append(None)

        elif ';' not in data[i][18] and ';' not in data[i][29]:
            blankfmt_exp['No'].append(data[i][0])
            blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
            blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
            blankfmt_exp['Bendera'].append(data[i][12])
            blankfmt_exp['Keagenan'].append(data[i][13])
            blankfmt_exp['GT'].append(data[i][5])
            blankfmt_exp['Tgl Tiba'].append(data[i][15])
            blankfmt_exp['Jam Tiba'].append(data[i][16])
            blankfmt_exp['Asal'].append(data[i][14])
            blankfmt_exp['Tgl Tambat'].append(data[i][15])
            blankfmt_exp['Jam Tambat'].append(data[i][16])
            blankfmt_exp['Tgl Tolak'].append(data[i][26])
            blankfmt_exp['Tujuan'].append(data[i][25])
            blankfmt_exp['Brg Bongkar D'].append(None)
            blankfmt_exp['Jml Bongkar D'].append(None)
            blankfmt_exp['1an Bongkar D'].append(None)
            blankfmt_exp['Brg Muat D'].append(None)
            blankfmt_exp['Jml Muat D'].append(None)
            blankfmt_exp['1an Muat D'].append(None)
            blankfmt_exp['Brg Bongkar E'].append(data[i][18])
            blankfmt_exp['Jml Bongkar E'].append(data[i][19])
            blankfmt_exp['1an Bongkar E'].append(data[i][20])
            blankfmt_exp['Brg Muat E'].append(data[i][29])
            blankfmt_exp['Jml Muat E'].append(data[i][30])
            blankfmt_exp['1an Muat E'].append(data[i][31])
            blankfmt_exp['KET'].append(None)

    return blankfmt_exp

# Format Number 03
def blankrows_tkii03(dataexp):
    data = [dataexp.columns.values.tolist()] + dataexp.values.tolist()

    blankfmt_exp = {'No':[],'Kode Kapal':[],'Nama Kapal':[],'Bendera':[],'Keagenan':[],'GT':[],'Tgl Tiba':[],
                    'Jam Tiba':[],'Asal':[],'Tgl Tambat':[],'Jam Tambat':[],'Tgl Tolak':[],'Tujuan':[]}

    for i in range(1,len(data)):
        if data[i][18] == 'NIHIL' and data[i][29] == 'NIHIL':
            blankfmt_exp['No'].append(data[i][0])
            blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
            blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
            blankfmt_exp['Bendera'].append(data[i][12])
            blankfmt_exp['Keagenan'].append(data[i][13])
            blankfmt_exp['GT'].append(data[i][5])
            blankfmt_exp['Tgl Tiba'].append(data[i][15])
            blankfmt_exp['Jam Tiba'].append(data[i][16])
            blankfmt_exp['Asal'].append(data[i][14])
            blankfmt_exp['Tgl Tambat'].append(data[i][15])
            blankfmt_exp['Jam Tambat'].append(data[i][16])
            blankfmt_exp['Tgl Tolak'].append(data[i][26])
            blankfmt_exp['Tujuan'].append(data[i][25])

        elif data[i][4][:data[i][4].find('. ')] in ['TB','TK','OB','BG']:
            blankfmt_exp['No'].append(data[i][0])
            blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
            blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
            blankfmt_exp['Bendera'].append(data[i][12])
            blankfmt_exp['Keagenan'].append(data[i][13])
            blankfmt_exp['GT'].append(data[i][5])
            blankfmt_exp['Tgl Tiba'].append(data[i][15])
            blankfmt_exp['Jam Tiba'].append(data[i][16])
            blankfmt_exp['Asal'].append(data[i][14])
            blankfmt_exp['Tgl Tambat'].append(data[i][15])
            blankfmt_exp['Jam Tambat'].append(data[i][16])
            blankfmt_exp['Tgl Tolak'].append(data[i][26])
            blankfmt_exp['Tujuan'].append(data[i][25])

    for i in range(len(blankfmt_exp['No'])):
        blankfmt_exp['No'][i] = i+1

    return blankfmt_exp

# for DOM format
def blankrows_dmstcs(datadom):
    data = [datadom.columns.values.tolist()] + datadom.values.tolist()

    blankfmt_dms = {'No':[],'Kode Kapal':[],'Nama Kapal':[],'Keagenan':[],'Bendera':[],'GT':[],'Trayek':[],'Tgl Tiba':[],
                    'Tgl Tolak':[],'Brg Bongkar D':[],'Jml Bongkar D':[],'1an Bongkar D':[],'Asal':[],'Brg Muat D':[],
                    'Jml Muat D':[],'1an Muat D':[],'Tujuan':[]}

    for i in range(1,len(data)):
        if ';' in data[i][18] and ';' in data[i][29]:
            arr_load = data[i][18].split('; ')
            arr_num,arr_mu = data[i][19].split('; '),data[i][20].split('; ')

            depar_load = data[i][29].split('; ')
            depar_num,depar_mu = data[i][30].split('; '),data[i][31].split('; ')
            
            if len(arr_load) == len(depar_load):
                for p in range(len(arr_load)):
                    if isinstance(arr_load[p],str) and p == 0:
                        blankfmt_dms['No'].append(data[i][0])
                        blankfmt_dms['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                        blankfmt_dms['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                        blankfmt_dms['Keagenan'].append(data[i][13])
                        blankfmt_dms['Bendera'].append(data[i][12])
                        blankfmt_dms['GT'].append(data[i][5])
                        blankfmt_dms['Trayek'].append('T')
                        blankfmt_dms['Tgl Tiba'].append(data[i][15])
                        blankfmt_dms['Tgl Tolak'].append(data[i][26])
                        blankfmt_dms['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dms['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dms['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dms['Asal'].append(data[i][14])
                        blankfmt_dms['Brg Muat D'].append(depar_load[p])
                        blankfmt_dms['Jml Muat D'].append(depar_num[p])
                        blankfmt_dms['1an Muat D'].append(depar_mu[p])
                        blankfmt_dms['Tujuan'].append(data[i][25])
                    else:
                        blankfmt_dms['No'].append(None)
                        blankfmt_dms['Kode Kapal'].append(None)
                        blankfmt_dms['Nama Kapal'].append(None)
                        blankfmt_dms['Keagenan'].append(None)
                        blankfmt_dms['Bendera'].append(None)
                        blankfmt_dms['GT'].append(None)
                        blankfmt_dms['Trayek'].append(None)
                        blankfmt_dms['Tgl Tiba'].append(None)
                        blankfmt_dms['Tgl Tolak'].append(None)
                        blankfmt_dms['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dms['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dms['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dms['Asal'].append(None)
                        blankfmt_dms['Brg Muat D'].append(depar_load[p])
                        blankfmt_dms['Jml Muat D'].append(depar_num[p])
                        blankfmt_dms['1an Muat D'].append(depar_mu[p])
                        blankfmt_dms['Tujuan'].append(None)

            elif len(arr_load) < len(depar_load):
                arr_load.extend(np.full([len(depar_load)-len(arr_load),1],None))
                arr_num.extend(np.full([len(depar_num)-len(arr_num),1],None))
                arr_mu.extend(np.full([len(depar_mu)-len(arr_mu),1],None))

                for p in range(len(depar_load)):
                    if isinstance(arr_load[p],str) and p == 0:
                        blankfmt_dms['No'].append(data[i][0])
                        blankfmt_dms['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                        blankfmt_dms['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                        blankfmt_dms['Keagenan'].append(data[i][13])
                        blankfmt_dms['Bendera'].append(data[i][12])
                        blankfmt_dms['GT'].append(data[i][5])
                        blankfmt_dms['Trayek'].append('T')
                        blankfmt_dms['Tgl Tiba'].append(data[i][15])
                        blankfmt_dms['Tgl Tolak'].append(data[i][26])
                        blankfmt_dms['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dms['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dms['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dms['Asal'].append(data[i][14])
                        blankfmt_dms['Brg Muat D'].append(depar_load[p])
                        blankfmt_dms['Jml Muat D'].append(depar_num[p])
                        blankfmt_dms['1an Muat D'].append(depar_mu[p])
                        blankfmt_dms['Tujuan'].append(data[i][25])
                    elif isinstance(arr_load[p],str) and p != 0:
                        blankfmt_dms['No'].append(None)
                        blankfmt_dms['Kode Kapal'].append(None)
                        blankfmt_dms['Nama Kapal'].append(None)
                        blankfmt_dms['Keagenan'].append(None)
                        blankfmt_dms['Bendera'].append(None)
                        blankfmt_dms['GT'].append(None)
                        blankfmt_dms['Trayek'].append(None)
                        blankfmt_dms['Tgl Tiba'].append(None)
                        blankfmt_dms['Tgl Tolak'].append(None)
                        blankfmt_dms['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dms['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dms['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dms['Asal'].append(None)
                        blankfmt_dms['Brg Muat D'].append(depar_load[p])
                        blankfmt_dms['Jml Muat D'].append(depar_num[p])
                        blankfmt_dms['1an Muat D'].append(depar_mu[p])
                        blankfmt_dms['Tujuan'].append(None)
                    elif not isinstance(arr_load[p],str) and p != 0:
                        blankfmt_dms['No'].append(None)
                        blankfmt_dms['Kode Kapal'].append(None)
                        blankfmt_dms['Nama Kapal'].append(None)
                        blankfmt_dms['Keagenan'].append(None)
                        blankfmt_dms['Bendera'].append(None)
                        blankfmt_dms['GT'].append(None)
                        blankfmt_dms['Trayek'].append(None)
                        blankfmt_dms['Tgl Tiba'].append(None)
                        blankfmt_dms['Tgl Tolak'].append(None)
                        blankfmt_dms['Brg Bongkar D'].append(None)
                        blankfmt_dms['Jml Bongkar D'].append(None)
                        blankfmt_dms['1an Bongkar D'].append(None)
                        blankfmt_dms['Asal'].append(None)
                        blankfmt_dms['Brg Muat D'].append(depar_load[p])
                        blankfmt_dms['Jml Muat D'].append(depar_num[p])
                        blankfmt_dms['1an Muat D'].append(depar_mu[p])
                        blankfmt_dms['Tujuan'].append(None)
            
            elif len(arr_load) > len(depar_load):
                depar_load.extend(np.full([1,len(arr_load)-len(depar_load)],None))
                depar_num.extend(np.full([1,len(arr_num)-len(depar_num)],None))
                depar_mu.extend(np.full([1,len(arr_mu)-len(depar_mu)],None))

                for p in range(len(arr_load)):
                    if isinstance(depar_load[p],str) and p == 0:
                        blankfmt_dms['No'].append(data[i][0])
                        blankfmt_dms['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                        blankfmt_dms['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                        blankfmt_dms['Keagenan'].append(data[i][13])
                        blankfmt_dms['Bendera'].append(data[i][12])
                        blankfmt_dms['GT'].append(data[i][5])
                        blankfmt_dms['Trayek'].append('T')
                        blankfmt_dms['Tgl Tiba'].append(data[i][15])
                        blankfmt_dms['Tgl Tolak'].append(data[i][26])
                        blankfmt_dms['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dms['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dms['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dms['Asal'].append(data[i][14])
                        blankfmt_dms['Brg Muat D'].append(depar_load[p])
                        blankfmt_dms['Jml Muat D'].append(depar_num[p])
                        blankfmt_dms['1an Muat D'].append(depar_mu[p])
                        blankfmt_dms['Tujuan'].append(data[i][25])
                    elif isinstance(depar_load[p],str) and p != 0:
                        blankfmt_dms['No'].append(None)
                        blankfmt_dms['Kode Kapal'].append(None)
                        blankfmt_dms['Nama Kapal'].append(None)
                        blankfmt_dms['Keagenan'].append(None)
                        blankfmt_dms['Bendera'].append(None)
                        blankfmt_dms['GT'].append(None)
                        blankfmt_dms['Trayek'].append(None)
                        blankfmt_dms['Tgl Tiba'].append(None)
                        blankfmt_dms['Tgl Tolak'].append(None)
                        blankfmt_dms['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dms['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dms['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dms['Asal'].append(None)
                        blankfmt_dms['Brg Muat D'].append(depar_load[p])
                        blankfmt_dms['Jml Muat D'].append(depar_num[p])
                        blankfmt_dms['1an Muat D'].append(depar_mu[p])
                        blankfmt_dms['Tujuan'].append(None)
                    elif not isinstance(depar_load[p],str) and p != 0:
                        blankfmt_dms['No'].append(None)
                        blankfmt_dms['Kode Kapal'].append(None)
                        blankfmt_dms['Nama Kapal'].append(None)
                        blankfmt_dms['Keagenan'].append(None)
                        blankfmt_dms['Bendera'].append(None)
                        blankfmt_dms['GT'].append(None)
                        blankfmt_dms['Trayek'].append(None)
                        blankfmt_dms['Tgl Tiba'].append(None)
                        blankfmt_dms['Tgl Tolak'].append(None)
                        blankfmt_dms['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_dms['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_dms['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_dms['Asal'].append(None)
                        blankfmt_dms['Brg Muat D'].append(None)
                        blankfmt_dms['Jml Muat D'].append(None)
                        blankfmt_dms['1an Muat D'].append(None)
                        blankfmt_dms['Tujuan'].append(None)
                        
        elif ';' in data[i][18] and ';' not in data[i][29]:
            arr_load = data[i][18].split('; ')
            arr_num = data[i][19].split('; ')
            arr_mu = data[i][20].split('; ')

            for p in range(len(arr_load)):
                if p == 0:
                    blankfmt_dms['No'].append(data[i][0])
                    blankfmt_dms['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                    blankfmt_dms['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                    blankfmt_dms['Keagenan'].append(data[i][13])
                    blankfmt_dms['Bendera'].append(data[i][12])
                    blankfmt_dms['GT'].append(data[i][5])
                    blankfmt_dms['Trayek'].append('T')
                    blankfmt_dms['Tgl Tiba'].append(data[i][15])
                    blankfmt_dms['Tgl Tolak'].append(data[i][26])
                    blankfmt_dms['Brg Bongkar D'].append(arr_load[p])
                    blankfmt_dms['Jml Bongkar D'].append(arr_num[p])
                    blankfmt_dms['1an Bongkar D'].append(arr_mu[p])
                    blankfmt_dms['Asal'].append(data[i][14])
                    blankfmt_dms['Brg Muat D'].append(data[i][29])
                    blankfmt_dms['Jml Muat D'].append(data[i][30])
                    blankfmt_dms['1an Muat D'].append(data[i][31])
                    blankfmt_dms['Tujuan'].append(data[i][25])
                elif p != 0:
                    blankfmt_dms['No'].append(None)
                    blankfmt_dms['Kode Kapal'].append(None)
                    blankfmt_dms['Nama Kapal'].append(None)
                    blankfmt_dms['Keagenan'].append(None)
                    blankfmt_dms['Bendera'].append(None)
                    blankfmt_dms['GT'].append(None)
                    blankfmt_dms['Trayek'].append(None)
                    blankfmt_dms['Tgl Tiba'].append(None)
                    blankfmt_dms['Tgl Tolak'].append(None)
                    blankfmt_dms['Brg Bongkar D'].append(arr_load[p])
                    blankfmt_dms['Jml Bongkar D'].append(arr_num[p])
                    blankfmt_dms['1an Bongkar D'].append(arr_mu[p])
                    blankfmt_dms['Asal'].append(None)
                    blankfmt_dms['Brg Muat D'].append(None)
                    blankfmt_dms['Jml Muat D'].append(None)
                    blankfmt_dms['1an Muat D'].append(None)
                    blankfmt_dms['Tujuan'].append(None)

        elif ';' not in data[i][18] and ';' in data[i][29]:
            depar_load = data[i][29].split('; ')
            depar_num = data[i][30].split('; ')
            depar_mu = data[i][31].split('; ')

            for p in range(len(depar_load)):
                if p == 0:
                    blankfmt_dms['No'].append(data[i][0])
                    blankfmt_dms['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                    blankfmt_dms['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                    blankfmt_dms['Keagenan'].append(data[i][13])
                    blankfmt_dms['Bendera'].append(data[i][12])
                    blankfmt_dms['GT'].append(data[i][5])
                    blankfmt_dms['Trayek'].append('T')
                    blankfmt_dms['Tgl Tiba'].append(data[i][15])
                    blankfmt_dms['Tgl Tolak'].append(data[i][26])
                    blankfmt_dms['Brg Bongkar D'].append(data[i][18])
                    blankfmt_dms['Jml Bongkar D'].append(data[i][19])
                    blankfmt_dms['1an Bongkar D'].append(data[i][20])
                    blankfmt_dms['Asal'].append(data[i][14])
                    blankfmt_dms['Brg Muat D'].append(depar_load[p])
                    blankfmt_dms['Jml Muat D'].append(depar_num[p])
                    blankfmt_dms['1an Muat D'].append(depar_mu[p])
                    blankfmt_dms['Tujuan'].append(data[i][25])
                elif p != 0:
                    blankfmt_dms['No'].append(None)
                    blankfmt_dms['Kode Kapal'].append(None)
                    blankfmt_dms['Nama Kapal'].append(None)
                    blankfmt_dms['Keagenan'].append(None)
                    blankfmt_dms['Bendera'].append(None)
                    blankfmt_dms['GT'].append(None)
                    blankfmt_dms['Trayek'].append(None)
                    blankfmt_dms['Tgl Tiba'].append(None)
                    blankfmt_dms['Tgl Tolak'].append(None)
                    blankfmt_dms['Brg Bongkar D'].append(None)
                    blankfmt_dms['Jml Bongkar D'].append(None)
                    blankfmt_dms['1an Bongkar D'].append(None)
                    blankfmt_dms['Asal'].append(None)
                    blankfmt_dms['Brg Muat D'].append(depar_load[p])
                    blankfmt_dms['Jml Muat D'].append(depar_num[p])
                    blankfmt_dms['1an Muat D'].append(depar_mu[p])
                    blankfmt_dms['Tujuan'].append(None)

        elif ';' not in data[i][18] and ';' not in data[i][29]:
            blankfmt_dms['No'].append(data[i][0])
            blankfmt_dms['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
            blankfmt_dms['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
            blankfmt_dms['Keagenan'].append(data[i][13])
            blankfmt_dms['Bendera'].append(data[i][12])
            blankfmt_dms['GT'].append(data[i][5])
            blankfmt_dms['Trayek'].append('T')
            blankfmt_dms['Tgl Tiba'].append(data[i][15])
            blankfmt_dms['Tgl Tolak'].append(data[i][26])
            blankfmt_dms['Brg Bongkar D'].append(data[i][18])
            blankfmt_dms['Jml Bongkar D'].append(data[i][19])
            blankfmt_dms['1an Bongkar D'].append(data[i][20])
            blankfmt_dms['Asal'].append(data[i][14])
            blankfmt_dms['Brg Muat D'].append(data[i][29])
            blankfmt_dms['Jml Muat D'].append(data[i][30])
            blankfmt_dms['1an Muat D'].append(data[i][31])
            blankfmt_dms['Tujuan'].append(data[i][25])

    return blankfmt_dms

# for EXP format
def blankrows_export(dataexp):
    data = [dataexp.columns.values.tolist()] + dataexp.values.tolist()

    blankfmt_exp = {'No':[],'Kode Kapal':[],'Nama Kapal':[],'Keagenan':[],'Bendera':[],'GT':[],'Trayek':[],'Tgl Tiba':[],
                    'Tgl Tolak':[],'Brg Bongkar D':[],'Jml Bongkar D':[],'1an Bongkar D':[],'Asal':[],'Brg Muat D':[],
                    'Jml Muat D':[],'1an Muat D':[],'Tujuan':[],'Shipper':[]}

    for i in range(1,len(data)):
        if ';' in data[i][18] and ';' in data[i][29]:
            arr_load = data[i][18].split('; ')
            arr_num,arr_mu = data[i][19].split('; '),data[i][20].split('; ')

            depar_load = data[i][29].split('; ')
            depar_num,depar_mu = data[i][30].split('; '),data[i][31].split('; ')
            
            if len(arr_load) == len(depar_load):
                for p in range(len(arr_load)):
                    if isinstance(arr_load[p],str) and p == 0:
                        blankfmt_exp['No'].append(data[i][0])
                        blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                        blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                        blankfmt_exp['Keagenan'].append(data[i][13])
                        blankfmt_exp['Bendera'].append(data[i][12])
                        blankfmt_exp['GT'].append(data[i][5])
                        blankfmt_exp['Trayek'].append('T')
                        blankfmt_exp['Tgl Tiba'].append(data[i][15])
                        blankfmt_exp['Tgl Tolak'].append(data[i][26])
                        blankfmt_exp['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_exp['Asal'].append(data[i][14])
                        blankfmt_exp['Brg Muat D'].append(depar_load[p])
                        blankfmt_exp['Jml Muat D'].append(depar_num[p])
                        blankfmt_exp['1an Muat D'].append(depar_mu[p])
                        blankfmt_exp['Tujuan'].append(data[i][25])
                        blankfmt_exp['Shipper'].append(data[i][36])
                    else:
                        blankfmt_exp['No'].append(None)
                        blankfmt_exp['Kode Kapal'].append(None)
                        blankfmt_exp['Nama Kapal'].append(None)
                        blankfmt_exp['Keagenan'].append(None)
                        blankfmt_exp['Bendera'].append(None)
                        blankfmt_exp['GT'].append(None)
                        blankfmt_exp['Trayek'].append(None)
                        blankfmt_exp['Tgl Tiba'].append(None)
                        blankfmt_exp['Tgl Tolak'].append(None)
                        blankfmt_exp['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_exp['Asal'].append(None)
                        blankfmt_exp['Brg Muat D'].append(depar_load[p])
                        blankfmt_exp['Jml Muat D'].append(depar_num[p])
                        blankfmt_exp['1an Muat D'].append(depar_mu[p])
                        blankfmt_exp['Tujuan'].append(None)
                        blankfmt_exp['Shipper'].append(None)

            elif len(arr_load) < len(depar_load):
                arr_load.extend(np.full([len(depar_load)-len(arr_load),1],None))
                arr_num.extend(np.full([len(depar_num)-len(arr_num),1],None))
                arr_mu.extend(np.full([len(depar_mu)-len(arr_mu),1],None))

                for p in range(len(depar_load)):
                    if isinstance(arr_load[p],str) and p == 0:
                        blankfmt_exp['No'].append(data[i][0])
                        blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                        blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                        blankfmt_exp['Keagenan'].append(data[i][13])
                        blankfmt_exp['Bendera'].append(data[i][12])
                        blankfmt_exp['GT'].append(data[i][5])
                        blankfmt_exp['Trayek'].append('T')
                        blankfmt_exp['Tgl Tiba'].append(data[i][15])
                        blankfmt_exp['Tgl Tolak'].append(data[i][26])
                        blankfmt_exp['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_exp['Asal'].append(data[i][14])
                        blankfmt_exp['Brg Muat D'].append(depar_load[p])
                        blankfmt_exp['Jml Muat D'].append(depar_num[p])
                        blankfmt_exp['1an Muat D'].append(depar_mu[p])
                        blankfmt_exp['Tujuan'].append(data[i][25])
                        blankfmt_exp['Shipper'].append(data[i][36])
                    elif isinstance(arr_load[p],str) and p != 0:
                        blankfmt_exp['No'].append(None)
                        blankfmt_exp['Kode Kapal'].append(None)
                        blankfmt_exp['Nama Kapal'].append(None)
                        blankfmt_exp['Keagenan'].append(None)
                        blankfmt_exp['Bendera'].append(None)
                        blankfmt_exp['GT'].append(None)
                        blankfmt_exp['Trayek'].append(None)
                        blankfmt_exp['Tgl Tiba'].append(None)
                        blankfmt_exp['Tgl Tolak'].append(None)
                        blankfmt_exp['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_exp['Asal'].append(None)
                        blankfmt_exp['Brg Muat D'].append(depar_load[p])
                        blankfmt_exp['Jml Muat D'].append(depar_num[p])
                        blankfmt_exp['1an Muat D'].append(depar_mu[p])
                        blankfmt_exp['Tujuan'].append(None)
                        blankfmt_exp['Shipper'].append(None)
                    elif not isinstance(arr_load[p],str) and p != 0:
                        blankfmt_exp['No'].append(None)
                        blankfmt_exp['Kode Kapal'].append(None)
                        blankfmt_exp['Nama Kapal'].append(None)
                        blankfmt_exp['Keagenan'].append(None)
                        blankfmt_exp['Bendera'].append(None)
                        blankfmt_exp['GT'].append(None)
                        blankfmt_exp['Trayek'].append(None)
                        blankfmt_exp['Tgl Tiba'].append(None)
                        blankfmt_exp['Tgl Tolak'].append(None)
                        blankfmt_exp['Brg Bongkar D'].append(None)
                        blankfmt_exp['Jml Bongkar D'].append(None)
                        blankfmt_exp['1an Bongkar D'].append(None)
                        blankfmt_exp['Asal'].append(None)
                        blankfmt_exp['Brg Muat D'].append(depar_load[p])
                        blankfmt_exp['Jml Muat D'].append(depar_num[p])
                        blankfmt_exp['1an Muat D'].append(depar_mu[p])
                        blankfmt_exp['Tujuan'].append(None)
                        blankfmt_exp['Shipper'].append(None)
            
            elif len(arr_load) > len(depar_load):
                depar_load.extend(np.full([1,len(arr_load)-len(depar_load)],None))
                depar_num.extend(np.full([1,len(arr_num)-len(depar_num)],None))
                depar_mu.extend(np.full([1,len(arr_mu)-len(depar_mu)],None))

                for p in range(len(arr_load)):
                    if isinstance(depar_load[p],str) and p == 0:
                        blankfmt_exp['No'].append(data[i][0])
                        blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                        blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                        blankfmt_exp['Keagenan'].append(data[i][13])
                        blankfmt_exp['Bendera'].append(data[i][12])
                        blankfmt_exp['GT'].append(data[i][5])
                        blankfmt_exp['Trayek'].append('T')
                        blankfmt_exp['Tgl Tiba'].append(data[i][15])
                        blankfmt_exp['Tgl Tolak'].append(data[i][26])
                        blankfmt_exp['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_exp['Asal'].append(data[i][14])
                        blankfmt_exp['Brg Muat D'].append(depar_load[p])
                        blankfmt_exp['Jml Muat D'].append(depar_num[p])
                        blankfmt_exp['1an Muat D'].append(depar_mu[p])
                        blankfmt_exp['Tujuan'].append(data[i][25])
                        blankfmt_exp['Shipper'].append(data[i][36])
                    elif isinstance(depar_load[p],str) and p != 0:
                        blankfmt_exp['No'].append(None)
                        blankfmt_exp['Kode Kapal'].append(None)
                        blankfmt_exp['Nama Kapal'].append(None)
                        blankfmt_exp['Keagenan'].append(None)
                        blankfmt_exp['Bendera'].append(None)
                        blankfmt_exp['GT'].append(None)
                        blankfmt_exp['Trayek'].append(None)
                        blankfmt_exp['Tgl Tiba'].append(None)
                        blankfmt_exp['Tgl Tolak'].append(None)
                        blankfmt_exp['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_exp['Asal'].append(None)
                        blankfmt_exp['Brg Muat D'].append(depar_load[p])
                        blankfmt_exp['Jml Muat D'].append(depar_num[p])
                        blankfmt_exp['1an Muat D'].append(depar_mu[p])
                        blankfmt_exp['Tujuan'].append(None)
                        blankfmt_exp['Shipper'].append(None)
                    elif not isinstance(depar_load[p],str) and p != 0:
                        blankfmt_exp['No'].append(None)
                        blankfmt_exp['Kode Kapal'].append(None)
                        blankfmt_exp['Nama Kapal'].append(None)
                        blankfmt_exp['Keagenan'].append(None)
                        blankfmt_exp['Bendera'].append(None)
                        blankfmt_exp['GT'].append(None)
                        blankfmt_exp['Trayek'].append(None)
                        blankfmt_exp['Tgl Tiba'].append(None)
                        blankfmt_exp['Tgl Tolak'].append(None)
                        blankfmt_exp['Brg Bongkar D'].append(arr_load[p])
                        blankfmt_exp['Jml Bongkar D'].append(arr_num[p])
                        blankfmt_exp['1an Bongkar D'].append(arr_mu[p])
                        blankfmt_exp['Asal'].append(None)
                        blankfmt_exp['Brg Muat D'].append(None)
                        blankfmt_exp['Jml Muat D'].append(None)
                        blankfmt_exp['1an Muat D'].append(None)
                        blankfmt_exp['Tujuan'].append(None)
                        blankfmt_exp['Shipper'].append(None)
                        
        elif ';' in data[i][18] and ';' not in data[i][29]:
            arr_load = data[i][18].split('; ')
            arr_num = data[i][19].split('; ')
            arr_mu = data[i][20].split('; ')

            for p in range(len(arr_load)):
                if p == 0:
                    blankfmt_exp['No'].append(data[i][0])
                    blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                    blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                    blankfmt_exp['Keagenan'].append(data[i][13])
                    blankfmt_exp['Bendera'].append(data[i][12])
                    blankfmt_exp['GT'].append(data[i][5])
                    blankfmt_exp['Trayek'].append('T')
                    blankfmt_exp['Tgl Tiba'].append(data[i][15])
                    blankfmt_exp['Tgl Tolak'].append(data[i][26])
                    blankfmt_exp['Brg Bongkar D'].append(arr_load[p])
                    blankfmt_exp['Jml Bongkar D'].append(arr_num[p])
                    blankfmt_exp['1an Bongkar D'].append(arr_mu[p])
                    blankfmt_exp['Asal'].append(data[i][14])
                    blankfmt_exp['Brg Muat D'].append(data[i][29])
                    blankfmt_exp['Jml Muat D'].append(data[i][30])
                    blankfmt_exp['1an Muat D'].append(data[i][31])
                    blankfmt_exp['Tujuan'].append(data[i][25])
                    blankfmt_exp['Shipper'].append(data[i][36])
                elif p != 0:
                    blankfmt_exp['No'].append(None)
                    blankfmt_exp['Kode Kapal'].append(None)
                    blankfmt_exp['Nama Kapal'].append(None)
                    blankfmt_exp['Keagenan'].append(None)
                    blankfmt_exp['Bendera'].append(None)
                    blankfmt_exp['GT'].append(None)
                    blankfmt_exp['Trayek'].append(None)
                    blankfmt_exp['Tgl Tiba'].append(None)
                    blankfmt_exp['Tgl Tolak'].append(None)
                    blankfmt_exp['Brg Bongkar D'].append(arr_load[p])
                    blankfmt_exp['Jml Bongkar D'].append(arr_num[p])
                    blankfmt_exp['1an Bongkar D'].append(arr_mu[p])
                    blankfmt_exp['Asal'].append(None)
                    blankfmt_exp['Brg Muat D'].append(None)
                    blankfmt_exp['Jml Muat D'].append(None)
                    blankfmt_exp['1an Muat D'].append(None)
                    blankfmt_exp['Tujuan'].append(None)
                    blankfmt_exp['Shipper'].append(None)

        elif ';' not in data[i][18] and ';' in data[i][29]:
            depar_load = data[i][29].split('; ')
            depar_num = data[i][30].split('; ')
            depar_mu = data[i][31].split('; ')

            for p in range(len(depar_load)):
                if p == 0:
                    blankfmt_exp['No'].append(data[i][0])
                    blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                    blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                    blankfmt_exp['Keagenan'].append(data[i][13])
                    blankfmt_exp['Bendera'].append(data[i][12])
                    blankfmt_exp['GT'].append(data[i][5])
                    blankfmt_exp['Trayek'].append('T')
                    blankfmt_exp['Tgl Tiba'].append(data[i][15])
                    blankfmt_exp['Tgl Tolak'].append(data[i][26])
                    blankfmt_exp['Brg Bongkar D'].append(data[i][18])
                    blankfmt_exp['Jml Bongkar D'].append(data[i][19])
                    blankfmt_exp['1an Bongkar D'].append(data[i][20])
                    blankfmt_exp['Asal'].append(data[i][14])
                    blankfmt_exp['Brg Muat D'].append(depar_load[p])
                    blankfmt_exp['Jml Muat D'].append(depar_num[p])
                    blankfmt_exp['1an Muat D'].append(depar_mu[p])
                    blankfmt_exp['Tujuan'].append(data[i][25])
                    blankfmt_exp['Shipper'].append(data[i][36])
                elif p != 0:
                    blankfmt_exp['No'].append(None)
                    blankfmt_exp['Kode Kapal'].append(None)
                    blankfmt_exp['Nama Kapal'].append(None)
                    blankfmt_exp['Keagenan'].append(None)
                    blankfmt_exp['Bendera'].append(None)
                    blankfmt_exp['GT'].append(None)
                    blankfmt_exp['Trayek'].append(None)
                    blankfmt_exp['Tgl Tiba'].append(None)
                    blankfmt_exp['Tgl Tolak'].append(None)
                    blankfmt_exp['Brg Bongkar D'].append(None)
                    blankfmt_exp['Jml Bongkar D'].append(None)
                    blankfmt_exp['1an Bongkar D'].append(None)
                    blankfmt_exp['Asal'].append(None)
                    blankfmt_exp['Brg Muat D'].append(depar_load[p])
                    blankfmt_exp['Jml Muat D'].append(depar_num[p])
                    blankfmt_exp['1an Muat D'].append(depar_mu[p])
                    blankfmt_exp['Tujuan'].append(None)
                    blankfmt_exp['Shipper'].append(None)

        elif ';' not in data[i][18] and ';' not in data[i][29]:
            blankfmt_exp['No'].append(data[i][0])
            blankfmt_exp['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
            blankfmt_exp['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
            blankfmt_exp['Keagenan'].append(data[i][13])
            blankfmt_exp['Bendera'].append(data[i][12])
            blankfmt_exp['GT'].append(data[i][5])
            blankfmt_exp['Trayek'].append('T')
            blankfmt_exp['Tgl Tiba'].append(data[i][15])
            blankfmt_exp['Tgl Tolak'].append(data[i][26])
            blankfmt_exp['Brg Bongkar D'].append(data[i][18])
            blankfmt_exp['Jml Bongkar D'].append(data[i][19])
            blankfmt_exp['1an Bongkar D'].append(data[i][20])
            blankfmt_exp['Asal'].append(data[i][14])
            blankfmt_exp['Brg Muat D'].append(data[i][29])
            blankfmt_exp['Jml Muat D'].append(data[i][30])
            blankfmt_exp['1an Muat D'].append(data[i][31])
            blankfmt_exp['Tujuan'].append(data[i][25])
            blankfmt_exp['Shipper'].append(data[i][36])

    return blankfmt_exp

# For CLRNC Format
def blankrows_clrnc(dataclr):
    data = [dataclr.columns.values.tolist()] + dataclr.values.tolist()

    blankfmt_clr = {'No':[],'Kode SPB':[],'No Seri':[],'No Reg':[],'Kode Kapal':[],'Nama Kapal':[],'Nahkoda':[],
                    'Bendera':[],'GT':[],'SIPI':[],'SIKPI':[],'SLO':[],'Asal':[],'Tgl Tiba':[],'Jml Kru':[],
                    'Tujuan':[],'Tgl Tolak':[],'Brg Muat':[],'Jml Muat':[],'1an Muat':[],'Keagenan':[]}

    for i in range(1,len(data)):
        if ';' in data[i][29]:
            depar_load = data[i][29].split('; ')
            depar_num = data[i][30].split('; ')
            depar_mu = data[i][31].split('; ')

            for p in range(len(depar_load)):
                if p == 0:
                    blankfmt_clr['No'].append(data[i][0])
                    blankfmt_clr['Kode SPB'].append('T58')
                    blankfmt_clr['No Seri'].append(data[i][1])
                    if data[i][2] != '--' and data[i][3] == '--':
                        blankfmt_clr['No Reg'].append(data[i][2])
                    elif data[i][2] == '--' and data[i][3] != '--':
                        blankfmt_clr['No Reg'].append(data[i][3])
                    blankfmt_clr['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
                    blankfmt_clr['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
                    blankfmt_clr['Nahkoda'].append(data[i][9])
                    blankfmt_clr['Bendera'].append(data[i][12])
                    blankfmt_clr['GT'].append(data[i][5])
                    blankfmt_clr['SIPI'].append('--')
                    blankfmt_clr['SIKPI'].append('--')
                    blankfmt_clr['SLO'].append('--')
                    blankfmt_clr['Asal'].append(data[i][14])
                    blankfmt_clr['Tgl Tiba'].append(data[i][15])
                    blankfmt_clr['Jml Kru'].append(data[i][11])
                    blankfmt_clr['Tujuan'].append(data[i][25])
                    blankfmt_clr['Tgl Tolak'].append(data[i][26])
                    blankfmt_clr['Brg Muat'].append(depar_load[p])
                    blankfmt_clr['Jml Muat'].append(depar_num[p])
                    blankfmt_clr['1an Muat'].append(depar_mu[p])
                    blankfmt_clr['Keagenan'].append(data[i][13])
                elif p != 0:
                    blankfmt_clr['No'].append(None)
                    blankfmt_clr['Kode SPB'].append(None)
                    blankfmt_clr['No Seri'].append(None)
                    blankfmt_clr['No Reg'].append(None)
                    blankfmt_clr['Kode Kapal'].append(None)
                    blankfmt_clr['Nama Kapal'].append(None)
                    blankfmt_clr['Nahkoda'].append(None)
                    blankfmt_clr['Bendera'].append(None)
                    blankfmt_clr['GT'].append(None)
                    blankfmt_clr['SIPI'].append(None)
                    blankfmt_clr['SIKPI'].append(None)
                    blankfmt_clr['SLO'].append(None)
                    blankfmt_clr['Asal'].append(None)
                    blankfmt_clr['Tgl Tiba'].append(None)
                    blankfmt_clr['Jml Kru'].append(None)
                    blankfmt_clr['Tujuan'].append(None)
                    blankfmt_clr['Tgl Tolak'].append(None)
                    blankfmt_clr['Brg Muat'].append(depar_load[p])
                    blankfmt_clr['Jml Muat'].append(depar_num[p])
                    blankfmt_clr['1an Muat'].append(depar_mu[p])
                    blankfmt_clr['Keagenan'].append(None)

        elif ';' not in data[i][29]:
            blankfmt_clr['No'].append(data[i][0])
            blankfmt_clr['Kode SPB'].append('T58')
            blankfmt_clr['No Seri'].append(data[i][1])
            if data[i][2] != '--' and data[i][3] == '--':
                blankfmt_clr['No Reg'].append(data[i][2])
            elif data[i][2] == '--' and data[i][3] != '--':
                blankfmt_clr['No Reg'].append(data[i][3])
            blankfmt_clr['Kode Kapal'].append(data[i][4][:data[i][4].find('. ')])
            blankfmt_clr['Nama Kapal'].append(data[i][4][data[i][4].find('. ')+2:])
            blankfmt_clr['Nahkoda'].append(data[i][9])
            blankfmt_clr['Bendera'].append(data[i][12])
            blankfmt_clr['GT'].append(data[i][5])
            blankfmt_clr['SIPI'].append('--')
            blankfmt_clr['SIKPI'].append('--')
            blankfmt_clr['SLO'].append('--')
            blankfmt_clr['Asal'].append(data[i][14])
            blankfmt_clr['Tgl Tiba'].append(data[i][15])
            blankfmt_clr['Jml Kru'].append(data[i][11])
            blankfmt_clr['Tujuan'].append(data[i][25])
            blankfmt_clr['Tgl Tolak'].append(data[i][26])
            blankfmt_clr['Brg Muat'].append(data[i][29])
            blankfmt_clr['Jml Muat'].append(data[i][30])
            blankfmt_clr['1an Muat'].append(data[i][31])
            blankfmt_clr['Keagenan'].append(data[i][13])

    return blankfmt_clr

# core functions, for monthly report

# 1 SIB report functions
# Handler
def sib_based(dfsibk,dfsibg,xlwriter):
    #declare dict for writing the output
    dict_sibkecil = {'Nomor':[],'Nama Kapal':[],'Bendera':[],'Nama Nakhoda':[],
                    'GT':[],'NT':[],'Tanda Selar':[],'Tempat Pendaftaran':[],'Tanggal Tiba':[],
                    'Asal Kapal':[],'Kode Muatan Tiba':[],'Tanggal Tolak':[],'Tujuan Kapal':[],
                    'Kode Muatan Tolak':[],'Keagenan':[],'Keterangan':[]
                    }

    dict_sibgede = {'Nomor':[],'Nama Kapal':[],'Bendera':[],'Nama Nakhoda':[],'Tempat Pendaftaran':[],
                    'GT':[],'NT':[],'Tanggal Tiba':[],'Asal Kapal':[],'Kode Muatan Tiba':[],'Tanggal Tolak':[],
                    'Tujuan Kapal':[],'Kode Muatan Tolak':[],'Keagenan':[],'Keterangan':[]
                    }

    #recreate list for data assignment
    sibkassg = ['NO','NAMA KAPAL','BENDERA','NAKHODA','GT','NT','TANDA SELAR',
                'TEMPAT PENDAFTARAN','TANGGAL TIBA','TIBA DARI','TIBA ISI / KOSONG',
                'TANGGAL BERTOLAK','TUJUAN','ISI / KOSONG',
                'PEMILIK / AGEN','KET']

    sibgassg = ['NO','NAMA KAPAL','BENDERA','NAKHODA','TEMPAT PENDAFTARAN','GT',
                'NT','TANGGAL TIBA','TIBA DARI','TIBA ISI / KOSONG','TANGGAL BERTOLAK',
                'TUJUAN','ISI / KOSONG','PEMILIK / AGEN','KET']

    #assigning dataframe's values into declared dictionary
    ini = 0
    for keys in dict_sibkecil:
        dict_sibkecil[keys] = dfsibk[sibkassg[ini]].values.tolist()
        ini += 1

    ini = 0
    for keys in dict_sibgede:
        dict_sibgede[keys] = dfsibg[sibgassg[ini]].values.tolist()
        ini += 1

    writesib(xlwriter,[dict_sibkecil,dict_sibgede])

# writer
def writesib(xlwriter,listofdf):
    xlworkbk = xlwriter.book

    #WORKSHEET FORMATS

    #title
    fmt_title_nonbold = xlworkbk.add_format({'font_name':'arial','font_size':12,'align':'center','valign':'vcenter'})
    fmt_title_bold = xlworkbk.add_format({'bold':True,'font_name':'arial','font_size':12,'align':'center','valign':'vcenter'})
    fmt_title_bold_underline = xlworkbk.add_format({'bold':True,'underline':True,'font_name':'arial','font_size':12,'align':'center','valign':'vcenter'})

    #header formats
    fmt_mheader = xlworkbk.add_format({'font_name':'arial','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':1,'align':'center','valign':'vcenter'})
    fmt_lheader = xlworkbk.add_format({'font_name':'arial','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rheader = xlworkbk.add_format({'font_name':'arial','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_header_subs = xlworkbk.add_format({'font_name':'arial','font_size':12,'text_wrap':True,'bottom':6,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter'})
    fmt_header_bold = xlworkbk.add_format({'font_name':'arial','font_size':12,'bold':True,'bottom':1,'top':6,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #filler formats
    fmt_lfiller = xlworkbk.add_format({'font_name':'arial','font_size':12,'text_wrap':True,'bottom':1,'top':6,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rfiller = xlworkbk.add_format({'font_name':'arial','font_size':12,'text_wrap':True,'bottom':1,'top':6,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_mfiller = xlworkbk.add_format({'font_name':'arial','font_size':12,'text_wrap':True,'bottom':1,'top':6,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #main data formats
    #upper part
    fmt_lumain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rumain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_mumain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #middle part
    fmt_lmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_mmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #downer part
    fmt_ldmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rdmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_mdmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #name, tonnage measurement number, agent's name
    fmt_ulefts = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'left','valign':'vcenter'})
    fmt_mlefts = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'left','valign':'vcenter'})
    fmt_dlefts = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'left','valign':'vcenter'})

    #gross and net tonnage
    fmt_urights = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'right','valign':'vcenter'})
    fmt_mrights = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'right','valign':'vcenter'})
    fmt_drights = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'right','valign':'vcenter'})

    #dates
    fmt_udates = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'dd/mm/yyyy'})
    fmt_mdates = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'dd/mm/yyyy'})
    fmt_ddates = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'dd/mm/yyyy'})

    #footer formats
    fmt_row1 = xlworkbk.add_format({'font_name':'arial','font_size':11,'text_wrap':True,'bold':True,'bottom':None,'top':2,'right':2,'left':2,'align':'center','valign':'vcenter'})
    fmt_row2 = xlworkbk.add_format({'bottom':None,'top':None,'right':2,'left':2})
    fmt_row3 = xlworkbk.add_format({'bottom':2,'top':None,'right':2,'left':2})

    #signing format
    fmt_signbold = xlworkbk.add_format({'font_name':'arial','font_size':16,'bold':True})
    fmt_signdflt = xlworkbk.add_format({'font_name':'arial','font_size':16,'bold':False})

    #REQUIRED VARIABLES

    #worksheet's name
    wsname = ['SIB KECIL PPK 29','SIB BESAR PPK 27']

    #month and year
    prevmonth = str(trnslt.translate(dt.strftime(dt.today().replace(day=1) - td(days=1),'%B')).upper())
    prevyear = dt.strftime(dt.today().replace(day=1) - td(days=1),'%Y')

    #list of header's rowcol, contents, and format sequences
    rcwmsibk = ['A5:A6','B5:D5','E5:E6','F5:F6','G5:G6','H5:H6','I5:K5','L5:N5','O5:O6','P5:P6']
    rcnmsibk = ['B6','C6','D6','I6','J6','K6','L6','M6','N6']
    cowmsibk = ['No.','N A M A','Berat Kotor (GT)','Berat Bersih (NT)','Tanda Selar Menurut Pas Tahunan',
                'Tempat kedudukan kapal','T I B A','B E R A N G K A T','Diageni/charter/kapal milik, disebutkan juga','KET.']
    conmsibk = ['Kapal','Bendera','Nakhoda','Tanggal','Tempat terakhir yang disinggahi','Bermuatan/kosong','Tanggal',
                'Tempat pertama yang disinggahi','Bermuatan/kosong']
    fmtmsibk = [fmt_lheader,fmt_header_bold,fmt_mheader,fmt_mheader,fmt_mheader,
                fmt_mheader,fmt_header_bold,fmt_header_bold,fmt_mheader,fmt_rheader]

    rcwmsibg = ['A5:A6','B5:D5','E5:E6','F5:G5','H5:J5','K5:M5','N5:N6','O5:O6']
    rcnmsibg = ['B6','C6','D6','F6','G6','H6','I6','J6','K6','L6','M6']
    cowmsibg = ['No.','N A M A','Tempat Kedudukan Kapal','T O N A S E','T I B A','B E R A N G K A T',
                'Diageni/charter/kapal milik, disebutkan juga','KET.']
    conmsibg = ['Kapal','Bendera','Nakhoda','Berat Kotor (GT)','Berat Bersih (NT)','Tanggal','Tempat terakhir yang disinggahi',
                'Bermuatan/kosong','Tanggal','Tempat tujuan terakhir','Bermuatan/kosong']
    fmtmsibg = [fmt_lheader,fmt_mheader,fmt_mheader,fmt_mheader,fmt_mheader,fmt_mheader,fmt_mheader,fmt_rheader]

    rcwm,rcnm,cowm,conm,fmtm = [rcwmsibk,rcwmsibg], [rcnmsibk,rcnmsibg], [cowmsibk,cowmsibg], [conmsibk,conmsibg], [fmtmsibk,fmtmsibg]

    #filler's length and column's number for signing place
    fillerlen,cnfsp = [16,15],[13,12]

    #inside footer's box
    footsibk = 'DKP.V - 10B = PPK - 29'
    footsibg = 'DKP.V - 10B = PPK - 27'
    foot = [footsibk,footsibg]

    #list of column's width and row's height
    cowisibk = [44,265,135,287,71,71,91,147,102,169,84,98,187,84,358,82]
    cowisibg = [35,265,120,287,144,78,78,127,162,89,113,171,95,388,78]
    rohesib = [22,22,22,22,35,100,22,35]

    cowi = [cowisibk,cowisibg]

    #EXECUTE WRITING SESSION
    
    for wr in range(2):

        #naming worksheet
        worksh = xlworkbk.add_worksheet(wsname[wr])

        #writing universal titles
        rowcols = ['A1:D1','A2:D2','A3:D3','F1:M1','O3:P3']
        contents = ['KEMENTERIAN PERHUBUNGAN','DIREKTORAT JENDERAL PERHUBUNGAN LAUT','UNIT PENYELENGGARA PELABUHAN KELAS III PULAU BUNYU',
                    'DAFTAR KAPAL YANG KELUAR MASUK DI PELABUHAN PULAU BUNYU','SELAMA BULAN:'+' '+prevmonth+' '+prevyear]
        fmts = [fmt_title_nonbold,fmt_title_nonbold,fmt_title_bold_underline,fmt_title_bold,fmt_title_bold]
        for val in range(len(rowcols)):
            if wr == 1 and val == len(rowcols)-1:
                rowcols[val] = 'N3:O3'
            worksh.merge_range(rowcols[val],contents[val],fmts[val])

        #writing table's header
        rowcols, contents, fmts = rcwm[wr], cowm[wr], fmtm[wr]
        for val in range(len(rowcols)):
            worksh.merge_range(rowcols[val],contents[val],fmts[val])

        rowcols, contents = rcnm[wr], conm[wr]
        for val in range(len(rowcols)):
            worksh.write(rowcols[val],contents[val],fmt_header_subs)

        #write filler
        for i in range(fillerlen[wr]):
            if i == 0:
                worksh.write(6,0,'('+str(i+1)+')',fmt_lfiller)
            elif i == fillerlen[wr]-1:
                worksh.write(6,fillerlen[wr]-1,'('+str(i+1)+')',fmt_rfiller)
            elif i != 0 and i!=fillerlen[wr]-1:
                worksh.write(6,i,'('+str(i+1)+')',fmt_mfiller)
        
        #write main data
        ini,dfsib = 0,listofdf[wr]
        for keys in dfsib:
            for i in range(len(dfsib[keys])):
                if keys == 'Nomor':
                    if i == 0:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_lumain)
                    elif i == len(dfsib[keys])-1:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_ldmain)
                    elif i != 0 and i != len(dfsib[keys])-1:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_lmain)
                
                elif keys == 'Keterangan':
                    if i == 0:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_rumain)
                    elif i == len(dfsib[keys])-1:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_rdmain)
                    elif i != 0 and i != len(dfsib[keys])-1:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_rmain)

                elif keys in ['Tanggal Tiba','Tanggal Tolak']:
                    if i == 0:
                        worksh.write(i+7,ini,dt.fromtimestamp((dfsib[keys][i]/1e6) * 0.001),fmt_udates)
                    elif i == len(dfsib[keys])-1:
                        worksh.write(i+7,ini,dt.fromtimestamp((dfsib[keys][i]/1e6) * 0.001),fmt_ddates)
                    elif i != 0 and i != len(dfsib[keys])-1:
                        worksh.write(i+7,ini,dt.fromtimestamp((dfsib[keys][i]/1e6) * 0.001),fmt_mdates)

                elif keys in ['Nama Kapal','Tanda Selar','Keagenan']:
                    if i == 0:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_ulefts)
                    elif i == len(dfsib[keys])-1:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_dlefts)
                    elif i != 0 and i != len(dfsib[keys])-1:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_mlefts)

                elif keys in ['GT','NT']:
                    if i == 0:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_urights)
                    elif i == len(dfsib[keys])-1:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_drights)
                    elif i != 0 and i != len(dfsib[keys])-1:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_mrights)
                
                else:
                    if i == 0:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_mumain)
                    elif i == len(dfsib[keys])-1:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_mdmain)
                    elif i != 0 and i != len(dfsib[keys])-1:
                        worksh.write(i+7,ini,dfsib[keys][i],fmt_mmain)
            ini += 1

        worksh.write(7+len(dfsib[keys])+4,1,foot[wr],fmt_row1)
        worksh.write(7+len(dfsib[keys])+5,1,' ',fmt_row2)
        worksh.write(7+len(dfsib[keys])+6,1,' ',fmt_row3)

        #set cell's width and height

        colwid = cowi[wr]
        for num in range(len(colwid)):
            worksh.set_column_pixels(num,num,colwid[num])

        for num in range(len(rohesib)):
            if num == len(rohesib)-1:
                for i in range(len(dfsib[keys])):
                    worksh.set_row_pixels(i+7,rohesib[num])
            elif num != len(rohesib)-1:
                worksh.set_row_pixels(num,rohesib[num])

        #write signing place
        datetext = 'Pulau Bunyu, '+str(dt.strftime(dt.today(),'%d'))+' '+str(trnslt.translate(dt.strftime(dt.today(),'%B')))+' '+str(dt.strftime(dt.today(),'%Y'))
        signtext = [None,'Kepala Kantor','Unit Penyelenggara Pelabuhan','Kelas III Pulau Bunyu',
                    None,None,None,None,'Abdul Wahid','NIP. 19710515 199803 1 006']

        worksh.write(7+len(dfsib[keys])+2,cnfsp[wr],datetext,fmt_signdflt)
        worksh.write_column(7+len(dfsib[keys])+3,cnfsp[wr],signtext,fmt_signbold)

#2 TKII style report
# Handler
def tkii_based(tkdata,place,xlwriter):
    if place == 'tkiiupt':
        tkdata = tkdata.loc[(tkdata['LOKASI BONGKAR'].str.contains('TK II UPT',case=False)) | (tkdata['LOKASI MUAT'].str.contains('TK II UPT',case=False))]
        tkdata['NO'] = range(1,len(tkdata)+1)
        tkdata = tkdata.reset_index(drop=True)
        dicttkdata = blankrows_tkii01(tkdata)
        writing_tkii('dtkii',xlwriter,dicttkdata,'NON TERMINAL')

    if place == 'prtmn':
        tkdata = tkdata.loc[(tkdata['LOKASI BONGKAR'].str.contains('PERTAMINA BUNYU',case=False)) | (tkdata['LOKASI MUAT'].str.contains('PERTAMINA BUNYU',case=False))]
        tkdata['NO'] = range(1,len(tkdata)+1)
        tkdata = tkdata.reset_index(drop=True)
        dicttkdata = blankrows_tkii01(tkdata)
        writing_tkii('dtkii',xlwriter,dicttkdata,'TUKS PT. PERTAMINA RU V BUNYU')
    
    elif place == 'mipsk':
        tkdata = tkdata.loc[(tkdata['LOKASI BONGKAR'].str.contains('MIP SEI KRASSI',case=False)) | (tkdata['LOKASI MUAT'].str.contains('MIP SEI KRASSI',case=False))]
        tkdata['NO'] = range(1,len(tkdata)+1)
        tkdata = tkdata.reset_index(drop=True)
        dicttkdata = blankrows_tkii01(tkdata)
        writing_tkii('dtkii',xlwriter,dicttkdata,'TERSUS PT. MIP SEI KRASSI')

    elif place == 'mipmj':
        tkdata = tkdata.loc[(tkdata['LOKASI BONGKAR'].str.contains('MIP MANJELUTUNG',case=False)) | (tkdata['LOKASI MUAT'].str.contains('MIP MANJELUTUNG',case=False))]
        tkdata['NO'] = range(1,len(tkdata)+1)
        tkdata = tkdata.reset_index(drop=True)
        dicttkdata = blankrows_tkii01(tkdata)
        writing_tkii('dtkii',xlwriter,dicttkdata,'TERSUS PT. MIP MANJELUTUNG')
        
    elif place == 'skkms':
        tkdata = tkdata.loc[(tkdata['LOKASI BONGKAR'].str.contains('SKK MIGAS SEMBAKUNG',case=False)) | (tkdata['LOKASI MUAT'].str.contains('SKK MIGAS SEMBAKUNG',case=False))]
        tkdata['NO'] = range(1,len(tkdata)+1)
        tkdata = tkdata.reset_index(drop=True)
        dicttkdata = blankrows_tkii01(tkdata)
        writing_tkii('dtkii',xlwriter,dicttkdata,'TERSUS SKK MIGAS SEMBAKUNG')

    elif place == 'pttum':
        tkdata = tkdata.loc[(tkdata['LOKASI BONGKAR'].str.contains('TUM SESAYAP',case=False)) | (tkdata['LOKASI MUAT'].str.contains('TUM SESAYAP',case=False))]
        tkdata['NO'] = range(1,len(tkdata)+1)
        tkdata = tkdata.reset_index(drop=True)
        dicttkdata = blankrows_tkii01(tkdata)
        writing_tkii('dtkii',xlwriter,dicttkdata,'TERSUS PT. TEKNIK UTAMA MANDIRI')

    elif place == 'ptssp':
        tkdata = tkdata.loc[(tkdata['LOKASI BONGKAR'].str.contains('SEBAUNG SAWIT PLANTATIONS',case=False)) | (tkdata['LOKASI MUAT'].str.contains('SEBAUNG SAWIT PLANTATIONS',case=False))]
        tkdata['NO'] = range(1,len(tkdata)+1)
        tkdata = tkdata.reset_index(drop=True)
        dicttkdata = blankrows_tkii01(tkdata)
        writing_tkii('dtkii',xlwriter,dicttkdata,'TERSUS PT. SSP')

    elif place == 'klngn':
        tkdata = tkdata.loc[(tkdata['LOKASI BONGKAR'].str.contains('KAYAN TANAH MERAH',case=False)) | (tkdata['LOKASI MUAT'].str.contains('KAYAN TANAH MERAH',case=False))]
        tkdata['NO'] = range(1,len(tkdata)+1)
        tkdata = tkdata.reset_index(drop=True)
        dicttkdata = blankrows_tkii01(tkdata)
        writing_tkii('dtkii',xlwriter,dicttkdata,'TERSUS PT. KAYAN LNG')

    elif place == 'ptser':
        tkdata = tkdata.loc[(tkdata['LOKASI BONGKAR'].str.contains('SER MANJELUTUNG',case=False)) | (tkdata['LOKASI MUAT'].str.contains('SER MANJELUTUNG',case=False))]
        tkdata['NO'] = range(1,len(tkdata)+1)
        tkdata = tkdata.reset_index(drop=True)
        dicttkdata = blankrows_tkii01(tkdata)
        writing_tkii('dtkii',xlwriter,dicttkdata,'TERSUS PT. SARANA ENERGI')

    elif place == 'jobsi':
        tkdata = tkdata.loc[(tkdata['LOKASI BONGKAR'].str.contains('JOB SIMENGGARIS TANAH MERAH',case=False)) | (tkdata['LOKASI MUAT'].str.contains('JOB SIMENGGARIS TANAH MERAH',case=False))]
        tkdata['NO'] = range(1,len(tkdata)+1)
        tkdata = tkdata.reset_index(drop=True)
        dicttkdata = blankrows_tkii01(tkdata)
        writing_tkii('dtkii',xlwriter,dicttkdata,'TERSUS JOB SIMENGGARIS')
    
    elif place == 'ptlim':
        tkdom,tkexp = tkdata

        tkdom = tkdom.loc[(tkdom['LOKASI BONGKAR'].str.contains('LAMINDO INTER MULTIKON',case=False)) | (tkdom['LOKASI MUAT'].str.contains('LAMINDO INTER MULTIKON',case=False))]
        tkdom['NO'] = range(1,len(tkdom)+1)
        tkdom = tkdom.reset_index(drop=True)
        dicttkdom = blankrows_tkii01(tkdom)

        tkexp = tkexp.loc[(tkexp['LOKASI BONGKAR'].str.contains('LAMINDO INTER MULTIKON',case=False)) | (tkexp['LOKASI MUAT'].str.contains('LAMINDO INTER MULTIKON',case=False))]
        tkexp['NO'] = range(1,len(tkexp)+1)
        tkexp = tkexp.reset_index(drop=True)
        dicttkexp = blankrows_tkii02(tkexp)
        
        writing_tkii('detkii',xlwriter,[dicttkdom,dicttkexp],'TERSUS PT. LIM BUNYU')

    elif place == 'ptgtb':
        tkdom,tkexp = tkdata

        tkdom = tkdom.loc[(tkdom['LOKASI BONGKAR'].str.contains('GARDA TUJUH BUANA',case=False)) | (tkdom['LOKASI MUAT'].str.contains('GARDA TUJUH BUANA',case=False))]
        tkdom['NO'] = range(1,len(tkdom)+1)
        tkdom = tkdom.reset_index(drop=True)
        dicttkdom = blankrows_tkii01(tkdom)

        tkexp = tkexp.loc[(tkexp['LOKASI BONGKAR'].str.contains('GARDA TUJUH BUANA',case=False)) | (tkexp['LOKASI MUAT'].str.contains('GARDA TUJUH BUANA',case=False))]
        tkexp['NO'] = range(1,len(tkexp)+1)
        tkexp = tkexp.reset_index(drop=True)
        dicttkexp = blankrows_tkii02(tkexp)
        
        writing_tkii('detkii',xlwriter,[dicttkdom,dicttkexp],'TERSUS PT. GARDA')

    elif place == 'butkii':
        dicttkbun = blankrows_tkii01(tkdata)
        writing_tkii(place,xlwriter,dicttkbun,'B')

    elif place == 'tmtkii':
        dicttktm = blankrows_tkii03(tkdata)
        writing_tkii(place,xlwriter,dicttktm,'T')
        
# Writer
def writing_tkii(mode,xlwriter,givendata,sheetname):
    xlworkbk = xlwriter.book

    #WORKSHEET FORMATS

    #title
    fmt_title_regu = xlworkbk.add_format({'font_name':'arial','font_size':20,'align':'center','valign':'vcenter'})
    fmt_title_bold = xlworkbk.add_format({'bold':True,'font_name':'arial','font_size':20,'align':'center','valign':'vcenter'})
    fmt_title_bold_left = xlworkbk.add_format({'bold':True,'font_name':'arial','font_size':20,'align':'left','valign':'vcenter'})
    fmt_title_regu_unli = xlworkbk.add_format({'underline':True,'font_name':'arial','font_size':20,'align':'center','valign':'vcenter'})
    fmt_title_bold_unli = xlworkbk.add_format({'bold':True,'underline':True,'font_name':'arial','font_size':20,'align':'center','valign':'vcenter'})

    #header
    fmt_lheader = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':18,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_mheader = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':18,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':1,'align':'center','valign':'vcenter'})
    fmt_rheader = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':18,'text_wrap':True,'bottom':6,'top':6,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_sheader = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':18,'text_wrap':True,'bottom':6,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #filler
    fmt_lfiller = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':18,'text_wrap':True,'bottom':1,'top':6,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rfiller = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':18,'text_wrap':True,'bottom':1,'top':6,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_mfiller = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':18,'text_wrap':True,'bottom':1,'top':6,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #main data upper part
    fmt_lumain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rumain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':1,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_mumain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #main data middle part
    fmt_lmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':7,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_mmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #main data downer part
    fmt_ldmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rdmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':6,'top':7,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_mdmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #main data name and agent's name section
    fmt_ulefts = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'left','valign':'vcenter'})
    fmt_mlefts = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'left','valign':'vcenter'})
    fmt_dlefts = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'left','valign':'vcenter'})

    #main data gross tonnage section
    fmt_urights = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'right','valign':'vcenter'})
    fmt_mrights = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'right','valign':'vcenter'})
    fmt_drights = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'right','valign':'vcenter'})

    #main data dates
    fmt_udates = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'dd/mm/yyyy'})
    fmt_mdates = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'dd/mm/yyyy'})
    fmt_ddates = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'dd/mm/yyyy'})

    #main data times
    fmt_utimes = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'hh:mm'})
    fmt_mtimes = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'hh:mm'})
    fmt_dtimes = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'hh:mm'})
    
    #main data splitter
    fmt_lsplt = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':6,'align':'center','valign':'vcenter','bg_color':'orange'})
    fmt_rsplt = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':7,'right':6,'left':1,'align':'center','valign':'vcenter','bg_color':'orange'})
    fmt_msplt = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter','bg_color':'orange'})

    #summary format
    fmt_hdnume = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':1,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter'})
    fmt_sumnam = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':20,'text_wrap':True,'bottom':1,'top':1,'right':1,'left':1,'align':'left','valign':'vcenter'})
    
    #signing format
    fmt_signbold = xlworkbk.add_format({'font_name':'arial','font_size':30,'bold':True})
    fmt_signdflt = xlworkbk.add_format({'font_name':'arial','font_size':30,'bold':False})
    
    #REQUIRED VARIABLES

    #month and year
    prevmonth = str(trnslt.translate(dt.strftime(dt.today().replace(day=1) - td(days=1),'%B')).upper())
    prevyear = dt.strftime(dt.today().replace(day=1) - td(days=1),'%Y')

    #list of titles, header's rowcol, contents, and format sequences
    tiwmtkii,ctwmtkii = ['A1:D1','E1:V1','A2:D2','E2:V2','A3:D3','Y7:Z8'],['KEMENTERIAN PERHUBUNGAN',
                        'L A P O R A N  B U L A N A N  K E G I A T A N  O P E R A S I O N A L  D I  P E L A B U H A N',
                        'DIREKTORAT JENDERAL PERHUBUNGAN LAUT','Y A N G  T I D A K  D I U S A H A K A N  T K. I I  U P T',
                        'UNIT PENYELENGGARA PELABUHAN PULAU BUNYU','TK II UPT']
    tifmtwm = [fmt_title_regu,fmt_title_bold,fmt_title_regu,fmt_title_bold_unli,fmt_title_regu_unli,fmt_title_regu]
    tinmtkii,ctnmtkii = ['L3','M3','N3','L4','M4','N4'],['BULAN',':',prevmonth,'TAHUN',':',prevyear]
    tifmtnm = [fmt_title_bold,fmt_title_bold,fmt_title_bold,fmt_title_bold,fmt_title_bold,fmt_title_bold]

    henmtkii,chnmtkii = ['F9','N11','Q11','T11','W11'],['UKURAN KAPAL','Jenis Brg/Hewan',
                        'Jenis Brg/Hewan','Jenis Brg/Hewan','Jenis Brg/Hewan']
    hefmtnm = [fmt_mheader,fmt_mheader,fmt_mheader,fmt_mheader,fmt_mheader]
    hewmtkii = ['A9:A11','B9:E9','B10:C11','D10:D11','E10:E11','F10:F11','G9:I9','G10:G11',
                'H10:H11','I10:I11','J9:K9','J10:J11','K10:K11','L9:M9','L10:L11','M10:M11',
                'N9:S9','N10:P10','O11:P11','Q10:S10','R11:S11','T9:Y9','T10:V10','U11:V11',
                'W10:Y10','X11:Y11','Z9:Z11']
    chwmtkii = ['NO','N A M A','Kapal','Bendera','Pemilik/Agen','GT','T I B A',
                'Tanggal','Jam','Pelabuhan Asal','T A M B A T','Tanggal','Jam',
                'B E R A N G K A T','Tanggal','Pelabuhan Tujuan','P E R D A G A N G A N  D A L A M  N E G E R I',
                'B O N G K A R','Jumlah Muatan','M U A T','Jumlah Muatan','P E R D A G A N G A N  L U A R  N E G E R I',
                'B O N G K A R','Jumlah Muatan','M U A T','Jumlah Muatan','Ket']
    hefmtwm = [fmt_lheader,fmt_mheader,fmt_sheader,fmt_sheader,fmt_sheader,fmt_sheader,fmt_mheader,fmt_sheader,
               fmt_sheader,fmt_sheader,fmt_mheader,fmt_sheader,fmt_sheader,fmt_mheader,fmt_sheader,fmt_sheader,
               fmt_mheader,fmt_sheader,fmt_sheader,fmt_mheader,fmt_sheader,fmt_mheader,fmt_sheader,fmt_sheader,
               fmt_sheader,fmt_sheader,fmt_rheader]

    fiwmtkii = ['B12:C12','O12:P12','R12:S12','U12:V12','X12:Y12']
    cfwmtkii = ['(2)','(14)','(17)','(19)','(20)']
    fitfmtwm = [fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller]
    finmtkii = ['A12','D12','E12','F12','G12','H12','I12','J12','K12','L12','M12',
                'N12','Q12','T12','W12','Z12']
    cfnmtkii = ['(1)','(3)','(4)','(5)','(7)','(8)','(9)','(10)','(11)','(12)','(12)',
                '(13)','(15)','(18)','(110)','(21)']
    fitfmtnm = [fmt_lfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,
                fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,
                fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_rfiller]

    #list of column's width and row's height
    cowitkii = [72,96,480,270,545,123,181,110,301,170,124,200,302,565,176,130,
                525,205,136,142,133,67,220,145,113,95]
    rohetkii = [33,33,33,33,33,33,33,33,113,33,66,33,70]

    if mode=='dtkii':
        #working on load summary
        dg = goodstkii(givendata,'dom')

        #naming worksheet
        worksh = xlworkbk.add_worksheet(sheetname)

        #writing titles, headers, and fillers
        thflist = [[tiwmtkii, ctwmtkii, tifmtwm, tinmtkii, ctnmtkii, tifmtnm],
                   [hewmtkii, chwmtkii, hefmtwm, henmtkii, chnmtkii, hefmtnm],
                   [fiwmtkii, cfwmtkii, fitfmtwm, finmtkii, cfnmtkii, fitfmtnm]]
        
        for num in range(3):
            rowcols, contents, fmts = thflist[num][0], thflist[num][1], thflist[num][2]
            for val in range(len(rowcols)):
                worksh.merge_range(rowcols[val],contents[val],fmts[val])

            rowcols, contents, fmts = thflist[num][3], thflist[num][4], thflist[num][5]
            for val in range(len(rowcols)):
                worksh.write(rowcols[val],contents[val],fmts[val])

        worksh.merge_range('E6:V6',sheetname,fmt_title_bold_unli)

        #writing main data [domestic]
        ini = 0
        for keys in givendata: 
            for i in range(len(givendata[keys])):
                if keys == 'No':
                    if i == 0:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_lumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_ldmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_lmain)

                elif keys in ['Tgl Tiba','Tgl Tambat', 'Tgl Tolak']:
                    if i == 0:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_udates)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_ddates)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_mdates)

                elif keys in ['Jam Tiba','Jam Tambat']:
                    if i == 0:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_utimes)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_dtimes)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_mtimes)

                elif keys in ['Nama Kapal','GT','Keagenan']:
                    if i == 0:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_ulefts)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_dlefts)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_mlefts)

                elif keys == 'KET':
                    if i == 0:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_rumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_rdmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_rmain)

                else:
                    if i == 0:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_mumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_mdmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_mmain)
            ini += 1

        #set column's width and row's height
        for num in range(len(cowitkii)):
            worksh.set_column_pixels(num,num,cowitkii[num])

        for num in range(len(rohetkii)):
            if num == len(rohetkii)-1:
                for i in range(len(givendata['No'])):
                    worksh.set_row_pixels(i+12,rohetkii[num])
            elif num != len(rohetkii)-1:
                worksh.set_row_pixels(num,rohetkii[num])

        #write loads summary
        #calculating row
        domrhdroco = ['B'+str(17+len(givendata['GT']))+':E'+str(17+len(givendata['GT'])),
                      'B'+str(18+len(givendata['GT']))+':C'+str(18+len(givendata['GT']))]
        donmhdroco = ['D'+str(18+len(givendata['GT'])),'E'+str(18+len(givendata['GT']))]
        docmhdroco, docnhdroco = ['JUMLAH MUATAN KAPAL DALAM NEGERI','JENIS BARANG'],['BONGKAR','MUAT']

        exmrhdroco = ['G'+str(17+len(givendata['GT']))+':N'+str(17+len(givendata['GT'])),
                      'G'+str(18+len(givendata['GT']))+':I'+str(18+len(givendata['GT'])),
                      'J'+str(18+len(givendata['GT']))+':L'+str(18+len(givendata['GT'])),
                      'M'+str(18+len(givendata['GT']))+':N'+str(18+len(givendata['GT']))]
        excmhdroco = ['JUMLAH MUATAN KAPAL LUAR NEGERI','JENIS BARANG','BONGKAR','MUAT']
        
        for val in range(len(domrhdroco)):
            worksh.merge_range(domrhdroco[val],docmhdroco[val],fmt_hdnume)
        for val in range(len(donmhdroco)):
            worksh.write(donmhdroco[val],docnhdroco[val],fmt_hdnume)

        for col in range(3):
            if col == 0:
                for row in range(len(dg[col])):
                    roco = 'B'+str(19+len(givendata['GT'])+row)+':C'+str(19+len(givendata['GT'])+row)
                    worksh.merge_range(roco,dg[col][row],fmt_sumnam)
            elif col == 1:
                for row in range(len(dg[col])):
                    roco = 'D'+str(19+len(givendata['GT'])+row)
                    worksh.write(roco,dg[col][row],fmt_hdnume)
            elif col == 2:
                for row in range(len(dg[col])):
                    roco = 'E'+str(19+len(givendata['GT'])+row)
                    worksh.write(roco,dg[col][row],fmt_hdnume)

        for val in range(len(exmrhdroco)):
            worksh.merge_range(exmrhdroco[val],excmhdroco[val],fmt_hdnume)
            
        roco = 'G'+str(19+len(givendata['GT']))+':I'+str(19+len(givendata['GT']))
        worksh.merge_range(roco,'NIHIL',fmt_hdnume)
        roco = 'J'+str(19+len(givendata['GT']))+':L'+str(19+len(givendata['GT']))
        worksh.merge_range(roco,'--',fmt_hdnume)
        roco = 'M'+str(19+len(givendata['GT']))+':N'+str(19+len(givendata['GT']))
        worksh.merge_range(roco,'--',fmt_hdnume)

        #write signing place
        datetext = 'Pulau Bunyu, '+str(dt.strftime(dt.today(),'%d'))+' '+str(trnslt.translate(dt.strftime(dt.today(),'%B')))+' '+str(dt.strftime(dt.today(),'%Y'))
        signtext = [None,'Kepala Kantor','Unit Penyelenggara Pelabuhan','Kelas III Pulau Bunyu',
                    None,None,None,None,None,'Abdul Wahid','NIP. 19710515 199803 1 006']

        worksh.write(12+len(givendata[keys])+4,18,datetext,fmt_signdflt)
        worksh.write_column(12+len(givendata[keys])+6,18,signtext,fmt_signbold)

    elif mode=='etkii':
        #divide data and calculate load's summary
        diexp = givendata
        eg = goodstkii(diexp,'exp')

        #naming worksheet
        worksh = xlworkbk.add_worksheet(sheetname)

        #writing titles, headers, and fillers
        thflist = [[tiwmtkii, ctwmtkii, tifmtwm, tinmtkii, ctnmtkii, tifmtnm],
                   [hewmtkii, chwmtkii, hefmtwm, henmtkii, chnmtkii, hefmtnm],
                   [fiwmtkii, cfwmtkii, fitfmtwm, finmtkii, cfnmtkii, fitfmtnm]]
        
        for num in range(3):
            rowcols, contents, fmts = thflist[num][0], thflist[num][1], thflist[num][2]
            for val in range(len(rowcols)):
                worksh.merge_range(rowcols[val],contents[val],fmts[val])

            rowcols, contents, fmts = thflist[num][3], thflist[num][4], thflist[num][5]
            for val in range(len(rowcols)):
                worksh.write(rowcols[val],contents[val],fmts[val])

        worksh.merge_range('E6:V6',sheetname,fmt_title_bold_unli)

        #writing main data [export]
        ini = 0
        for keys in diexp:
            for i in range(len(diexp[keys])):
                if keys == 'No':
                    if i == len(diexp[keys])-1:
                        worksh.write(i+12,ini,diexp[keys][i],fmt_ldmain)
                    else:
                        worksh.write(i+12,ini,diexp[keys][i],fmt_lmain)

                elif keys in ['Tgl Tiba','Tgl Tambat', 'Tgl Tolak']:
                    if i == len(diexp[keys])-1:
                        worksh.write(i+12,ini,diexp[keys][i],fmt_ddates)
                    else:
                        worksh.write(i+12,ini,diexp[keys][i],fmt_mdates)

                elif keys in ['Jam Tiba','Jam Tambat']:
                    if i == len(diexp[keys])-1:
                        worksh.write(i+12,ini,diexp[keys][i],fmt_dtimes)
                    else:
                        worksh.write(i+10,ini,diexp[keys][i],fmt_mtimes)

                elif keys in ['Nama Kapal','GT','Keagenan']:
                    if i == len(diexp[keys])-1:
                        worksh.write(i+12,ini,diexp[keys][i],fmt_dlefts)
                    else:
                        worksh.write(i+12,ini,diexp[keys][i],fmt_mlefts)

                elif keys == 'KET':
                    if i == len(diexp[keys])-1:
                        worksh.write(i+12,ini,diexp[keys][i],fmt_rdmain)
                    else:
                        worksh.write(i+12,ini,diexp[keys][i],fmt_rmain)

                else:
                    if i == len(diexp[keys])-1:
                        worksh.write(i+12,ini,diexp[keys][i],fmt_mdmain)
                    else:
                        worksh.write(i+12,ini,diexp[keys][i],fmt_mmain)
            ini += 1

        #set column's width and row's height
        for num in range(len(cowitkii)):
            worksh.set_column_pixels(num,num,cowitkii[num])

        for num in range(len(rohetkii)):
            if num == len(rohetkii)-1:
                for i in range(len(diexp['No'])):
                    worksh.set_row_pixels(i+10,rohetkii[num])
            elif num != len(rohetkii)-1:
                worksh.set_row_pixels(num,rohetkii[num])

        #write loads summary
        #calculating row
        domrhdroco = ['B'+str(17+len(givendata['GT']))+':E'+str(17+len(givendata['GT'])),
                      'B'+str(18+len(givendata['GT']))+':C'+str(18+len(givendata['GT']))]
        donmhdroco = ['D'+str(18+len(givendata['GT'])),'E'+str(18+len(givendata['GT']))]
        docmhdroco, docnhdroco = ['JUMLAH MUATAN KAPAL DALAM NEGERI','JENIS BARANG'],['BONGKAR','MUAT']

        exmrhdroco = ['G'+str(17+len(givendata['GT']))+':N'+str(17+len(givendata['GT'])),
                      'G'+str(18+len(givendata['GT']))+':I'+str(18+len(givendata['GT'])),
                      'J'+str(18+len(givendata['GT']))+':L'+str(18+len(givendata['GT'])),
                      'M'+str(18+len(givendata['GT']))+':N'+str(18+len(givendata['GT']))]
        excmhdroco = ['JUMLAH MUATAN KAPAL LUAR NEGERI','JENIS BARANG','BONGKAR','MUAT']
        
        for val in range(len(domrhdroco)):
            worksh.merge_range(domrhdroco[val],docmhdroco[val],fmt_hdnume)
        for val in range(len(donmhdroco)):
            worksh.write(donmhdroco[val],docnhdroco[val],fmt_hdnume)
        
        for val in range(len(domrhdroco)):
            worksh.merge_range(domrhdroco[val],docmhdroco[val],fmt_hdnume)
        for val in range(len(donmhdroco)):
            worksh.write(donmhdroco[val],docnhdroco[val],fmt_hdnume)

        roco = 'B'+str(19+len(diexp['No']))+':C'+str(19+len(diexp['No']))
        worksh.merge_range(roco,'NIHIL',fmt_hdnume)
        roco = 'D'+str(19+len(diexp['No']))
        worksh.write(roco,'--',fmt_hdnume)
        roco = 'E'+str(19+len(diexp['No']))
        worksh.write(roco,'--',fmt_hdnume)

        for val in range(len(exmrhdroco)):
            worksh.merge_range(exmrhdroco[val],excmhdroco[val],fmt_hdnume)
            
        for col in range(3):
            if col == 0:
                for row in range(len(eg[col])):
                    roco = 'G'+str(19+len(diexp['No'])+row)+':I'+str(19+len(diexp['No'])+row)
                    worksh.merge_range(roco,eg[col][row],fmt_hdnume)
            elif col == 1:
                for row in range(len(eg[col])):
                    roco = 'J'+str(19+len(diexp['No'])+row)+':L'+str(19+len(diexp['No'])+row)
                    worksh.merge_range(roco,eg[col][row],fmt_hdnume)
            elif col == 2:
                for row in range(len(eg[col])):
                    roco = 'M'+str(19+len(diexp['No'])+row)+':N'+str(19+len(diexp['No'])+row)
                    worksh.merge_range(roco,eg[col][row],fmt_hdnume)
        
        #write signing place
        datetext = 'Pulau Bunyu, '+str(dt.strftime(dt.today(),'%d'))+' '+str(trnslt.translate(dt.strftime(dt.today(),'%B')))+' '+str(dt.strftime(dt.today(),'%Y'))
        signtext = [None,'Kepala Kantor','Unit Penyelenggara Pelabuhan','Kelas III Pulau Bunyu',
                    None,None,None,None,'Abdul Wahid','NIP. 19710515 199803 1 006']

        worksh.write(12+len(diexp['No'])+4,18,datetext,fmt_signdflt)
        worksh.write_column(12+len(diexp['No'])+5,18,signtext,fmt_signbold)

    elif mode=='detkii':
        #divide data and calculate load's summary
        didom, diexp = givendata
        dg,eg = goodstkii(didom,'dom'),goodstkii(diexp,'exp')

        #naming worksheet
        worksh = xlworkbk.add_worksheet(sheetname)

        #writing titles, headers, and fillers
        thflist = [[tiwmtkii, ctwmtkii, tifmtwm, tinmtkii, ctnmtkii, tifmtnm],
                   [hewmtkii, chwmtkii, hefmtwm, henmtkii, chnmtkii, hefmtnm],
                   [fiwmtkii, cfwmtkii, fitfmtwm, finmtkii, cfnmtkii, fitfmtnm]]
        
        for num in range(3):
            rowcols, contents, fmts = thflist[num][0], thflist[num][1], thflist[num][2]
            for val in range(len(rowcols)):
                worksh.merge_range(rowcols[val],contents[val],fmts[val])

            rowcols, contents, fmts = thflist[num][3], thflist[num][4], thflist[num][5]
            for val in range(len(rowcols)):
                worksh.write(rowcols[val],contents[val],fmts[val])

        worksh.merge_range('E6:V6',sheetname,fmt_title_bold_unli)

        #writing main data [domestic]
        ini = 0
        for keys in didom:
            for i in range(len(didom[keys])):
                if keys == 'No':
                    if i == 0:
                        worksh.write(i+12,ini,didom[keys][i],fmt_lumain)
                    elif i == len(didom[keys])-1:
                        worksh.write(i+12,ini,didom[keys][i],fmt_lmain)
                    elif i != 0 and i != len(didom[keys])-1:
                        worksh.write(i+12,ini,didom[keys][i],fmt_lmain)

                elif keys in ['Tgl Tiba','Tgl Tambat', 'Tgl Tolak']:
                    if i == 0:
                        worksh.write(i+12,ini,didom[keys][i],fmt_udates)
                    elif i == len(didom[keys])-1:
                        worksh.write(i+12,ini,didom[keys][i],fmt_mdates)
                    elif i != 0 and i != len(didom[keys])-1:
                        worksh.write(i+12,ini,didom[keys][i],fmt_mdates)

                elif keys in ['Jam Tiba','Jam Tambat']:
                    if i == 0:
                        worksh.write(i+12,ini,didom[keys][i],fmt_utimes)
                    elif i == len(didom[keys])-1:
                        worksh.write(i+12,ini,didom[keys][i],fmt_mtimes)
                    elif i != 0 and i != len(didom[keys])-1:
                        worksh.write(i+12,ini,didom[keys][i],fmt_mtimes)

                elif keys in ['Nama Kapal','GT','Keagenan']:
                    if i == 0:
                        worksh.write(i+12,ini,didom[keys][i],fmt_ulefts)
                    elif i == len(didom[keys])-1:
                        worksh.write(i+12,ini,didom[keys][i],fmt_mlefts)
                    elif i != 0 and i != len(didom[keys])-1:
                        worksh.write(i+12,ini,didom[keys][i],fmt_mlefts)

                elif keys == 'KET':
                    if i == 0:
                        worksh.write(i+12,ini,didom[keys][i],fmt_rumain)
                    elif i == len(didom[keys])-1:
                        worksh.write(i+12,ini,didom[keys][i],fmt_rmain)
                    elif i != 0 and i != len(didom[keys])-1:
                        worksh.write(i+12,ini,didom[keys][i],fmt_rmain)

                else:
                    if i == 0:
                        worksh.write(i+12,ini,didom[keys][i],fmt_mumain)
                    elif i == len(didom[keys])-1:
                        worksh.write(i+12,ini,didom[keys][i],fmt_mmain)
                    elif i != 0 and i != len(didom[keys])-1:
                        worksh.write(i+12,ini,didom[keys][i],fmt_mmain)
            ini += 1

        #writing main data [export]
        ini = 0
        for keys in diexp:
            for i in range(len(diexp[keys])):
                if keys == 'No':
                    if i == len(diexp[keys])-1:
                        worksh.write(i+12+len(didom[keys])+1,ini,diexp[keys][i],fmt_ldmain)
                    else:
                        worksh.write(i+12+len(didom[keys])+1,ini,diexp[keys][i],fmt_lmain)

                elif keys in ['Tgl Tiba','Tgl Tambat', 'Tgl Tolak']:
                    if i == len(diexp[keys])-1:
                        worksh.write(i+12+len(didom[keys])+1,ini,diexp[keys][i],fmt_ddates)
                    else:
                        worksh.write(i+12+len(didom[keys])+1,ini,diexp[keys][i],fmt_mdates)

                elif keys in ['Jam Tiba','Jam Tambat']:
                    if i == len(diexp[keys])-1:
                        worksh.write(i+12+len(didom[keys])+1,ini,diexp[keys][i],fmt_dtimes)
                    else:
                        worksh.write(i+12+len(didom[keys])+1,ini,diexp[keys][i],fmt_mtimes)

                elif keys in ['Nama Kapal','GT','Keagenan']:
                    if i == len(diexp[keys])-1:
                        worksh.write(i+12+len(didom[keys])+1,ini,diexp[keys][i],fmt_dlefts)
                    else:
                        worksh.write(i+12+len(didom[keys])+1,ini,diexp[keys][i],fmt_mlefts)

                elif keys == 'KET':
                    if i == len(diexp[keys])-1:
                        worksh.write(i+12+len(didom[keys])+1,ini,diexp[keys][i],fmt_rdmain)
                    else:
                        worksh.write(i+12+len(didom[keys])+1,ini,diexp[keys][i],fmt_rmain)

                else:
                    if i == len(diexp[keys])-1:
                        worksh.write(i+12+len(didom[keys])+1,ini,diexp[keys][i],fmt_mdmain)
                    else:
                        worksh.write(i+12+len(didom[keys])+1,ini,diexp[keys][i],fmt_mmain)
            ini += 1

        #splitter between d & e main data
        for keys in range(len(didom)):
            if keys == 0:
                worksh.write(12+len(didom['GT']),keys,'',fmt_lsplt)
            elif keys == len(didom)-1:
                worksh.write(12+len(didom['GT']),keys,'',fmt_rsplt)
            elif keys != 0 and i != len(didom):
                worksh.write(12+len(didom['GT']),keys,'',fmt_msplt)

        #set column's width and row's height
        for num in range(len(cowitkii)):
            worksh.set_column_pixels(num,num,cowitkii[num])

        for num in range(len(rohetkii)):
            if num == len(rohetkii)-1:
                for i in range(len(didom['No'])+1+len(diexp['No'])):
                    worksh.set_row_pixels(i+12,rohetkii[num])
            elif num != len(rohetkii)-1:
                worksh.set_row_pixels(num,rohetkii[num])

        #write loads summary
        #calculating row
        domrhdroco = ['B'+str(17+len(didom['GT'])+1+len(diexp['No']))+':E'+str(17+len(didom['GT'])+1+len(diexp['No'])),
                      'B'+str(18+len(didom['GT'])+1+len(diexp['No']))+':C'+str(18+len(didom['GT'])+1+len(diexp['No']))]
        donmhdroco = ['D'+str(18+len(didom['GT'])+1+len(diexp['No'])),'E'+str(18+len(didom['GT'])+1+len(diexp['No']))]
        docmhdroco, docnhdroco = ['JUMLAH MUATAN KAPAL DALAM NEGERI','JENIS BARANG'],['BONGKAR','MUAT']

        exmrhdroco = ['G'+str(17+len(didom['GT'])+1+len(diexp['No']))+':N'+str(17+len(didom['GT'])+1+len(diexp['No'])),
                      'G'+str(18+len(didom['GT'])+1+len(diexp['No']))+':I'+str(18+len(didom['GT'])+1+len(diexp['No'])),
                      'J'+str(18+len(didom['GT'])+1+len(diexp['No']))+':L'+str(18+len(didom['GT'])+1+len(diexp['No'])),
                      'M'+str(18+len(didom['GT'])+1+len(diexp['No']))+':N'+str(18+len(didom['GT'])+1+len(diexp['No']))]
        excmhdroco = ['JUMLAH MUATAN KAPAL LUAR NEGERI','JENIS BARANG','BONGKAR','MUAT']
        
        for val in range(len(domrhdroco)):
            worksh.merge_range(domrhdroco[val],docmhdroco[val],fmt_hdnume)
        for val in range(len(donmhdroco)):
            worksh.write(donmhdroco[val],docnhdroco[val],fmt_hdnume)

        for col in range(3):
            if col == 0:
                for row in range(len(dg[col])):
                    roco = 'B'+str(19+len(didom['GT'])+1+len(diexp['No'])+row)+':C'+str(19+len(didom['GT'])+1+len(diexp['No'])+row)
                    worksh.merge_range(roco,dg[col][row],fmt_sumnam)
            elif col == 1:
                for row in range(len(dg[col])):
                    roco = 'D'+str(19+len(didom['GT'])+1+len(diexp['No'])+row)
                    worksh.write(roco,dg[col][row],fmt_hdnume)
            elif col == 2:
                for row in range(len(dg[col])):
                    roco = 'E'+str(19+len(didom['GT'])+1+len(diexp['No'])+row)
                    worksh.write(roco,dg[col][row],fmt_hdnume)

        for val in range(len(exmrhdroco)):
            worksh.merge_range(exmrhdroco[val],excmhdroco[val],fmt_hdnume)
            
        for col in range(3):
            if col == 0:
                for row in range(len(eg[col])):
                    roco = 'G'+str(19+len(didom['GT'])+1+len(diexp['No'])+row)+':I'+str(19+len(didom['GT'])+1+len(diexp['No'])+row)
                    worksh.merge_range(roco,eg[col][row],fmt_hdnume)
            elif col == 1:
                for row in range(len(eg[col])):
                    roco = 'J'+str(19+len(didom['GT'])+1+len(diexp['No'])+row)+':L'+str(19+len(didom['GT'])+1+len(diexp['No'])+row)
                    worksh.merge_range(roco,eg[col][row],fmt_hdnume)
            elif col == 2:
                for row in range(len(eg[col])):
                    roco = 'M'+str(19+len(didom['GT'])+1+len(diexp['No'])+row)+':N'+str(19+len(didom['GT'])+1+len(diexp['No'])+row)
                    worksh.merge_range(roco,eg[col][row],fmt_hdnume)
        

        #write signing place
        datetext = 'Pulau Bunyu, '+str(dt.strftime(dt.today(),'%d'))+' '+str(trnslt.translate(dt.strftime(dt.today(),'%B')))+' '+str(dt.strftime(dt.today(),'%Y'))
        signtext = [None,'Kepala Kantor','Unit Penyelenggara Pelabuhan','Kelas III Pulau Bunyu',
                    None,None,None,None,'Abdul Wahid','NIP. 19710515 199803 1 006']

        worksh.write(12+len(didom['GT'])+1+len(diexp['No'])+4,18,datetext,fmt_signdflt)
        worksh.write_column(12+len(didom['GT'])+1+len(diexp['No'])+5,18,signtext,fmt_signbold)
    
    elif mode=='butkii':
        #naming worksheet
        worksh = xlworkbk.add_worksheet('BUNYU')

        #writing titles, headers, and fillers
        hefmtwm[16],hefmtwm[19],hefmtwm[20],fitfmtwm[2] = fmt_rheader,fmt_rheader,fmt_rheader,fmt_rfiller
        tiwmtkii[1],tiwmtkii[3],tifmtnm[2],tifmtnm[5] = 'E1:S1','E2:S2',fmt_title_bold_left,fmt_title_bold_left
        thflist = [[tiwmtkii[:5], ctwmtkii[:5], tifmtwm[:5], tinmtkii, ctnmtkii, tifmtnm],
                   [hewmtkii[:21], chwmtkii[:21], hefmtwm[:21], henmtkii[:3], chnmtkii[:3], hefmtnm[:3]],
                   [fiwmtkii[:3], cfwmtkii[:3], fitfmtwm[:3], finmtkii[:13], cfnmtkii[:13], fitfmtnm[:13]]]
        
        for num in range(3):
            rowcols, contents, fmts = thflist[num][0], thflist[num][1], thflist[num][2]
            for val in range(len(rowcols)):
                worksh.merge_range(rowcols[val],contents[val],fmts[val])

            rowcols, contents, fmts = thflist[num][3], thflist[num][4], thflist[num][5]
            for val in range(len(rowcols)):
                worksh.write(rowcols[val],contents[val],fmts[val])

        #writing main data [domestic]
        ini = 0
        for keys in givendata: 
            for i in range(len(givendata[keys])):
                if keys == 'No':
                    if i == 0:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_lumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_ldmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_lmain)

                elif keys in ['Tgl Tiba','Tgl Tambat','Tgl Tolak']:
                    if i == 0:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_udates)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_ddates)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_mdates)

                elif keys in ['Jam Tiba','Jam Tambat']:
                    if i == 0:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_utimes)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_dtimes)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_mtimes)

                elif keys in ['Nama Kapal','GT','Keagenan']:
                    if i == 0:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_ulefts)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_dlefts)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_mlefts)

                elif keys == '1an Muat D':
                    if i == 0:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_rumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_rdmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_rmain)
                
                elif keys in ['Brg Bongkar E','Jml Bongkar E','1an Bongkar E','Brg Muat E','Jml Muat E','1an Muat E','Ket']:
                    pass
                
                elif keys in ['Kode Kapal','Bendera','Asal','Tujuan','Brg Bongkar D','Jml Bongkar D','1an Bongkar D','Brg Muat D','Jml Muat D']:
                    if i == 0:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_mumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_mdmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+12,ini,givendata[keys][i],fmt_mmain)
            ini += 1

        #set column's width and row's height
        for num in range(len(cowitkii)):
            worksh.set_column_pixels(num,num,cowitkii[num])

        for num in range(len(rohetkii)):
            if num == len(rohetkii)-1:
                for i in range(len(givendata['No'])+1):
                    worksh.set_row_pixels(i+12,rohetkii[num])
            elif num != len(rohetkii)-1:
                worksh.set_row_pixels(num,rohetkii[num])

        #write signing place
        datetext = 'Pulau Bunyu, '+str(dt.strftime(dt.today(),'%d'))+' '+str(trnslt.translate(dt.strftime(dt.today(),'%B')))+' '+str(dt.strftime(dt.today(),'%Y'))
        signtext = [None,'Kepala Kantor','Unit Penyelenggara Pelabuhan','Kelas III Pulau Bunyu',
                    None,None,None,None,None,'Abdul Wahid','NIP. 19710515 199803 1 006']

        worksh.write(18+len(givendata[keys]),15,datetext,fmt_signdflt)
        worksh.write_column(19+len(givendata[keys]),15,signtext,fmt_signbold)

    elif mode=='tmtkii':
        #naming worksheet
        worksh = xlworkbk.add_worksheet('TANPA MUATAN')

        #writing titles, headers, and fillers
        tiwmtkii,ctwmtkii = ['A1:M1','A2:M2','L3:M3'],['LAPORAN BULANAN OPERASIONAL KAPAL',
                             'KANTOR UNIT PENYELENGGARA PELABUHAN KELAS III PULAU BUNYU',
                              prevmonth+' '+prevyear]
        tifmtwm = [fmt_title_bold,fmt_title_bold_unli,fmt_title_bold_unli]

        henmtkii,chnmtkii = ['D6','E6','F5','F6','G6','H6','I6','J6','K6','L6','M6'],['Bendera',
                             'Pemilik/Agen','UKURAN KAPAL','GT','Tanggal','Jam','Pelabuhan Asal',
                             'Tanggal','Jam','Tanggal','Pelabuhan Tujuan']
        hefmtnm = [fmt_sheader,fmt_sheader,fmt_mheader,fmt_sheader,fmt_sheader,fmt_sheader,
                   fmt_sheader,fmt_sheader,fmt_sheader,fmt_sheader,fmt_rheader]

        hewmtkii,chwmtkii = ['A5:A6','B5:E5','B6:C6','G5:I5','J5:K5','L5:M5'],['NO','N A M A',
                             'Kapal','T I B A','T A M B A T','B E R A N G K A T']
        hefmtwm = [fmt_lheader,fmt_mheader,fmt_sheader,fmt_mheader,fmt_mheader,fmt_rheader]

        fiwmtkii,cfwmtkii,fitfmtwm = ['B7:C7'],['(2)'],[fmt_mfiller]
        finmtkii = ['A7','D7','E7','F7','G7','H7','I7','J7','K7','L7','M7']
        cfnmtkii = ['(1)','(3)','(4)','(5)','(6)','(7)','(8)','(9)','(10)','(11)','(12)']
        fitfmtnm = [fmt_lfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,
                    fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_rfiller]

        thflist = [[tiwmtkii, ctwmtkii, tifmtwm, tinmtkii, ctnmtkii, tifmtnm],
                   [hewmtkii, chwmtkii, hefmtwm, henmtkii, chnmtkii, hefmtnm],
                   [fiwmtkii, cfwmtkii, fitfmtwm, finmtkii, cfnmtkii, fitfmtnm]]
        
        for num in range(3):
            rowcols, contents, fmts = thflist[num][0], thflist[num][1], thflist[num][2]
            for val in range(len(rowcols)):
                worksh.merge_range(rowcols[val],contents[val],fmts[val])

            rowcols, contents, fmts = thflist[num][3], thflist[num][4], thflist[num][5]
            if num == 0:
                pass
            elif num != 0:
                for val in range(len(rowcols)):
                    worksh.write(rowcols[val],contents[val],fmts[val])

        #writing main data [domestic]
        ini = 0
        for keys in givendata: 
            for i in range(len(givendata[keys])):
                if keys == 'No':
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_lumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_ldmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_lmain)

                elif keys in ['Tgl Tiba','Tgl Tambat','Tgl Tolak']:
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_udates)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_ddates)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mdates)

                elif keys in ['Jam Tiba','Jam Tambat']:
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_utimes)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_dtimes)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mtimes)

                elif keys in ['Nama Kapal','GT','Keagenan']:
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_ulefts)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_dlefts)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mlefts)

                elif keys == 'Tujuan':
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_rumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_rdmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_rmain)
                
                elif keys in ['Brg Bongkar D','Jml Bongkar D','1an Bongkar D','Brg Muat D','Jml Muat D','Brg Bongkar E','Jml Bongkar E','1an Bongkar E','Brg Muat E','Jml Muat E','1an Muat E','Ket']:
                    pass
                
                elif keys in ['Kode Kapal','Bendera','Asal']:
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mdmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mmain)
            ini += 1

        #set column's width and row's height
        for num in range(len(cowitkii)):
            worksh.set_column_pixels(num,num,cowitkii[num])

        rohetkii = [33,100,33,7,113,66,33,70]
        for num in range(len(rohetkii)):
            if num == len(rohetkii)-1:
                for i in range(len(givendata['No'])+1):
                    worksh.set_row_pixels(i+7,rohetkii[num])
            elif num != len(rohetkii)-1:
                worksh.set_row_pixels(num,rohetkii[num])

        #write signing place
        datetext = 'Pulau Bunyu, '+str(dt.strftime(dt.today(),'%d'))+' '+str(trnslt.translate(dt.strftime(dt.today(),'%B')))+' '+str(dt.strftime(dt.today(),'%Y'))
        signtext = [None,'Kepala Kantor','Unit Penyelenggara Pelabuhan','Kelas III Pulau Bunyu',
                    None,None,None,None,None,'Abdul Wahid','NIP. 19710515 199803 1 006']
        
        fmt_signbold = xlworkbk.add_format({'font_name':'arial','font_size':25,'bold':True})
        fmt_signdflt = xlworkbk.add_format({'font_name':'arial','font_size':25,'bold':False})

        worksh.write(10+len(givendata[keys]),9,datetext,fmt_signdflt)
        worksh.write_column(11+len(givendata[keys]),9,signtext,fmt_signbold)

#3 Domestic & Export style report
# Handler
def dmstc_based(rwdata,mode,xlwriter):
    if mode == 'dmstc':
        dictrwdata = blankrows_dmstcs(rwdata)
        writing_dmstc(xlwriter,'no-sh',dictrwdata,'DOMESTIK')
    elif mode == 'ekspr':
        dictrwdata = blankrows_export(rwdata)
        writing_dmstc(xlwriter,'wi-sh',dictrwdata,'EKSPOR')
    elif mode == 'sawit':
        rwdata = rwdata.loc[(rwdata['JENIS BARANG'].str.contains('(SWT)',case=False)) | (rwdata['JENIS MUATAN'].str.contains('(SWT)',case=False))]
        rwdata['NO'] = range(1,len(rwdata)+1)
        rwdata = rwdata.reset_index(drop=True)
        dictrwdata = blankrows_dmstcs(rwdata)
        writing_dmstc(xlwriter,'no-sh',dictrwdata,'SAWIT')
    elif mode == 'batbar':
        rwdata = rwdata.loc[(rwdata['JENIS BARANG'].str.contains('(BABA)',case=False)) | (rwdata['JENIS MUATAN'].str.contains('(BABA)',case=False))]
        rwdata['NO'] = range(1,len(rwdata)+1)
        rwdata = rwdata.reset_index(drop=True)
        dictrwdata = blankrows_dmstcs(rwdata)
        writing_dmstc(xlwriter,'no-sh',dictrwdata,'BATUBARA')
    elif mode == 'gencar':
        rwdata = rwdata.loc[(rwdata['JENIS BARANG'].str.contains('(GEAR)',case=False)) | (rwdata['JENIS MUATAN'].str.contains('(GEAR)',case=False))]
        rwdata['NO'] = range(1,len(rwdata)+1)
        rwdata = rwdata.reset_index(drop=True)
        dictrwdata = blankrows_dmstcs(rwdata)
        writing_dmstc(xlwriter,'no-sh',dictrwdata,'GENCAR')
    elif mode == 'batcah':
        rwdata = rwdata.loc[(rwdata['JENIS BARANG'].str.contains('(BAPE)',case=False)) | (rwdata['JENIS MUATAN'].str.contains('(BAPE)',case=False))]
        rwdata['NO'] = range(1,len(rwdata)+1)
        rwdata = rwdata.reset_index(drop=True)
        dictrwdata = blankrows_dmstcs(rwdata)
        writing_dmstc(xlwriter,'no-sh',dictrwdata,'BATU PECAH')
    elif mode == 'cruoil':
        rwdata = rwdata.loc[(rwdata['JENIS BARANG'].str.contains('(CRIL)',case=False)) | (rwdata['JENIS MUATAN'].str.contains('(CRIL)',case=False))]
        rwdata['NO'] = range(1,len(rwdata)+1)
        rwdata = rwdata.reset_index(drop=True)
        dictrwdata = blankrows_dmstcs(rwdata)
        writing_dmstc(xlwriter,'no-sh',dictrwdata,'CRUDE OIL')
    elif mode == 'alaber':
        rwdata = rwdata.loc[(rwdata['JENIS BARANG'].str.contains('(ALBE)',case=False)) | (rwdata['JENIS MUATAN'].str.contains('(ALBE)',case=False))]
        rwdata['NO'] = range(1,len(rwdata)+1)
        rwdata = rwdata.reset_index(drop=True)
        dictrwdata = blankrows_dmstcs(rwdata)
        writing_dmstc(xlwriter,'no-sh',dictrwdata,'ALAT BERAT')
    elif mode == 'bebeem':
        rwdata = rwdata.loc[(rwdata['JENIS BARANG'].str.contains('(BBM)',case=False)) | (rwdata['JENIS MUATAN'].str.contains('(BBM)',case=False))]
        rwdata['NO'] = range(1,len(rwdata)+1)
        rwdata = rwdata.reset_index(drop=True)
        dictrwdata = blankrows_dmstcs(rwdata)
        writing_dmstc(xlwriter,'no-sh',dictrwdata,'BBM')
    elif mode == 'kendrn':
        rwdata = rwdata.loc[(rwdata['JENIS BARANG'].str.contains('(KNDR)',case=False)) | (rwdata['JENIS MUATAN'].str.contains('(KNDR)',case=False))]
        rwdata['NO'] = range(1,len(rwdata)+1)
        rwdata = rwdata.reset_index(drop=True)
        dictrwdata = blankrows_dmstcs(rwdata)
        writing_dmstc(xlwriter,'no-sh',dictrwdata,'MOBIL MOTOR')
    elif mode == 'kaywoo':
        rwdata = rwdata.loc[(rwdata['JENIS BARANG'].str.contains('(KAYU)',case=False)) | (rwdata['JENIS MUATAN'].str.contains('(KAYU)',case=False))]
        rwdata['NO'] = range(1,len(rwdata)+1)
        rwdata = rwdata.reset_index(drop=True)
        dictrwdata = blankrows_dmstcs(rwdata)
        writing_dmstc(xlwriter,'no-sh',dictrwdata,'KAYU')
    elif mode == 'tansan':
        rwdata = rwdata.loc[(rwdata['JENIS BARANG'].str.contains('(TNH)',case=False)) | (rwdata['JENIS MUATAN'].str.contains('(TNH)',case=False))]
        rwdata['NO'] = range(1,len(rwdata)+1)
        rwdata = rwdata.reset_index(drop=True)
        dictrwdata = blankrows_dmstcs(rwdata)
        writing_dmstc(xlwriter,'no-sh',dictrwdata,'TANAH')
    elif mode == 'campur':
        rwdata = rwdata.loc[(rwdata['JENIS BARANG'].str.contains('(CMPR)',case=False)) | (rwdata['JENIS MUATAN'].str.contains('(CMPR)',case=False))]
        rwdata['NO'] = range(1,len(rwdata)+1)
        rwdata = rwdata.reset_index(drop=True)
        dictrwdata = blankrows_dmstcs(rwdata)
        writing_dmstc(xlwriter,'no-sh',dictrwdata,'CAMPURAN')
        
# writer
def writing_dmstc(xlwriter,mode,givendata,sheetname):
    xlworkbk = xlwriter.book
    
    #WORKSHEET FORMATS

    #title
    fmt_title_bold = xlworkbk.add_format({'bold':True,'font_name':'arial','font_size':12,'align':'center','valign':'vcenter'})

    #header
    fmt_lheader = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_mheader = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':1,'align':'center','valign':'vcenter'})
    fmt_mshhead = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':1,'top':6,'right':1,'left':1,'align':'center','valign':'vcenter'})
    fmt_rshhead = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':1,'top':6,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_rschead = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':1,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_rheader = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_sheader = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #filler
    fmt_lfiller = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rfiller = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_mfiller = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #main data upper part
    fmt_lumain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rumain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_mumain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #main data middle part
    fmt_lmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_mmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #main data downer part
    fmt_ldmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rdmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_mdmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #main data left align section
    fmt_ulefts = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'left','valign':'vcenter'})
    fmt_mlefts = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'left','valign':'vcenter'})
    fmt_dlefts = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'left','valign':'vcenter'})

    #main data end of line part
    fmt_lemain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_remain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_memain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':1,'align':'left','valign':'vcenter','bold':True})

    #main data dates
    fmt_udates = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'dd/mm/yyyy'})
    fmt_mdates = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'dd/mm/yyyy'})
    fmt_ddates = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'dd/mm/yyyy'})
    
    #main data splitter
    fmt_lsplt = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':6,'align':'center','valign':'vcenter','bg_color':'orange'})
    fmt_rsplt = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':6,'left':1,'align':'center','valign':'vcenter','bg_color':'orange'})
    fmt_msplt = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter','bg_color':'orange'})

    #summary format
    fmt_hdnume = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':1,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter','bold':True})
    fmt_sumnam = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':1,'top':1,'right':1,'left':1,'align':'left','valign':'vcenter','bold':True})
    
    #signing format
    fmt_signbold = xlworkbk.add_format({'font_name':'arial','font_size':16,'bold':True})
    fmt_signdflt = xlworkbk.add_format({'font_name':'arial','font_size':16,'bold':False})

    #REQUIRED VARIABLES

    #month and year
    prevmonth = str(trnslt.translate(dt.strftime(dt.today().replace(day=1) - td(days=1),'%B')).upper())
    prevyear = dt.strftime(dt.today().replace(day=1) - td(days=1),'%Y')

    #list of titles, header's rowcol, contents, and format sequences
    tiwmdmst,ctwmdmst = ['A1:Q1','A2:Q2','A3:Q3'],['L A P O R A N  ASAL DAN TUJUAN PENUMPANG, HEWAN, DAN BARANG UNTUK ANTAR PULAU',
                        'PELABUHAN BUNYU',prevmonth+' '+prevyear]
    tifmdwm = [fmt_title_bold,fmt_title_bold,fmt_title_bold]

    hewmdmst = ['A5:A6','B5:D5','B6:C6','E5:E6','F5:F6','G5:G6','H5:H6','I5:I6','J5:M5','K6:L6','N5:Q5','O6:P6']
    chwmdmst = ['NO','N A M A','Kapal','Bendera','GT','Trayek','Tgl Tiba','Tgl Berangkat',
                'B O N G K A R','Volume Org/MT/Ton','M U A T','Volume Org/MT/Ton']
    hefmdwm = [fmt_lheader,fmt_mshhead,fmt_sheader,fmt_mheader,fmt_mheader,fmt_mheader,
               fmt_mheader,fmt_mheader,fmt_mshhead,fmt_sheader,fmt_rshhead,fmt_mheader]

    henmdmst,chnmdmst = ['D6','J6','M6','N6','Q6'],['Pemilik/Agen','Jenis','Asal','Jenis','Tujuan']
    hefmdnm = [fmt_sheader,fmt_sheader,fmt_sheader,fmt_sheader,fmt_rschead]

    fiwmdmst,cfwmdmst,fitfmdwm = ['B7:C7','K7:L7','O7:P7'],['(2)','(10)','(13)'],[fmt_mfiller,fmt_mfiller,fmt_mfiller]
    finmdmst = ['A7','D7','E7','F7','G7','H7','I7','J7','M7','N7','Q7']
    cfnmdmst = ['(1)','(3)','(4)','(5)','(6)','(7)','(8)','(9)','(11)','(12)','(14)']
    fitfmdnm = [fmt_lfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,
                fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_rfiller]

    #list of column's width and row's height
    cowidmst = [50,60,255,345,205,105,70,125,125,240,105,70,180,240,105,70,180]
    rohedmst = [25,25,25,25,25,50,30,45]

    if mode == 'no-sh':
        #working on load summary
        dg = goodstkii(givendata,'dom')

        #naming worksheet
        worksh = xlworkbk.add_worksheet(sheetname)

        #writing titles, headers, and fillers
        dhflist = [[tiwmdmst, ctwmdmst, tifmdwm, None, None, None],
                   [hewmdmst, chwmdmst, hefmdwm, henmdmst, chnmdmst, hefmdnm],
                   [fiwmdmst, cfwmdmst, fitfmdwm, finmdmst, cfnmdmst, fitfmdnm]]

        for num in range(3):
            rowcols, contents, fmts = dhflist[num][0], dhflist[num][1], dhflist[num][2]
            for val in range(len(rowcols)):
                worksh.merge_range(rowcols[val],contents[val],fmts[val])

            rowcols, contents, fmts = dhflist[num][3], dhflist[num][4], dhflist[num][5]
            if num == 0:
                pass
            elif num != 0:
                for val in range(len(rowcols)):
                    worksh.write(rowcols[val],contents[val],fmts[val])

        #writing main data
        ini = 0
        for keys in givendata:
            for i in range(len(givendata[keys])):
                if keys == 'No':
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_lumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_ldmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_lmain)
                
                elif keys in ['Tgl Tiba', 'Tgl Tolak']:
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_udates)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_ddates)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mdates)

                elif keys in ['Nama Kapal','GT','Keagenan']:
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_ulefts)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_dlefts)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mlefts)

                elif keys == 'Tujuan':
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_rumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_rdmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_rmain)

                else:
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mdmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mmain)

            ini += 1

        for col in range(len(givendata)):
            if col == 0:
                worksh.write(7+len(givendata['No']),col,None,fmt_lemain)
            elif col == len(givendata)-1:
                worksh.write(7+len(givendata['No']),col,None,fmt_remain)
            elif col > 0 and col < len(givendata)-1:
                worksh.write(7+len(givendata['No']),col,None,fmt_memain)

        worksh.write(7+len(givendata['No']),5,'=SUM(F8:F'+str(7+len(givendata['No']))+')',fmt_memain)

        #set column's width and row's height
        for num in range(len(cowidmst)):
            worksh.set_column_pixels(num,num,cowidmst[num])

        for num in range(len(rohedmst)):
            if num == len(rohedmst)-1:
                for i in range(len(givendata['No'])+1):
                    worksh.set_row_pixels(i+7,rohedmst[num])
            elif num != len(rohedmst)-1:
                worksh.set_row_pixels(num,rohedmst[num])

        #write loads summary
        #calculating row
        domrhdroco = ['B'+str(11+len(givendata['GT']))+':E'+str(11+len(givendata['GT'])),
                      'B'+str(12+len(givendata['GT']))+':C'+str(12+len(givendata['GT']))]
        donmhdroco = ['D'+str(12+len(givendata['GT'])),'E'+str(12+len(givendata['GT']))]
        docmhdroco, docnhdroco = ['JUMLAH MUATAN KAPAL DALAM NEGERI','JENIS BARANG'],['BONGKAR','MUAT']
            
        for val in range(len(domrhdroco)):
            worksh.merge_range(domrhdroco[val],docmhdroco[val],fmt_hdnume)
        for val in range(len(donmhdroco)):
            worksh.write(donmhdroco[val],docnhdroco[val],fmt_hdnume)

        for col in range(3):
            if col == 0:
                for row in range(len(dg[col])):
                    roco = 'B'+str(13+len(givendata['GT'])+row)+':C'+str(13+len(givendata['GT'])+row)
                    worksh.merge_range(roco,dg[col][row],fmt_sumnam)
            elif col == 1:
                for row in range(len(dg[col])):
                    roco = 'D'+str(13+len(givendata['GT'])+row)
                    worksh.write(roco,dg[col][row],fmt_hdnume)
            elif col == 2:
                for row in range(len(dg[col])):
                    roco = 'E'+str(13+len(givendata['GT'])+row)
                    worksh.write(roco,dg[col][row],fmt_hdnume)

        #write signing place
        datetext = 'Pulau Bunyu, '+str(dt.strftime(dt.today(),'%d'))+' '+str(trnslt.translate(dt.strftime(dt.today(),'%B')))+' '+str(dt.strftime(dt.today(),'%Y'))
        signtext = [None,'Kepala Kantor','Unit Penyelenggara Pelabuhan','Kelas III Pulau Bunyu',
                    None,None,None,None,None,'Abdul Wahid','NIP. 19710515 199803 1 006']

        worksh.write(10+len(givendata[keys]),13,datetext,fmt_signdflt)
        worksh.write_column(10+len(givendata[keys])+1,13,signtext,fmt_signbold)
    
    elif mode == 'wi-sh':
        #working on load summary
        eg = goodstkii(givendata,'dom')

        #naming worksheet
        worksh = xlworkbk.add_worksheet(sheetname)

        hewmdmst.append('R5:R6')
        chwmdmst.append('Shipper')
        hefmdwm.append(fmt_rheader)

        finmdmst.append('R7')
        cfnmdmst.append('(15)')
        fitfmdnm.append(fmt_rfiller)

        #writing titles, headers, and fillers
        dhflist = [[tiwmdmst, ctwmdmst, tifmdwm, None, None, None],
                   [hewmdmst, chwmdmst, hefmdwm, henmdmst, chnmdmst, hefmdnm],
                   [fiwmdmst, cfwmdmst, fitfmdwm, finmdmst, cfnmdmst, fitfmdnm]]

        for num in range(3):
            rowcols, contents, fmts = dhflist[num][0], dhflist[num][1], dhflist[num][2]
            for val in range(len(rowcols)):
                worksh.merge_range(rowcols[val],contents[val],fmts[val])

            rowcols, contents, fmts = dhflist[num][3], dhflist[num][4], dhflist[num][5]
            if num == 0:
                pass
            elif num != 0:
                for val in range(len(rowcols)):
                    worksh.write(rowcols[val],contents[val],fmts[val])

        #writing main data
        ini = 0
        for keys in givendata:
            for i in range(len(givendata[keys])):
                if keys == 'No':
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_lumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_ldmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_lmain)
                
                elif keys in ['Tgl Tiba', 'Tgl Tolak']:
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_udates)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_ddates)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mdates)

                elif keys in ['Nama Kapal','GT','Keagenan']:
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_ulefts)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_dlefts)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mlefts)

                elif keys == 'Shipper':
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_rumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_rdmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_rmain)

                else:
                    if i == 0:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mumain)
                    elif i == len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mdmain)
                    elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+7,ini,givendata[keys][i],fmt_mmain)

            ini += 1

        for col in range(len(givendata)):
            if col == 0:
                worksh.write(7+len(givendata['No']),col,None,fmt_lemain)
            elif col == len(givendata)-1:
                worksh.write(7+len(givendata['No']),col,None,fmt_remain)
            elif col > 0 and col < len(givendata)-1:
                worksh.write(7+len(givendata['No']),col,None,fmt_memain)

        worksh.write(7+len(givendata['No']),5,'=SUM(F8:F'+str(7+len(givendata['No']))+')',fmt_memain)

        cowidmst.append(170)

        #set column's width and row's height
        for num in range(len(cowidmst)):
            worksh.set_column_pixels(num,num,cowidmst[num])

        for num in range(len(rohedmst)):
            if num == len(rohedmst)-1:
                for i in range(len(givendata['No'])+1):
                    worksh.set_row_pixels(i+7,rohedmst[num])
            elif num != len(rohedmst)-1:
                worksh.set_row_pixels(num,rohedmst[num])

        #write loads summary
        #calculating row
        exmrhdroco = ['D'+str(11+len(givendata['GT']))+':I'+str(11+len(givendata['GT'])),
                      'E'+str(12+len(givendata['GT']))+':G'+str(12+len(givendata['GT'])),
                      'H'+str(12+len(givendata['GT']))+':I'+str(12+len(givendata['GT'])),
                      'D'+str(12+len(givendata['GT']))]
        excmhdroco = ['JUMLAH MUATAN KAPAL LUAR NEGERI','BONGKAR','MUAT','JENIS BARANG']
            
        for val in range(len(exmrhdroco)):
            if val != len(exmrhdroco)-1:
                worksh.merge_range(exmrhdroco[val],excmhdroco[val],fmt_hdnume)
            elif val == len(exmrhdroco)-1:
                worksh.write(exmrhdroco[val],excmhdroco[val],fmt_hdnume)
            
        for col in range(3):
            if col == 0:
                for row in range(len(eg[col])):
                    roco = 'D'+str(13+len(givendata['GT'])+row)
                    worksh.write(roco,eg[col][row],fmt_sumnam)
            elif col == 1:
                for row in range(len(eg[col])):
                    roco = 'E'+str(13+len(givendata['GT'])+row)+':G'+str(13+len(givendata['GT'])+row)
                    worksh.merge_range(roco,eg[col][row],fmt_hdnume)
            elif col == 2:
                for row in range(len(eg[col])):
                    roco = 'H'+str(13+len(givendata['GT'])+row)+':I'+str(13+len(givendata['GT'])+row)
                    worksh.merge_range(roco,eg[col][row],fmt_hdnume)

        #write signing place
        datetext = 'Pulau Bunyu, '+str(dt.strftime(dt.today(),'%d'))+' '+str(trnslt.translate(dt.strftime(dt.today(),'%B')))+' '+str(dt.strftime(dt.today(),'%Y'))
        signtext = [None,'Kepala Kantor','Unit Penyelenggara Pelabuhan','Kelas III Pulau Bunyu',
                    None,None,None,None,None,'Abdul Wahid','NIP. 19710515 199803 1 006']

        worksh.write(10+len(givendata[keys]),13,datetext,fmt_signdflt)
        worksh.write_column(10+len(givendata[keys])+1,13,signtext,fmt_signbold)

#4 Port Clearance style report
def clrnc_based(xlwriter,rwdata):
    #Preparing the data and workbook
    givendata = blankrows_clrnc(rwdata)
    xlworkbk = xlwriter.book
    
    #WORKSHEET FORMATS

    #title
    fmt_title_bold = xlworkbk.add_format({'bold':True,'font_name':'verdana','font_size':18,'align':'center','valign':'vcenter'})
    fmt_title_sub = xlworkbk.add_format({'bold':True,'font_name':'verdana','font_size':14,'align':'left','valign':'vcenter'})

    #header
    fmt_lheader = xlworkbk.add_format({'bold':True,'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_mheader = xlworkbk.add_format({'bold':True,'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':1,'align':'center','valign':'vcenter'})
    fmt_rheader = xlworkbk.add_format({'bold':True,'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':6,'left':1,'align':'center','valign':'vcenter'})    
    fmt_mshhead = xlworkbk.add_format({'bold':True,'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':1,'top':6,'right':1,'left':1,'align':'center','valign':'vcenter'})
    fmt_sheader = xlworkbk.add_format({'bold':True,'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #filler
    fmt_lfiller = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rfiller = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':6,'left':1,'align':'center','valign':'vcenter'})
    fmt_mfiller = xlworkbk.add_format({'font_name':'microsoft sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':6,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #main data upper part
    fmt_lumain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rumain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':6,'left':1,'align':'left','valign':'vcenter'})
    fmt_mumain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #main data middle part
    fmt_lmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':6,'left':1,'align':'left','valign':'vcenter'})
    fmt_mmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #main data downer part
    fmt_ldmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':6,'align':'center','valign':'vcenter'})
    fmt_rdmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':6,'left':1,'align':'left','valign':'vcenter'})
    fmt_mdmain = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter'})

    #main data code, name, and GT
    fmt_ulefts = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'left','valign':'vcenter'})
    fmt_mlefts = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'left','valign':'vcenter'})
    fmt_dlefts = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'left','valign':'vcenter'})

    #main data dates
    fmt_udates = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'dd/mm/yyyy'})
    fmt_mdates = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'dd/mm/yyyy'})
    fmt_ddates = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':6,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter','num_format':'dd/mm/yyyy'})
    
    #main data splitter
    fmt_lsplt = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':6,'align':'center','valign':'vcenter','bg_color':'orange'})
    fmt_rsplt = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':6,'left':1,'align':'center','valign':'vcenter','bg_color':'orange'})
    fmt_msplt = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':7,'top':7,'right':1,'left':1,'align':'center','valign':'vcenter','bg_color':'orange'})

    #summary format
    fmt_hdnume = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':1,'top':1,'right':1,'left':1,'align':'center','valign':'vcenter'})
    fmt_sumnam = xlworkbk.add_format({'font_name':'microsofot sans serif','font_size':12,'text_wrap':True,'bottom':1,'top':1,'right':1,'left':1,'align':'left','valign':'vcenter'})
    
    #signing format
    fmt_signbold = xlworkbk.add_format({'font_name':'arial','font_size':20,'bold':True})
    fmt_signdflt = xlworkbk.add_format({'font_name':'arial','font_size':20,'bold':False})

    #REQUIRED VARIABLES

    #month and year
    prevmonth = str(trnslt.translate(dt.strftime(dt.today().replace(day=1) - td(days=1),'%B')).upper())
    prevyear = dt.strftime(dt.today().replace(day=1) - td(days=1),'%Y')

    #list of titles, header's rowcol, contents, and format sequences
    tiwmclr,ctwmclr,tifmcwm = ['A1:U1','A2:U2'],['LAPORAN BULANAN PENGGUNAAN SPB (SURAT PERSETUJUAN BERLAYAR)',
                                 'KELUAR / MASUK KAPAL'],[fmt_title_bold,fmt_title_bold]

    tinmclr,ctnmclr,tifmcnm = ['A5','D5','A6','D6','U6'],['NAMA KANTOR',': KANTOR UNIT PENYELENGGARA PELABUHAN KELAS III P. BUNYU',
                                 'PROVINSI',': KALIMANTAN UTARA','BULAN: '+prevmonth+' '+prevyear],[fmt_title_sub,fmt_title_sub,
                                 fmt_title_sub,fmt_title_sub,fmt_title_sub]

    hewmclr = ['A8:A9','B8:C9','D8:D9','E8:G8','E9:F9','H8:H9','I8:I9','J8:J9','K8:K9',
                'L8:L9','M8:N8','O8:O9','P8:Q8','R8:T8','S9:T9','U8:U9']
    chwmclr = ['NO.','NO. SERI','NO. REG.','N A M A','KAPAL','BENDERA','GT','NO. SIPI',
                'NO. SIKPI','NO. SLO','T I B A','JUMLAH ABK','B E R A N G K A T','M U A T','JUMLAH',
                'PERUSAHAAN DAN / ATAU AGEN KAPAL']
    hefmcwm = [fmt_lheader,fmt_mheader,fmt_mheader,fmt_mshhead,fmt_sheader,fmt_mheader,
               fmt_mheader,fmt_mheader,fmt_mheader,fmt_mheader,fmt_mshhead,fmt_mheader,
               fmt_mshhead,fmt_mshhead,fmt_sheader,fmt_rheader]

    henmclr,chnmclr = ['G9','M9','N9','P9','Q9','R9'],['NAHKODA','DARI','TANGGAL','TUJUAN','TANGGAL','JENIS BARANG']
    hefmcnm = [fmt_sheader,fmt_sheader,fmt_sheader,fmt_sheader,fmt_sheader,fmt_sheader]

    fiwmclr,cfwmclr,fitfmcwm = ['B10:C10','E10:F10','S10:T10'],['(2)','(4)','(17)'],[fmt_mfiller,fmt_mfiller,fmt_mfiller]
    finmclr = ['A10','D10','G10','H10','I10','J10','K10','L10','M10','N10','O10','P10','Q10','R10','U10']
    cfnmclr = ['(1)','(3)','(5)','(6)','(7)','(8)','(9)','(10)','(11)','(12)','(13)','(14)','(15)','(16)','(18)']
    fitfmcnm = [fmt_lfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,
                fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_mfiller,fmt_rfiller]

    #list of column's width and row's height
    cowiclr = [70,50,125,65,70,370,445,250,145,75,75,75,255,138,70,295,135,480,160,145,465]
    roheclr = [33,33,33,33,33,33,33,35,45,25,25]

    #WRITING START

    #naming worksheet
    worksh = xlworkbk.add_worksheet('SPB')

    #writing titles, headers, and fillers
    chflist = [[tiwmclr, ctwmclr, tifmcwm, tinmclr, ctnmclr, tifmcnm],
               [hewmclr, chwmclr, hefmcwm, henmclr, chnmclr, hefmcnm],
               [fiwmclr, cfwmclr, fitfmcwm, finmclr, cfnmclr, fitfmcnm]]

    for num in range(3):
        rowcols, contents, fmts = chflist[num][0], chflist[num][1], chflist[num][2]
        for val in range(len(rowcols)):
            worksh.merge_range(rowcols[val],contents[val],fmts[val])
    
        rowcols, contents, fmts = chflist[num][3], chflist[num][4], chflist[num][5]
        for val in range(len(rowcols)):
            worksh.write(rowcols[val],contents[val],fmts[val])

    #writing main data
    ini = 0
    for keys in givendata:
        for i in range(len(givendata[keys])):
            if keys == 'No':
                if i == 0:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_lumain)
                elif i == len(givendata[keys])-1:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_ldmain)
                elif i != 0 and i != len(givendata[keys])-1:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_lmain)
                
            elif keys in ['Tgl Tiba', 'Tgl Tolak']:
                if i == 0:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_udates)
                elif i == len(givendata[keys])-1:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_ddates)
                elif i != 0 and i != len(givendata[keys])-1:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_mdates)

            elif keys in ['Nama Kapal','GT']:
                if i == 0:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_ulefts)
                elif i == len(givendata[keys])-1:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_dlefts)
                elif i != 0 and i != len(givendata[keys])-1:
                        worksh.write(i+10,ini,givendata[keys][i],fmt_mlefts)

            elif keys == 'Keagenan':
                if i == 0:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_rumain)
                elif i == len(givendata[keys])-1:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_rdmain)
                elif i != 0 and i != len(givendata[keys])-1:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_rmain)

            else:
                if i == 0:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_mumain)
                elif i == len(givendata[keys])-1:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_mdmain)
                elif i != 0 and i != len(givendata[keys])-1:
                    worksh.write(i+10,ini,givendata[keys][i],fmt_mmain)

        ini += 1

        #set column's width and row's height
        for num in range(len(cowiclr)):
            worksh.set_column_pixels(num,num,cowiclr[num])

        for num in range(len(roheclr)):
            if num == len(roheclr)-1:
                for i in range(len(givendata['No'])):
                    worksh.set_row_pixels(i+10,roheclr[num])
            elif num != len(roheclr)-1:
                worksh.set_row_pixels(num,roheclr[num])

        #write signing place
        datetext = 'Pulau Bunyu, '+str(dt.strftime(dt.today(),'%d'))+' '+str(trnslt.translate(dt.strftime(dt.today(),'%B')))+' '+str(dt.strftime(dt.today(),'%Y'))
        signtext = [None,'Kepala Kantor','Unit Penyelenggara Pelabuhan','Kelas III Pulau Bunyu',
                    None,None,None,None,None,'Abdul Wahid','NIP. 19710515 199803 1 006']

        worksh.write(10+len(givendata[keys])+3,18,datetext,fmt_signdflt)
        worksh.write_column(10+len(givendata[keys])+4,18,signtext,fmt_signbold)

#master function
def main():
    datadf = read_excel_file(filename='Tempat Input Kunjungan Kapal.xlsx')
    under500, above500, dom, exp, bun = categorizing(datadf)

    filewriter = pd.ExcelWriter('Untuk Hasil Olah Data.xlsx',engine='xlsxwriter')
    clrnc_based(filewriter,datadf)
    sib_based(under500,above500,filewriter)
    tkii_based(datadf,'prtmn',filewriter)
    tkii_based(datadf,'skkms',filewriter)
    tkii_based([dom,exp],'ptlim',filewriter)
    tkii_based(datadf,'mipsk',filewriter)
    tkii_based([dom,exp],'ptgtb',filewriter)
    tkii_based(datadf,'pttum',filewriter)
    tkii_based(datadf,'ptssp',filewriter)
    tkii_based(datadf,'klngn',filewriter)
    tkii_based(datadf,'ptser',filewriter)
    tkii_based(datadf,'mipmj',filewriter)
    tkii_based(datadf,'jobsi',filewriter)
    tkii_based(datadf,'tkiiupt',filewriter)
    dmstc_based(dom,'dmstc',filewriter)
    dmstc_based(exp,'ekspr',filewriter)
    tkii_based(bun,'butkii',filewriter)
    tkii_based(datadf,'tmtkii',filewriter)
    dmstc_based(dom,'sawit',filewriter)
    dmstc_based(dom,'batbar',filewriter)
    dmstc_based(dom,'gencar',filewriter)
    dmstc_based(dom,'batcah',filewriter)
    dmstc_based(dom,'cruoil',filewriter)
    dmstc_based(dom,'alaber',filewriter)
    dmstc_based(dom,'bebeem',filewriter)
    dmstc_based(dom,'kendrn',filewriter)
    dmstc_based(dom,'kaywoo',filewriter)
    dmstc_based(dom,'tansan',filewriter)
    dmstc_based(dom,'campur',filewriter)

    filewriter.save()
    
if __name__ == '__main__':
    main()

''' ------------------------------End of The Code Writing---------------------------'''
