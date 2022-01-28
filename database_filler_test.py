# -*- coding: windows-1251 -*-
import csv
import time

import pandas as pd
import docx2txt
import numpy as np
import os
import re
from glob import glob
import re
import os
import win32com.client as win32
from win32com.client import constants


def filldb(path):
    # preparing separators
    blankspc = 'BBBBB'

    f = open("mainheadlist.txt", "r")
    lines = f.readlines()
    mheadlist = [i.replace('\n', '') for i in lines]

    sepmheadlist = ['$$$' + i for i in mheadlist]

    rawtext = docx2txt.process(path)
    rawtext = rawtext.replace('изменеия', 'изменения')
    rawtext = rawtext.replace('ВРАЧ', 'Врач')
    rawtext = rawtext.replace('  ', ' ')
    rawtext = rawtext.replace(',', '.')
    rawtext = rawtext.split('Врач')

    rawtext = [blankspc if i.find('ЭХОКАРДИОГРАФИЯ') == -1 else i for i in rawtext]

    rawtext = [x for x in rawtext if x != blankspc]

    # 1st level parsing

    rawdata = []

    #rawdata = [i.replace(mheadlist[j], sepmheadlist[j]) for j in range(len(mheadlist)) for i in rawtext]
    for i in rawtext:
        for j in range(len(mheadlist)):
            i = i.replace(mheadlist[j], sepmheadlist[j])
        rawdata.append(i)

    rawdata1 = [i.split('$$$') for i in rawdata]

    rawdata1 = [i[1:-1] + i[-1].split('Дополнительные данные') for i in rawdata1]

    # collecting numerical data
    data = []
    for i in range(len(rawdata1)):
        data.append([])
        for j in rawdata1[i]:
            for y in j.split():
                try:
                    data[i].append(float(y))
                except ValueError:
                    pass
            if len(data[i]) > 21:
                data[i] = data[i][:21]
    for i in range(len(data)):
        while len(data[i]) < 21:
            data[i].append(np.nan)

    # collecting verbal data

    f = open("verbdata1.txt", "r")
    verbsamples = f.readlines()
    verbmatch = [verbsamples[i].rstrip('\n') if i % 2 == 0 else None for i in range(len(verbsamples))]
    while verbmatch.__contains__(None):
        verbmatch.remove(None)
    verbmatch = [i.split(', ') for i in verbmatch]

    f = open("verbhead.txt", "r")
    keys = f.readlines()
    for i in range(len(keys)):
        keys[i] = keys[i].rstrip(' \n')

    verbdict = dict(zip(keys, verbmatch))
    print(verbdict)
    protodict = {}

    for k in range(len(rawdata1)):
        for j in rawdata1[k]:
            for key in verbdict.keys():
                if j.find(key) != -1:
                    a = key
                    b = j.split(key)[1]
                    protodict[a] = b
        print(protodict)
        p = True
        for key in protodict:
            for x in verbdict[key]:
                if x in protodict[key].lower():
                    data[k].append(x)
                    p = False
            if p:
                data[k].append('')
            p = True
    print(data)
    # collecting diagnoses

    f = open("diag.txt", "r")
    diag = f.readlines()

    diag = [i.rstrip(' \n') for i in diag]

    diagdata = []
    for i in range(len(rawdata)):
        diagdata.append([])
        for j in range(len(diag)):
            if rawdata[i].find(diag[j]) != -1:
                diagdata[i] = diag[j]
                break
            if j == len(diag)-1:
                diagdata[i] = 'Здоров'
                break

    diagdata = [str(i) for i in diagdata]

    for i in data:
        while len(i) != 53:
            i.append(np.nan)


    # creating dataset

    f = open("headerslist.txt", "r")
    lines = f.readlines()
    headerslist = [i.replace('\n', '') for i in lines]

    df = pd.DataFrame(columns=headerslist, index=range(len(data)))
    for i in range(len(data)):
        df.reset_index(inplace=True, drop=True)
        df.iloc[i] = data[i]
    df["Диагноз"] = diagdata

#    df = df.dropna(axis=1)

    print(df)
    print(df.columns)
    return df
start = time.time()
# processing docx


folderpath = r"D:\Ivan\CPP_1\db_med_test"
filepaths = [os.path.join(folderpath, name) for name in os.listdir(folderpath)]
dfpaths = [os.path.join(name) for name in os.listdir()]

k = 0
frames = []
for q in range(len(filepaths)):
    if str(filepaths[q]).endswith("docx"):
        df1 = filldb(filepaths[q])
        df1.to_csv('dfendver %d.csv' % k)
        k += 1
    # if str(filepaths[q]).endswith("doc"):
        # word = win32.gencache.EnsureDispatch('Word.Application')
        # doc = word.Documents.Open(filepaths[q])
        # doc.Activate()
        #
        # # Rename path with .docx
        # new_file_abs = os.path.abspath(filepaths[q])
        # new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
        #
        # # Save and Close
        # word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
        # doc.Close(False)
        # df1 = filldb(filepaths[q] + 'x')
        # df1.to_csv('dfendver %d.csv' % k)
        # k += 1

dfpaths = [os.path.join(name) for name in os.listdir()]
print(dfpaths)
df = pd.read_csv('dfendver 0.csv')

for q in range(1, len(dfpaths)):
    if re.search(r'\bdfendver\b\s\d', str(dfpaths[q]).split('.')[0]):
        dfbuf = pd.read_csv(dfpaths[q])
        df = df.append(dfbuf)
        df.to_csv('dfendver.csv')
        os.remove(dfpaths[q])

db = pd.read_csv('dfendver.csv')
db = db.T.drop_duplicates().T
db = db.drop(labels='Unnamed: 0', axis='columns')
db.to_csv('dfendver.csv')

end = time.time()
print(end - start)
