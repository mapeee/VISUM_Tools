# -*- coding: iso-8859-1 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      mape
#
# Created:     06/02/2017
# Copyright:   (c) mape 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(r'C:'+path),'VISUM_Tools','Fake_Matrizen.txt')
f = path.read_text()
f = f.split('\n')


import xlrd
import win32com.client.dynamic

XLS = 'C:'+f[0]
Netz = 'C:'+f[1]

VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.15")
VISUM.loadversion(Netz)

Dokument = xlrd.open_workbook(XLS)
Blatt = Dokument.sheet_by_index(0)
Zeilen = Blatt.nrows-1

for i in range(Zeilen):
    Nr = int(Blatt.cell_value(rowx=1+i,colx=0))
    Name = Blatt.cell_value(rowx=1+i,colx=1)

    VISUM.Net.AddStop(Nr)
    VISUM.Net.AddStopArea(Nr,Nr,0)
    VISUM.Net.StopAreas.ItemByKey(Nr).SetAttValue("Name",Name)

print ("fertig!")