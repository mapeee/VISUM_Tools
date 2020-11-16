# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        splitting VISUM links
# Purpose:
#
# Author:      mape
#
# Created:     18/09/2014
# Copyright:   (c) mape 2014
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import win32com.client.dynamic
import xlrd
import time
start_time = time.time()
from datetime import datetime

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(path),'VISUM_Tools','Haltestellen.txt')
f = path.read_text()
f = f.split('\n')

#--Parameters--#
Netz = f[0]
XLS = f[1]

print("excel "+XLS)

#--outputs--#
date = datetime.now()
txt_name = 'StreckeTeilen_'+date.strftime('%m%d%Y')+'.txt'
f = open(Path.home() / 'Desktop' / txt_name,'w+')
f.write(time.ctime()+"\n")
f.write("Input Network: "+Netz+"\n")
f.write("Change Excel: "+XLS+"\n")

f.write("\n\n\n\n")
#--open VISUM--#
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.20")
VISUM.loadversion(Netz)
VISUM.Filters.InitAll()

#--open xlsx--#
Dokument = xlrd.open_workbook(XLS)
Blatt = Dokument.sheet_by_index(0) ##first sheet
Zeilen = Blatt.nrows-1 ##-1 due to headlines
print(" ",Zeilen,"rows!")

#--splitting--#
f.write("Splitting links \n")
for i in range(Zeilen):
    FromNode = int(Blatt.cell_value(rowx=1+i,colx=0)) ##1+i ignoring headlines
    ToNode = int(Blatt.cell_value(rowx=1+i,colx=1))
    NodeNr = int(Blatt.cell_value(rowx=1+i,colx=2))
    x_coord = Blatt.cell_value(rowx=1+i,colx=3)
    y_coord = Blatt.cell_value(rowx=1+i,colx=4)
    
    try:
        VISUM.Net.Links.SplitViaNode(FromNode,ToNode,NodeNr)
        f.write("node connected: "+str(NodeNr)+"\n")
    except:
        f.write("!!Error: node: "+str(NodeNr)+"\n")
        print("!!Error: node: "+str(NodeNr)+"\n")
        
    xynode = VISUM.Net.Nodes.ItemByKey(NodeNr)
    xynode.SetAttValue("XCoord",x_coord)
    xynode.SetAttValue("YCoord",y_coord)
    

f.write("\n\n\n")

##end
Sekunden = int(time.time() - start_time)
print("--finished after ",Sekunden,"seconds--")

f.write("--finished after "+str(Sekunden)+" seconds--")
f.close()