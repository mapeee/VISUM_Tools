# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        Calculate shortest path from excel
# Purpose:
#
# Author:      mape
#
# Created:     23/07/2021
# Copyright:   (c) mape 2021
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import win32com.client.dynamic
from openpyxl import load_workbook
import time
start_time = time.time()

from pathlib import Path
f = open(Path.home() / 'python32' / 'python_dir.txt', mode='r')
for i in f: path = i
path = Path.joinpath(Path(path),'VISUM_Tools','TimeProfiles.txt')
f = path.read_text().split('\n')

#--Path--#
Network = f[0]
Excel = f[1]

def VISUM_open(Net):
    VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.20")
    VISUM.loadversion(Net)
    VISUM.Filters.InitAll()
    return VISUM   

Excel = r''
Network = r''

#Excel
wb = load_workbook(filename = Excel)
ws = wb.worksheets[0]

#VISUM
VISUM = VISUM_open(Network)
ShortPath = VISUM.Analysis.RouteSearchPrT

for row in ws.iter_rows(min_row=2): #skip header
    NE = VISUM.CreateNetElements()
    Node = VISUM.Net.Nodes.ItemByKey(row[0].value)
    NE.Add(Node)
    Node = VISUM.Net.Nodes.ItemByKey(row[1].value)
    NE.Add(Node)
    
    ShortPath.Execute(NE,"C",1)
    row[2].value = int(ShortPath.AttValue("tCur"))
    ShortPath.Clear()
    
    print(row[0].row)

#END
wb.save(Excel)
wb.close()
seconds = int(time.time() - start_time)
print("--finished after ",seconds,"seconds--") 