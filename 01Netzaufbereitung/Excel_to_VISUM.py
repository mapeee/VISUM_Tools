# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        Networkelements from Excel to Visum
# Purpose:
#
# Author:      mape
#
# Created:     08/04/2021
# Copyright:   (c) mape 2021
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import win32com.client.dynamic
import openpyxl
import time
start_time = time.time()

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(path),'VISUM_Tools','Excel_to_VISUM.txt')
f = path.read_text()
f = f.split('\n')

def Excel(path):
    XLSX = path
    XLSX = openpyxl.load_workbook(XLSX)
    return XLSX

def VISUM_open(Net):
    VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.20")
    VISUM.loadversion(Net)
    VISUM.Filters.InitAll()
    return VISUM

def Stops(VISUM,Excel):
    for row in range(2,Elements.max_row+1):
        Nr = Excel.cell(row,1).value
        Name = Excel.cell(row,2).value
        Node = VISUM.Net.Nodes.ItemByKey(Nr)
        X = Node.AttValue("XCoord")
        Y = Node.AttValue("YCoord")
        #Add Stops
        try:
            Node.SetAttValue("Name",Name)
            Stop = VISUM.Net.AddStop(Nr,X,Y)
            Stop.SetAttValue("Name",Name)
            StopArea = VISUM.Net.AddStopArea(Nr,Stop,Node)
            StopArea.SetAttValue("XCoord",X)
            StopArea.SetAttValue("YCoord",Y)
            StopArea.SetAttValue("Name",Name)
            StopPoint = VISUM.Net.AddStopPointOnNode(Nr,StopArea,Node)
            StopPoint.SetAttValue("Name",Name)
            StopPoint.SetAttValue("AddVal2",1)
            print("added: "+Name)
        except: print("error at: "+(str(Nr)))
        

#--Parameter--#
Network = f[0]
XLSX = Excel(f[1])
Elements = XLSX.active


#--Processing--#
VISUM = VISUM_open(Network)
#--Stops--#
Stops(VISUM,Elements)


#End
XLSX.close()
seconds = int(time.time() - start_time)
print("--finished after ",seconds,"seconds--")