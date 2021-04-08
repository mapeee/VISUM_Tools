# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        PuT-Maps in VISUM
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

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(path),'VISUM_Tools','Mapwork.txt')
f = path.read_text()
f = f.split('\n')

#--Path--#
XLSX = f[0]

#--Excel--#
XLSX = openpyxl.load_workbook(XLSX)
Stops = XLSX.active

#--VISUM--#
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.20")
dirTo = VISUM.Net.Directions.GetAll[0]
Read = VISUM.CreateNetReadRouteSearchTsys()
Read.SetAttValue("ChangeLinkTypeOfOpenedLinks",False)
Read.SetAttValue("IncludeBlockedTurns",True)
Read.SetAttValue("HowToHandleIncompleteRoute", 2) ##search shortest path
Read.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths",1)
Read.SetAttValue("ShortestPathCriterion",3) ##link length (1 = travel time; 3 = link length)
Read.SetAttValue("MaxDeviationFactor",3)
Read.SetAttValue("WhatToDoIfShortestPathNotFound",1) ##open turns
Read.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths",1)
Read.SetAttValue("WhatToDoIfStopPointIsBlocked", 2) ##Open the StopPoint
Read.SetAttValue("WhatToDoIfStopPointNotFound", 0) ##do not read lineroute

#--Build up Network--#
i = 1
for node in range(2,Stops.max_row+1):
    Name = Stops.cell(node,1).value
    X = Stops.cell(node,2).value
    Y = Stops.cell(node,3).value
    Nr = i
    print(Name)
    
    #Add
    Node = VISUM.Net.AddNode(Nr,X,Y)
    Node.SetAttValue("Name",Name)
    Stop = VISUM.Net.AddStop(Nr,X,Y)
    Stop.SetAttValue("Name",Name)
    StopArea = VISUM.Net.AddStopArea(Nr,Stop,Node)
    StopArea.SetAttValue("Name",Name)
    StopPoint = VISUM.Net.AddStopPointOnNode(Nr,StopArea,Node)
    StopPoint.SetAttValue("Name",Name)
 
    #Helper
    Node = VISUM.Net.AddNode(int(str(Nr)+"001"),X+1,Y) #east
    Node = VISUM.Net.AddNode(int(str(Nr)+"002"),X+1,Y-10) #southeast
    Node = VISUM.Net.AddNode(int(str(Nr)+"003"),X,Y-10) #south
    Node = VISUM.Net.AddNode(int(str(Nr)+"004"),X-1,Y-10) #southwest
    Node = VISUM.Net.AddNode(int(str(Nr)+"005"),X-1,Y) #west
    Node = VISUM.Net.AddNode(int(str(Nr)+"006"),X-1,Y+10) #northwest
    Node = VISUM.Net.AddNode(int(str(Nr)+"007"),X,Y+10) #north
    Node = VISUM.Net.AddNode(int(str(Nr)+"008"),X+1,Y+10) #northeast
    
    #Links
    VISUM.Net.AddLink(int(str(Nr)+"001"),int(str(Nr)+"001"),int(str(Nr)+"002"),1)
    VISUM.Net.AddLink(int(str(Nr)+"002"),int(str(Nr)+"002"),int(str(Nr)+"003"),1)
    VISUM.Net.AddLink(int(str(Nr)+"003"),int(str(Nr)+"003"),int(str(Nr)+"004"),1)
    VISUM.Net.AddLink(int(str(Nr)+"004"),int(str(Nr)+"004"),int(str(Nr)+"005"),1)
    VISUM.Net.AddLink(int(str(Nr)+"005"),int(str(Nr)+"005"),int(str(Nr)+"006"),1)
    VISUM.Net.AddLink(int(str(Nr)+"006"),int(str(Nr)+"006"),int(str(Nr)+"007"),1)
    VISUM.Net.AddLink(int(str(Nr)+"007"),int(str(Nr)+"007"),int(str(Nr)+"008"),1)
    VISUM.Net.AddLink(int(str(Nr)+"008"),int(str(Nr)+"008"),int(str(Nr)+"001"),1)
    
    VISUM.Net.AddLink(int(str(Nr)+"010"),int(str(Nr)+"005"),Nr,1)
    VISUM.Net.AddLink(int(str(Nr)+"011"),int(str(Nr)+"001"),Nr,1)
    
    if i > 1: VISUM.Net.AddLink(int(str(Nr)+"012"),int(str(Nr-1)+"001"),int(str(Nr)+"005"),1)
     
    i+=1

#--Build up Lines--#
Line = VISUM.Net.AddLine("Line1","B")

Stops = VISUM.CreateNetElements()
Stops.Add(VISUM.Net.StopPoints.ItemByKey(1))
Stops.Add(VISUM.Net.Nodes.ItemByKey(2004))
Stops.Add(VISUM.Net.Nodes.ItemByKey(2002))
Stops.Add(VISUM.Net.StopPoints.ItemByKey(3))
VISUM.Net.AddLineRoute("Route1",Line,dirTo,Stops,Read)

Stops = VISUM.CreateNetElements()
Stops.Add(VISUM.Net.StopPoints.ItemByKey(1))
Stops.Add(VISUM.Net.StopPoints.ItemByKey(2))
Stops.Add(VISUM.Net.StopPoints.ItemByKey(3))
VISUM.Net.AddLineRoute("Route2",Line,dirTo,Stops,Read)


#End
XLSX.close()