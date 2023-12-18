# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        PuT from Excel to PTV VISUM
# Author:      mape
# Created:     08/04/2021 // 15/05/2023
# Copyright:   (c) mape 2021
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import win32com.client.dynamic
import openpyxl

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
path = open(path, mode='r').readlines()
path = path[0] +"\\"+'VISUM_Tools'+"\\"+'Excel_to_PuT.txt'
f = open(path, mode='r').read().splitlines()
Network = f[0]
XLSX = f[1]


def CreateHeadway(Headway, Dep, Dep_next, No, NoVJ):
    HeadwayJourneys = int(((Dep_next-Dep)/Headway)-1)
    for HeadwayJourney in range(HeadwayJourneys):
        NoHead = int(str(No)+str(HeadwayJourney+1))
        DepHead = Dep + (Headway*(HeadwayJourney+1))
        CreateVehJourney(VISUM, TimeProfile, NoHead, NoVJ, DepHead)

def CreateLine(VISUM, XLSX):
    LineName = XLSX.cell(1,2).value
    try:
        Line = VISUM.Net.Lines.ItemByKey(LineName)
        print("> Line "+LineName+ " already exists")
        return Line
    except:
        print("-> creating Line "+LineName)
        Line = VISUM.Net.AddLine(LineName,XLSX.cell(3,2).value)
        Line.SetAttValue("LINIENNUMMER",XLSX.cell(2,2).value)
        Line.SetAttValue("OPERATORNO",XLSX.cell(4,2).value)
        Line.SetAttValue("FARESYSTEMSET",XLSX.cell(5,2).value)
        return Line

def CreateLineRoute(VISUM, Line, No, Stops, Direct):
    RouteName = str(No)+"_"+Direct
    RouteSearch = VISUM_para(VISUM)
    Route = VISUM.Net.AddLineRoute(RouteName,Line,Direct,Stops,RouteSearch)
    Route.SetAttValue("Szenario","Script")
    return Route

def CreateTimeProfile(VISUM, Route, Times):
    TProfile = VISUM.Net.AddTimeProfile("1",Route)
    for Item in TProfile.TimeProfileItems:
        Index = Item.AttValue("Index")
        prevarr = Item.AttValue("PREVTIMEPROFILEITEM\ARR")
        if Index==1:continue
        arr = prevarr+Times[int(Index-2)]
        Item.SetAttValue("arr",arr)
    return TProfile    

def CreateVehJourney(VISUM, TimeProfile, No, NoVJ, Dep):
    VJ = VISUM.Net.AddVehicleJourney(NoVJ,TimeProfile)
    VJ.SetAttValue("Dep",Dep)
    VJ.SetAttValue("Name",str(No))
    Operator = VJ.AttValue('LINEROUTE\LINE\OPERATORNO')
    VJ.SetAttValue('OPERATORNO',Operator)
    return VJ

def DepartureTime(XLSX, Col):
    for row in range(10,XLSX.max_row+10):
        Time = XLSX.cell(row,Col).value
        if Time in [None,"|","S"]:continue
        Departure = Times_to_seconds(Time,None)
        break
    return Departure

def Excel(path):
    XLSX = path
    XLSX = openpyxl.load_workbook(XLSX)
    return XLSX

def Filter(VISUM):
    VISUM.Filters.InitAll()
    Lines = VISUM.Filters.LineGroupFilter()
    Lines.UseFilterForLineRoutes = True
    Lines.UseFilterForLines = True
    LineRouteF = Lines.LineRouteFilter()
    LineRouteF.AddCondition("OP_NONE",False,"Szenario",9,"Script")

def HeadwayCheck(Check):
    if Check is None:return None
    if "Min" not in str(Check) and "min" not in str(Check):return None
    freq = int(Check[:2])*60
    return freq  

def Stop_Container(VISUM, XLSX, Col):
    Stops = VISUM.CreateNetElements()
    Times = []
    Time_from = None
    Time_to = None
    for row in range(10,XLSX.max_row+10):
        Stop = XLSX.cell(row,Col).value
        if Stop in [None,"|","S"]:continue
        Stops.Add(VISUM.Net.StopPoints.ItemByKey(XLSX.cell(row,2).value))
        Time_to = XLSX.cell(row,Col).value
        seconds = Times_to_seconds(Time_from, Time_to)
        if seconds != None:Times.append(seconds)
        Time_from = XLSX.cell(row,Col).value
    return Stops, Times

def Times_to_seconds(Time_from, Time_to):
    if Time_from == None:
        return None
    if Time_to == None:
        departure = (int(str(Time_from)[-2:])+(int(str(Time_from)[:-2])*60))*60
        return departure
    minutes_from = int(str(Time_from)[-2:])+(int(str(Time_from)[:-2])*60)
    minutes_to = int(str(Time_to)[-2:])+(int(str(Time_to)[:-2])*60)
    seconds = (minutes_to - minutes_from)*60
    return seconds

def VISUM_open(Net):
    VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.22")
    VISUM.loadversion(Net)
    VISUM.Filters.InitAll()
    return VISUM

def VISUM_para(VISUM):
    Routing = VISUM.CreateNetReadRouteSearchTsys()
    Routing.SetAttValue("ChangeLinkTypeOfOpenedLinks",False)
    Routing.SetAttValue("IncludeBlockedTurns",False)
    Routing.SetAttValue("HowToHandleIncompleteRoute", 2) ##search shortest path
    Routing.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths",1)
    Routing.SetAttValue("ShortestPathCriterion",1) ##link length (1 = travel time; 3 = link length)
    Routing.SetAttValue("MaxDeviationFactor",8)
    Routing.SetAttValue("WhatToDoIfShortestPathNotFound",2) ##insert link
    Routing.SetAttValue("WhatToDoIfStopPointIsBlocked", 0) ##do not read lineroute
    Routing.SetAttValue("WhatToDoIfStopPointNotFound", 0) ##do not read lineroute
    return Routing

#--Parameter--#
VISUM = VISUM_open(Network)
VHNo = VISUM.Net.VehicleJourneys.GetAll
VHNoMax = VHNo[len(VHNo)-1].AttValue("No")
VHNo = VHNoMax+1
Direct = "R"
Elements = Excel(XLSX)[Direct]
Line = CreateLine(VISUM,Elements)
No=Dep=Stops=Times=LineRoute=TimeProfile=VehJourney = None

for Journey in range(3,Elements.max_column+1):
    No = Elements.cell(9,Journey).value
    print("> Journey: "+str(No))
    if HeadwayCheck(Elements.cell(10,Journey).value) != None:
        print("> Headway Journey: "+str(No))
        Headway = HeadwayCheck(Elements.cell(10,Journey).value)
        Dep_next = DepartureTime(Elements,Journey+1)
        CreateHeadway(Headway, Dep, Dep_next, No, VHNo)
        VHNo+=1
        continue          
    Dep = DepartureTime(Elements,Journey)
    Stops, Times = Stop_Container(VISUM, Elements, Journey)
    LineRoute = CreateLineRoute(VISUM, Line, No, Stops, Direct)
    TimeProfile = CreateTimeProfile(VISUM, LineRoute, Times)
    VehJourney = CreateVehJourney(VISUM, TimeProfile, No, VHNo, Dep)
    VHNo+=1
    
Filter(VISUM)

#End
# VISUM.SaveVersion(Network)
print("> finished")