# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        importing lines for new routes
# Purpose:
# Author:      mape
# Created:     05/01/2021
# Copyright:   (c) mape 2021
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import win32com.client.dynamic
import pyodbc
from send2trash import send2trash

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
with open(path, mode='r') as f: path = f.readlines()
path = path[0] +"\\"+'VISUM_Tools'+"\\"+'Line_import.txt'
with open(path, mode='r') as path: f = path.read().splitlines()

#--Path--#
Network = f[0]
layout_path = f[1]
access_db = f[2].replace("accdb","mdb")


def VISUM_open(Net):
    VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.22")
    VISUM.loadversion(Net)
    VISUM.Filters.InitAll()
    return VISUM

def VISUM_filter(VISUM):
    VISUM.Filters.InitAll()
    Lines = VISUM.Filters.LineGroupFilter()
    Lines.UseFilterForLineRoutes = True
    Lines.UseFilterForLineRouteItems = True
    Lines.UseFilterForTimeProfiles = True
    Lines.UseFilterForTimeProfileItems = True
    Lines.UseFilterForVehJourneys = True
    Lines.UseFilterForVehJourneySections = True
    Lines.UseFilterForVehJourneyItems = True
    Lines = Lines.LineRouteFilter()
    Lines.AddCondition("OP_NONE",False,"AddVal1","EqualVal",1)##no more filter,not complementary
    
    HST = VISUM.Filters.StopGroupFilter()
    HST.UseFilterForStopAreas = True
    HST.UseFilterForStopPoints = True
    HST = HST.StopPointFilter()
    HST.AddCondition("OP_NONE",False,r"Node\AddVal1","EqualVal",1)
    
    Connector = VISUM.Filters.ConnectorFilter()
    Connector.AddCondition("OP_NONE",False,r"Node\AddVal1","EqualVal",1)
    
    Nodes = VISUM.Filters.NodeFilter()
    Nodes.AddCondition("OP_NONE",False,"AddVal1","EqualVal",1)
    
    InsertedLinks = VISUM.Filters.LinkFilter()
    InsertedLinks.AddCondition("OP_NONE",False,"TYPENO", "ContainedIn", str(1))
    
def VISUM_export(VISUM,layout,access, Stops):
    global journeys_before
    global servingstops_before
    journeys_before = VISUM.Net.VehicleJourneys.Count ##number of journeys before export
    servingstops_before = VISUM.Net.StopPoints.GetMultipleAttributes(["Count:ServingVehJourneys"])
    servingstops_before = int(sum(list(map(sum, list(servingstops_before)))))
    
    VISUM.IO.SaveAccessDatabase(access,layout,True,False,True)
    VISUM.Net.LineRoutes.RemoveAll()
    VISUM.Net.Connectors.RemoveAll()
    VISUM.Net.StopAreas.RemoveAll()
    VISUM.Net.StopPoints.RemoveAll()

    ## Access
    conn = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="+access+";")
    cursor = conn.cursor()
    
    cursor.execute('''
                    UPDATE LINEROUTEITEM
                    INNER JOIN LINEROUTE
                    ON LINEROUTEITEM.LINEROUTENAME = LINEROUTE.NAME
                    SET LINEROUTEITEM.ADDVAL = LINEROUTE.ADDVAL1
                    ''')
    conn.commit()       
    cursor.execute('''
                    DELETE FROM LINEROUTEITEM 
                    WHERE ISROUTEPOINT = 0 or ISROUTEPOINT is NULL
                    ''')
    conn.commit()

    if Stops == False: return
    for i in Stops:
        cursor.execute('''
                        UPDATE LINEROUTEITEM
                        SET NODENO = '''+str(i[1])+''', STOPPOINTNO = '''+str(i[1])+'''
                        WHERE STOPPOINTNO = '''+str(i[0])+''' and ADDVAL = '''+str(1)+'''
                        ''')              
        conn.commit()
    conn.close()
 
def VISUM_import(VISUM,access,LinkType,shortcrit,open_blocked):
    VISUM_filter(VISUM)
    import_setting = VISUM.CreateNetReadRouteSearchTsys()
    import_setting.SetAttValue("ChangeLinkTypeOfOpenedLinks",open_blocked)
    import_setting.SetAttValue("IncludeBlockedTurns",open_blocked)
    import_setting.SetAttValue("HowToHandleIncompleteRoute", 2) ##search shortest path
    import_setting.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths",1)
    import_setting.SetAttValue("ShortestPathCriterion",shortcrit) ##1 = travel time; 3 = link length
    import_setting.SetAttValue("MaxDeviationFactor",50)
    import_setting.SetAttValue("WhatToDoIfShortestPathNotFound",2) ##insert link if necessary
    import_setting.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths",LinkType)
    import_setting.SetAttValue("WhatToDoIfStopPointIsBlocked", 2) ##Open the StopPoint
    import_setting.SetAttValue("WhatToDoIfStopPointNotFound", 0) ##do not read lineroute
    PuT_import = VISUM.CreateNetReadRouteSearch()
    PuT_import.SetForAllTSys(import_setting)
    VISUM.IO.LoadAccessDatabase(access,True,PuT_import)
    
    VISUM.Filters.NodeFilter().Init()
    Nodes = VISUM.Filters.NodeFilter()
    Nodes.AddCondition("OP_NONE",False,"AddVal1","GreaterVal",0)
    for N in VISUM.Net.Nodes.GetAllActive:
        N.SetAttValue("Name","")
        N.SetAttValue("AddVal1",0)
    
    VISUM.Filters.NodeFilter().Init()
    VISUM.Filters.StopGroupFilter().Init()
    VISUM.Filters.ConnectorFilter().Init()

    #--testing after the import--#
    InsertedLinks = VISUM.Filters.LinkFilter().Init() 
    InsertedLinks = VISUM.Filters.LinkFilter()
    InsertedLinks.AddCondition("OP_NONE",False,"TYPENO", "ContainedIn", str(LinkType))
    if VISUM.Net.Links.CountActive > 0: print("--"+str(VISUM.Net.Links.CountActive)+" new links added--") 
    VISUM.Filters.LinkFilter().Init()
    
    global journeys_after
    journeys_after = VISUM.Net.VehicleJourneys.Count ##number of journeys after processing
    if journeys_before != journeys_after:
        print("> missing VehicleJourneys")
        print("> before: "+str(journeys_before))
        print("> after: "+str(journeys_after)+"\n")

    global servingstops_after
    servingstops_after = VISUM.Net.StopPoints.GetMultipleAttributes(["Count:ServingVehJourneys"])
    servingstops_after = int(sum(list(map(sum, list(servingstops_after)))))
    if servingstops_before != servingstops_after:
        print("> missing servings at stops")
        print("> before: "+str(servingstops_before))
        print("> after: "+str(servingstops_after)+"\n")

def VISUM_end(VISUM,access):
    send2trash(access)
    for Route in VISUM.Net.LineRoutes.GetAllActive: Route.SetAttValue("AddVal1",0)
    VISUM.Filters.InitAll()
    VISUM_filter(VISUM)

#--processing--#
V = VISUM_open(Net=Network)
VISUM_filter(VISUM=V)

Stop = False #False = no change of served stops
# Stop = [[11018,11039]]   #old, new
# Stop = [[44304,44301],[44307,44302]]   #old, new

VISUM_export(VISUM=V, layout=layout_path, access=access_db, Stops=Stop)
VISUM_import(VISUM=V, access=access_db, LinkType=1, shortcrit=1, open_blocked=False) #shortcrit( 1 = travel time; 3 = link length)
VISUM_end(VISUM=V, access=access_db)

##end
V.SaveVersion(Network)