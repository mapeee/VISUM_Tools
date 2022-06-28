# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        importing lines for new routes
# Purpose:
#
# Author:      mape
#
# Created:     05/01/2021
# Copyright:   (c) mape 2021
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import win32com.client.dynamic
import pyodbc
import time
start_time = time.time()
from send2trash import send2trash

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
with open(path, mode='r') as f: path = f.readlines()
path = path[0] +"\\"+'VISUM_Tools'+"\\"+'Line_import.txt'
with open(path, mode='r') as path: f = path.read().splitlines()

#--Path--#
Network = f[0]
layout_path = f[1]
access_db = f[2]
access_db = access_db.replace("accdb","mdb")


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
    Lines.AddCondition("OP_NONE",False,"AddVal1","GreaterVal",0)##no more filter,not complementary
    
    HST = VISUM.Filters.StopGroupFilter()
    HST.UseFilterForStopAreas = True
    HST.UseFilterForStopPoints = True
    HST = HST.StopPointFilter()
    HST.AddCondition("OP_NONE",False,r"Node\AddVal1","GreaterVal",0)
    
    Connector = VISUM.Filters.ConnectorFilter()
    Connector.AddCondition("OP_NONE",False,r"Node\AddVal1","GreaterVal",0)
    
    Nodes = VISUM.Filters.NodeFilter()
    Nodes.AddCondition("OP_NONE",False,"AddVal1","GreaterVal",0)
    
    InsertedLinks = VISUM.Filters.LinkFilter()
    InsertedLinks.AddCondition("OP_NONE",False,"TYPENO", "ContainedIn", str(1))
    

def VISUM_export(VISUM,layout,access):
    VISUM.IO.SaveAccessDatabase(access,layout,True,False,True)
    
    global journeys_b
    journeys_b = VISUM.Net.VehicleJourneys.Count ##number of journeys before export
    
    VISUM.Net.LineRoutes.RemoveAll()
    VISUM.Net.Connectors.RemoveAll()
    VISUM.Net.StopAreas.RemoveAll()
    VISUM.Net.StopPoints.RemoveAll()


def access_edit(access,Nodes,change):
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

    if change == False: return
    for AddValues in Nodes:
        for i in AddValues[1:]:
            cursor.execute('''
                            UPDATE LINEROUTEITEM
                            SET NODENO = '''+str(i[1])+''', STOPPOINTNO = '''+str(i[1])+'''
                            WHERE STOPPOINTNO = '''+str(i[0])+''' and ADDVAL = '''+str(AddValues[0])+'''
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
    for N in VISUM.Net.Nodes.GetAllActive: N.SetAttValue("AddVal1",0)
    
    VISUM.Filters.NodeFilter().Init()
    VISUM.Filters.StopGroupFilter().Init()
    VISUM.Filters.ConnectorFilter().Init()
    
    global journeys_a
    journeys_a = VISUM.Net.VehicleJourneys.Count ##number of journeys after processing

    #--testing after the import--#
    InsertedLinks = VISUM.Filters.LinkFilter().Init() 
    InsertedLinks = VISUM.Filters.LinkFilter()
    InsertedLinks.AddCondition("OP_NONE",False,"TYPENO", "ContainedIn", str(LinkType))
    if VISUM.Net.Links.CountActive > 0: print("--"+str(VISUM.Net.Links.CountActive)+" new links added--") 
    VISUM.Filters.LinkFilter().Init()   
    if journeys_b != journeys_a: print("missing VehicleJourneys!!!")


#--processing--#
V = VISUM_open(Net=Network)
VISUM_filter(VISUM=V)
Node = [[1,[32129,7565091]]]   #old, new
# Node = [[1,[44304,44301]],[1,[44307,44302]]]   #old, new

VISUM_export(VISUM=V, layout=layout_path, access=access_db)
access_edit(access=access_db, Nodes=Node, change=False)
##False = no editing of nodenumbers

VISUM_import(VISUM=V, access=access_db, LinkType=1, shortcrit=1, open_blocked=False)
##shortcrit( 1 = travel time; 3 = link length)

##end
send2trash(access_db)
for Route in V.Net.LineRoutes.GetAllActive: Route.SetAttValue("AddVal1",0)
V.Filters.InitAll()
VISUM_filter(V)
print("--fertig--")

V.SaveVersion(Network)

# del VISUM
Sekunden = int(time.time() - start_time)
print("--finished after ",Sekunden,"seconds--")
