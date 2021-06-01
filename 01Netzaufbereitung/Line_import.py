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
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(path),'VISUM_Tools','Line_import.txt')
f = path.read_text()
f = f.split('\n')


#--Path--#
Network = f[0]
layout = f[1]
access_db = f[2]


def VISUM_open(Net):
    VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.20")
    VISUM.loadversion(Net)
    VISUM.Filters.InitAll()
    return VISUM

def VISUM_filter(VISUM):
    VISUM.Filters.InitAll()
    Linien = VISUM.Filters.LineGroupFilter()
    Linien.UseFilterForLineRoutes = True
    Linien.UseFilterForLineRouteItems = True
    Linien.UseFilterForTimeProfiles = True
    Linien.UseFilterForTimeProfileItems = True
    Linien.UseFilterForVehJourneys = True
    Linien.UseFilterForVehJourneySections = True
    Linien.UseFilterForVehJourneyItems = True
    Linien = Linien.LineRouteFilter()
    Linien.AddCondition("OP_NONE",False,"AddVal1","GreaterVal",0)##no more filter,not complementary,AddVal1 > 0
    
    HST = VISUM.Filters.StopGroupFilter()
    HST.UseFilterForStopAreas = True
    HST.UseFilterForStopPoints = True
    HST = HST.StopPointFilter()
    HST.AddCondition("OP_NONE",False,"Node\AddVal1","GreaterVal",0)
    
    Connector = VISUM.Filters.ConnectorFilter()
    Connector.AddCondition("OP_NONE",False,"Node\AddVal1","GreaterVal",0)
    
    Node = VISUM.Filters.NodeFilter()
    Node.AddCondition("OP_NONE",False,"AddVal1","GreaterVal",0)
    
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
 
def VISUM_import(VISUM,access,LinkType,journeys_b,shortcrit):
    VISUM_filter(VISUM)
    import_setting = VISUM.CreateNetReadRouteSearchTsys()
    import_setting.SetAttValue("ChangeLinkTypeOfOpenedLinks",False)
    import_setting.SetAttValue("IncludeBlockedTurns",False)
    import_setting.SetAttValue("HowToHandleIncompleteRoute", 2) ##search shortest path
    import_setting.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths",1)
    import_setting.SetAttValue("ShortestPathCriterion",shortcrit) ##link length (1 = travel time; 3 = link length)
    import_setting.SetAttValue("MaxDeviationFactor",50)
    import_setting.SetAttValue("WhatToDoIfShortestPathNotFound",2) ##insert link if necessary
    import_setting.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths",LinkType)
    import_setting.SetAttValue("WhatToDoIfStopPointIsBlocked", 2) ##Open the StopPoint
    import_setting.SetAttValue("WhatToDoIfStopPointNotFound", 0) ##do not read lineroute
    PuT_import = VISUM.CreateNetReadRouteSearch()
    PuT_import.SetForAllTSys(import_setting)
    VISUM.IO.LoadAccessDatabase(access,True,PuT_import)
    
    VISUM.Filters.NodeFilter().Init()
    Node = VISUM.Filters.NodeFilter()
    Node.AddCondition("OP_NONE",False,"AddVal1","GreaterVal",0)
    for Node in VISUM.Net.Nodes.GetAllActive: Node.SetAttValue("AddVal1",0)
    
    VISUM.Filters.NodeFilter().Init()
    VISUM.Filters.StopGroupFilter().Init()
    VISUM.Filters.ConnectorFilter().Init()
    
    global journeys_a
    journeys_a = VISUM.Net.VehicleJourneys.Count ##number of journeys after processing

    #--testing after the import--#
    InsertedLinks = VISUM.Filters.LinkFilter().Init() 
    InsertedLinks = VISUM.Filters.LinkFilter()
    InsertedLinks.AddCondition("OP_NONE",False,"TYPENO", "ContainedIn", str(insert_type))
    if VISUM.Net.Links.CountActive > 0: print("--"+str(VISUM.Net.Links.CountActive)+" new links added--") 
    VISUM.Filters.LinkFilter().Init()   
    if journeys_b != journeys_a: print("missing VehicleJourneys!!!")


#--Parameter--#
insert_type = 1 ##type of inserted links
# Nodes = [[1,[60013,640097209]]]   #old, new
    

#--processing--#
VISUM = VISUM_open(Network)
VISUM_filter(VISUM)
Nodes = [[1,[11012,3784292]]]   #old, new
# Nodes = [[1,[61070,3406056]],[1,[610701,3520689]]]   #old, new

VISUM_export(VISUM,layout,access_db)
access_edit(access_db,Nodes,False) ##False = no editing of nodenumbers

VISUM_import(VISUM,access_db,insert_type,journeys_b,1) ##(shortcrit: 1 = travel time; 3 = link length)

##end
send2trash(access_db)
for Route in VISUM.Net.LineRoutes.GetAllActive: Route.SetAttValue("AddVal1",0)
VISUM.Filters.InitAll()
VISUM_filter(VISUM)
print("--fertig--")

VISUM.SaveVersion(Network)

# del VISUM
Sekunden = int(time.time() - start_time)
print("--finished after ",Sekunden,"seconds--")