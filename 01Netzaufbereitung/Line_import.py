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

#--Parameter--#
insert_type = 1 ##type of inserted links
Nodes = [[1,[1,33],[2,44]],
          [2,[4,55]]]


def VISUM_open(Net):
    VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.20")
    VISUM.loadversion(Net)
    VISUM.Filters.InitAll()
    return VISUM

def VISUM_filter(VISUM):
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

def VISUM_export(VISUM,layout,access):
    VISUM.IO.SaveAccessDatabase(access,layout,True,False,True)
    VISUM.Net.LineRoutes.RemoveAll()

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
                            WHERE NODENO = '''+str(i[0])+''' and ADDVAL = '''+str(AddValues[0])+'''
                            ''')              
            conn.commit()
    conn.close()
 
def VISUM_import(VISUM,access,LinkType):
    import_setting = VISUM.CreateNetReadRouteSearchTsys()
    import_setting.SetAttValue("ChangeLinkTypeOfOpenedLinks",False)
    import_setting.SetAttValue("IncludeBlockedTurns",False)
    import_setting.SetAttValue("HowToHandleIncompleteRoute", 2) ##search shortest path
    import_setting.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths",1)
    import_setting.SetAttValue("ShortestPathCriterion",0) ##direct distance (1 = travel time)
    import_setting.SetAttValue("MaxDeviationFactor",10)
    import_setting.SetAttValue("WhatToDoIfShortestPathNotFound",2) ##insert link if necessary
    import_setting.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths",LinkType)
    import_setting.SetAttValue("WhatToDoIfStopPointIsBlocked", 0) ##do not insert time profile
    import_setting.SetAttValue("WhatToDoIfStopPointNotFound", 0) ##do not read lineroute
    PuT_import = VISUM.CreateNetReadRouteSearch()
    PuT_import.SetForAllTSys(import_setting)
    VISUM.IO.LoadAccessDatabase(access,True,PuT_import)   

def LinksTest(VISUM,filter2):
    InsertedLinks = VISUM.Filters.LinkFilter()
    InsertedLinks.UseFilter = True
    InsertedLinks.AddCondition("OP_NONE",False,"TYPENO", "ContainedIn", str(insert_type))
    if filter2 == True: InsertedLinks.AddCondition("OP_AND",False,"DISPLAYTYPE", "EqualVal", "")
    InsertedLinks = VISUM.Net.Links.CountActive
    if InsertedLinks > 0: print("--"+str(InsertedLinks)+" new links added--")
    
    
#--processing--#
VISUM = VISUM_open(Network)
journeys_b = VISUM.Net.VehicleJourneys.Count ##number of journeys before processing

VISUM_filter(VISUM)
VISUM_export(VISUM,layout,access_db)
access_edit(access_db,Nodes,False) ##False = no changing of nodenumbers
VISUM_import(VISUM,access_db,insert_type)
journeys_a = VISUM.Net.VehicleJourneys.Count ##number of journeys after processing

for Route in VISUM.Net.LineRoutes.GetAllActive: Route.SetAttValue("AddVal1",0)


#--testing--#
if journeys_b == journeys_a: pass
else: print("missing VehicleJourneys!!!")
LinksTest(VISUM,False)

##end
send2trash(access_db)
# del VISUM
Sekunden = int(time.time() - start_time)
print("--finished after ",Sekunden,"seconds--")