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
import sys

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
path = open(path, mode='r').readlines()
path = path[0] +"\\"+'VISUM_Tools'+"\\"+'Line_import.txt'
f = open(path, mode='r').read().splitlines()
sys.path.insert(0,f[3])
from VisumPy.helpers import SetMulti

#--Path--#
Network = f[0]
layout_path = f[1]
access_db = f[2].replace("accdb","mdb")

def setSRtimeBus(_Visum):
    _Visum.Filters.InitAll()
    SRLinks = _Visum.Filters.LinkFilter()
    SRLinks.AddCondition("OP_NONE",False,r"COUNTACTIVE:SYSROUTES","GreaterVal",0)
    SRLinks.AddCondition("OP_AND",False,"TSYSSET","ContainsAll","Bus")
    SetMulti(Visum.Net.Links, r"T_PUTSYS(BUS)", Visum.Net.Links.CountActive*[1], True)
    print("> Set bus TravelTime on SystemRoutes")
    _Visum.Filters.InitAll()
    
def Visum_open(Net):
    _Visum = win32com.client.dynamic.Dispatch("Visum.Visum.24")
    _Visum.IO.loadversion(Net)
    _Visum.Filters.InitAll()
    return _Visum

def Visum_filter(_Visum,**optional):
    _Visum.Filters.InitAll()
    
    Lines = _Visum.Filters.LineGroupFilter()
    
    LineRouteItems = Lines.LineRouteItemFilter()
    LineRouteItems.AddCondition("OP_NONE",False,"ISROUTEPOINT","EqualVal",1)   
    
    LineRoutes = Lines.LineRouteFilter()
    LineRoutes.AddCondition("OP_NONE",False,"AddVal1","EqualVal",1) ##no more filter,not complementary
    LineRoutes.AddCondition("OP_OR",False,r"COUNTACTIVE:STOPPOINTS","GreaterVal",0) ##no more filter,not complementary
    
    Lines.UseFilterForLineRoutes = True
    Lines.UseFilterForLineRouteItems = True
    Lines.UseFilterForTimeProfiles = True
    Lines.UseFilterForTimeProfileItems = True
    Lines.UseFilterForVehJourneys = True
    Lines.UseFilterForVehJourneySections = True
    Lines.UseFilterForVehJourneyItems = True

    Con_type = optional.get("Con_type", False)
    Connector = _Visum.Filters.ConnectorFilter()
    Connector.AddCondition("OP_NONE",False,r"Node\AddVal1","EqualVal",1)
    if Con_type != False: Connector.AddCondition("OP_AND",False,"TYPENO","EqualVal",Con_type)
    
    Nodes = _Visum.Filters.NodeFilter()
    Nodes.AddCondition("OP_NONE",False,"AddVal1","EqualVal",1)
    
    HST = _Visum.Filters.StopGroupFilter()
    HST.UseFilterForStopAreas = True
    HST.UseFilterForStopPoints = True
    HST = HST.StopPointFilter()
    HST.AddCondition("OP_NONE",False,r"Node\AddVal1","EqualVal",1)

    InsertedLinks = _Visum.Filters.LinkFilter()
    InsertedLinks.AddCondition("OP_NONE",False,"TYPENO", "ContainedIn", str(1))
    
    Lines.UseFilterForLineRoutes = True
    
def Visum_export(_Visum,layout,access,**optional):
    global journeys_before
    global servingstops_before
    journeys_before = _Visum.Net.VehicleJourneys.Count ##number of journeys before export
    servingstops_before = _Visum.Net.StopPoints.GetMultipleAttributes(["Count:ServingVehJourneys"])
    servingstops_before = int(sum(list(map(sum, list(servingstops_before)))))
    
    _Visum.IO.SaveAccessDatabase(access,layout,True,False,True)
    _Visum.Net.LineRoutes.RemoveAll()
    _Visum.Net.Connectors.RemoveAll()
    _Visum.Net.StopAreas.RemoveAll()
    _Visum.Net.StopPoints.RemoveAll()

    global change_nodes
    change_nodes = _Visum.Net.Nodes.GetMultipleAttributes(["No"],True)

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

    Stops = optional.get("Stops", False)
    if Stops == False: return
    for i in Stops:
        cursor.execute('''
                        UPDATE LINEROUTEITEM
                        SET NODENO = '''+str(i[1])+''', STOPPOINTNO = '''+str(i[1])+'''
                        WHERE STOPPOINTNO = '''+str(i[0])+''' and ADDVAL = '''+str(1)+'''
                        ''')              
        conn.commit()
    conn.close()
 
def Visum_import(_Visum,access,LinkType,shortcrit,open_blocked):
    Visum_filter(_Visum)
    import_setting = _Visum.IO.CreateNetReadRouteSearchTsys()
    import_setting.SetAttValue("ChangeLinkTypeOfOpenedLinks",open_blocked)
    import_setting.SetAttValue("IncludeBlockedTurns",open_blocked)
    import_setting.SetAttValue("HowToHandleIncompleteRoute", 2) ##search shortest path
    import_setting.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths",1)
    import_setting.SetAttValue("ShortestPathCriterion",shortcrit) ##1 = travel time; 2 = travel time from linktype; 3 = link length
    import_setting.SetAttValue("MaxDeviationFactor",50)
    import_setting.SetAttValue("WhatToDoIfShortestPathNotFound",2) ##insert link if necessary
    import_setting.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths",LinkType)
    import_setting.SetAttValue("WhatToDoIfStopPointIsBlocked", 2) ##Open the StopPoint
    import_setting.SetAttValue("WhatToDoIfStopPointNotFound", 0) ##do not read lineroute
    PuT_import = _Visum.IO.CreateNetReadRouteSearch()
    PuT_import.SetForAllTSys(import_setting)
    _Visum.IO.LoadAccessDatabase(access,True,PuT_import)
    
    _Visum.Filters.NodeFilter().Init()
    Nodes = _Visum.Filters.NodeFilter()
    Nodes.AddCondition("OP_NONE",False,"AddVal1","GreaterVal",0)
    for N in _Visum.Net.Nodes.GetAllActive:
        N.SetAttValue("Name","")
        N.SetAttValue("AddVal1",0)
    
    _Visum.Filters.NodeFilter().Init()
    _Visum.Filters.StopGroupFilter().Init()
    _Visum.Filters.ConnectorFilter().Init()
    
    ##location / name of Nodes, Stops and StopAreas
    if len(change_nodes) > 0:
        for i in change_nodes:
            Node = _Visum.Net.Nodes.ItemByKey(int(i[0]))
            
            Node.SetAttValue("Name",Node.AttValue(r"MIN:STOPPOINTS\Name"))
            Area = _Visum.Net.StopAreas.ItemByKey(int(i[0]))
            Area.SetAttValue("XCOORD",Area.AttValue(r"MIN:STOPPOINTS\XCOORD"))
            Area.SetAttValue("YCOORD",Area.AttValue(r"MIN:STOPPOINTS\YCOORD"))
            try: Stop = _Visum.Net.Stops.ItemByKey(int(i[0]))
            except: continue
            if int(Stop.AttValue("NumStopAreas")) == 1:
                Stop.SetAttValue("XCOORD",Stop.AttValue(r"MIN:STOPAREAS\XCOORD"))
                Stop.SetAttValue("YCOORD",Stop.AttValue(r"MIN:STOPAREAS\YCOORD"))

    #--testing after the import--#
    InsertedLinks = _Visum.Filters.LinkFilter().Init() 
    InsertedLinks = _Visum.Filters.LinkFilter()
    InsertedLinks.AddCondition("OP_NONE",False,"TYPENO", "ContainedIn", str(LinkType))
    if _Visum.Net.Links.CountActive > 0: print("> "+str(_Visum.Net.Links.CountActive)+" new links of type "+str(LinkType)+" added \n") 
    _Visum.Filters.LinkFilter().Init()
    
    global journeys_after
    journeys_after = _Visum.Net.VehicleJourneys.Count ##number of journeys after processing
    if journeys_before != journeys_after:
        print("> missing VehicleJourneys")
        print("> before: "+str(journeys_before))
        print("> after: "+str(journeys_after)+"\n")

    global servingstops_after
    servingstops_after = _Visum.Net.StopPoints.GetMultipleAttributes(["Count:ServingVehJourneys"])
    servingstops_after = int(sum(list(map(sum, list(servingstops_after)))))
    if servingstops_before != servingstops_after:
        print("> missing servings at stops")
        print("> before: "+str(servingstops_before))
        print("> after: "+str(servingstops_after)+"\n")
        
    _Visum.Filters.LineGroupFilter().LineRouteFilter().RemoveCondition(2) ##Remove Aktive:StopPoint > 0 

def Visum_end(_Visum):
    SetMulti(_Visum.Net.LineRoutes, "AddVal1", Visum.Net.LineRoutes.CountActive*[0], True)
    _Visum.Filters.InitAll()
    Visum_filter(_Visum)

#--processing--#
Visum = Visum_open(Network)

setSRtimeBus(Visum)

'''Export Line only or StopPoint to different Node'''
Visum_filter(_Visum=Visum, Con_type=9)
Visum_export(_Visum=Visum, layout=layout_path, access=access_db)
Visum_import(_Visum=Visum, access=access_db, LinkType=1, shortcrit=1, open_blocked=False) ##1=time; 2=linktype; 3=length

'''Change StopPoint from LineRoutes'''
Stop = [[0,0],[0,0]]
Visum_filter(_Visum=Visum)
Visum_export(_Visum=Visum, layout=layout_path, access=access_db, Stops=Stop)
Visum_import(_Visum=Visum, access=access_db, LinkType=1, shortcrit=1, open_blocked=False) ##1=time; 2=linktype; 3=length

##end
Visum_end(_Visum=Visum)
Visum.SaveVersion(Network)
