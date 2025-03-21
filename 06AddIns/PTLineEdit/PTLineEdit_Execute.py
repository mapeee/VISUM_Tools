# -*- coding: utf-8 -*-
"""
Created on Wed Mar 12 15:58:59 2025
"""
import os
from pathlib import Path
import sqlite3
from VisumPy.AddIn import AddIn
from VisumPy.helpers import SetMulti
_ = AddIn.gettext


def InitFilter(Visum):
    Visum.Filters.InitAll()
    SetMulti(Visum.Net.Nodes,"AddVal1",[0]*Visum.Net.Nodes.Count,False)
    SetMulti(Visum.Net.LineRoutes,"AddVal1",[0]*Visum.Net.LineRoutes.Count,False)

def PTExport(Visum, directory, Stops):
    if Visum.Net.LineRoutes.CountActive > 300:
        Visum.Log(12288,_("To many LineRoutes: %s") %(Visum.Net.LineRoutes.CountActive))
        return False
    
    vehjourneys = Visum.Net.VehicleJourneys.Count
    servingstops = sum(i[0] for i in Visum.Net.StopPoints.GetMultipleAttributes(["Count:ServingVehJourneys"]))
    chainedVehSec = Visum.Net.ChainedUpVehicleJourneySections.Count
    Nodes = [int(x[0]) for x in Visum.Net.Nodes.GetMultipleAttributes(["No"],True)]
    if len(Nodes) == 0: Nodes = [0]
    else: PTFilter(Visum)
    
    desktop_path = _desktop()
    sqlite_path = os.path.join(desktop_path, "PTdata.sqlite3")
    Visum.IO.SaveSQLiteDatabase(sqlite_path, os.path.join(directory, "Access_export.net"), True, False, True)    
    
    Visum.Net.LineRoutes.RemoveAll()
    Visum.Net.Connectors.RemoveAll()
    Visum.Net.StopAreas.RemoveAll()
    Visum.Net.StopPoints.RemoveAll()

    ## Sqlite3
    conn = sqlite3.connect(sqlite_path)
    cursor = conn.cursor()
             
    if Stops[0][0] != "0":
        for i in Stops:
            if i[0] == "0": continue
            cursor.execute('''
                            UPDATE LINEROUTEITEM
                            SET NODENO = ?, STOPPOINTNO = ?
                            WHERE STOPPOINTNO = ?
                            ''', (int(i[1]), int(i[1]), int(i[0])))              
            conn.commit()
        conn.close()
        
    return True, [vehjourneys, servingstops, chainedVehSec], Nodes

def PTFilter(Visum):
    Visum.Filters.InitAll()

    Nodes = Visum.Filters.NodeFilter()
    Nodes.AddCondition("OP_NONE",False,"AddVal1","EqualVal",1)

    Connector = Visum.Filters.ConnectorFilter()
    Connector.AddCondition("OP_NONE",False,r"Node\AddVal1","EqualVal",1)
    Connector.AddCondition("OP_AND",False,"TYPENO","EqualVal",9)
    
    Stops = Visum.Filters.StopGroupFilter()
    Stops.UseFilterForStopAreas = True
    Stops.UseFilterForStopPoints = True
    Stops = Stops.StopPointFilter()
    Stops.AddCondition("OP_NONE",False,r"Node\AddVal1","EqualVal",1)

    InsertedLinks = Visum.Filters.LinkFilter()
    InsertedLinks.AddCondition("OP_NONE",False,"TYPENO", "ContainedIn", str(1))
    
    Lines = Visum.Filters.LineGroupFilter()
    LineRouteItems = Lines.LineRouteItemFilter()
    LineRouteItems.AddCondition("OP_NONE",False,"ISROUTEPOINT","EqualVal",1)   
    LineRoutes = Lines.LineRouteFilter()
    LineRoutes.AddCondition("OP_NONE",False,"AddVal1","EqualVal",1)
    LineRoutes.AddCondition("OP_OR",False,r"COUNTACTIVE:STOPPOINTS","GreaterVal",0)
    Lines.UseFilterForLineRoutes = True
    Lines.UseFilterForLineRouteItems = True
    Lines.UseFilterForTimeProfiles = True
    Lines.UseFilterForTimeProfileItems = True
    Lines.UseFilterForVehJourneys = True
    Lines.UseFilterForVehJourneySections = True
    Lines.UseFilterForVehJourneyItems = True

def PTImport(Visum, changed_nodes = [0], PTcounts = False):   
    desktop_path = _desktop()
    sqlite_path = os.path.join(desktop_path, "PTdata.sqlite3")
    
    PuT_import = Visum.IO.CreateNetReadRouteSearch()
    import_setting = Visum.IO.CreateNetReadRouteSearchTSys()
    import_setting.SetAttValue("ChangeLinkTypeOfOpenedLinks", False)
    import_setting.SetAttValue("IncludeBlockedTurns", False)
    import_setting.SetAttValue("HowToHandleIncompleteRoute", 2) ##search shortest path
    import_setting.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths", 1)
    import_setting.SetAttValue("ShortestPathCriterion", 1) ##1 = travel time; 2 = travel time from linktype; 3 = link length
    import_setting.SetAttValue("MaxDeviationFactor", 50)
    import_setting.SetAttValue("WhatToDoIfShortestPathNotFound", 2) ##insert link if necessary
    import_setting.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths", 1)
    import_setting.SetAttValue("WhatToDoIfStopPointIsBlocked", 2) ##Open the StopPoint
    import_setting.SetAttValue("WhatToDoIfStopPointNotFound", 0) ##do not read lineroute
    PuT_import.SetForAllTSys(import_setting)
    Visum.IO.LoadSQLiteDatabase(sqlite_path, True, PuT_import)
    
    Visum.Filters.NodeFilter().Init()
    Nodes = Visum.Filters.NodeFilter()
    Nodes.AddCondition("OP_NONE", False, "AddVal1", "GreaterVal", 0)
    for Node in Visum.Net.Nodes.GetAllActive:
        Node.SetAttValue("Name", "")
        Node.SetAttValue("AddVal1", 0)
    
    Visum.Filters.NodeFilter().Init()
    Visum.Filters.StopGroupFilter().Init()
    Visum.Filters.ConnectorFilter().Init()
    
    ##Attributes for new Stop locations
    for i in changed_nodes:
        if i == 0: break
        Node = Visum.Net.Nodes.ItemByKey(i)
        
        Node.SetAttValue("Name",Node.AttValue(r"MIN:STOPPOINTS\Name"))
        Area = Visum.Net.StopAreas.ItemByKey(i)
        Area.SetAttValue("XCOORD",Area.AttValue(r"MIN:STOPPOINTS\XCOORD"))
        Area.SetAttValue("YCOORD",Area.AttValue(r"MIN:STOPPOINTS\YCOORD"))
        for Turn in Node.ViaNodeTurns.GetAll:
            if Turn.AttValue("ANGLE") == 360:
                Turn.SetAttValue("TSYSSET","")
        try: Stop = Visum.Net.Stops.ItemByKey(i)
        except: continue
        if int(Stop.AttValue("NumStopAreas")) == 1:
            Stop.SetAttValue("XCOORD",Stop.AttValue(r"MIN:STOPAREAS\XCOORD"))
            Stop.SetAttValue("YCOORD",Stop.AttValue(r"MIN:STOPAREAS\YCOORD"))

    Visum.Filters.LineGroupFilter().LineRouteFilter().RemoveCondition(2) ##Remove Aktive:StopPoint > 0

    ##tests for missing elements
    PTFilter(Visum)
    InsertedLinks = Visum.Filters.LinkFilter().Init() 
    InsertedLinks = Visum.Filters.LinkFilter()
    InsertedLinks.AddCondition("OP_NONE",False,"TYPENO", "ContainedIn", str(1))
    if Visum.Net.Links.CountActive > 0:
        Visum.Log(16384, _("> "+str(Visum.Net.Links.CountActive)+" new links of type "+str(1)+" added"))
    Visum.Filters.LinkFilter().Init()
    
    if PTcounts:
        journeys = PTcounts[0] - Visum.Net.VehicleJourneys.Count
        if journeys != 0:
            Visum.Log(12288, _("missing VehicleJourneys: "+str(journeys)))
            return False
        servingstops = int(PTcounts[1] - sum(i[0] for i in Visum.Net.StopPoints.GetMultipleAttributes(["Count:ServingVehJourneys"])))
        if servingstops != 0:
            Visum.Log(12288, _("missing servings at stops: "+str(servingstops)))
            return False
        chainedVehSec = PTcounts[2] - Visum.Net.ChainedUpVehicleJourneySections.Count
        if chainedVehSec != 0:
            Visum.Log(12288, _("missing Chained VehicleJourneySections: "+str(chainedVehSec)))
            return False

def SRtimeBus(Visum, mode):
    Visum.Filters.InitAll()
    SRLinks = Visum.Filters.LinkFilter()
    SRLinks.AddCondition("OP_NONE",False,"TSYSSET","ContainsAll","Bus")
    if mode:
        SRLinks.AddCondition("OP_AND",False,r"COUNTACTIVE:SYSROUTES","GreaterVal",0)
        tBus = [i[0] / 200 * 60 * 60 for i in Visum.Net.Links.GetMultipleAttributes(["LENGTHPOLY"],True)]
    else:
        tBus = [i[0] / i[1] * 60 * 60 for i in Visum.Net.Links.GetMultipleAttributes(["LENGTHPOLY", r"LINKTYPE\VDEF_PUTSYS(BUS)"],True)]
    SetMulti(Visum.Net.Links, r"T_PUTSYS(BUS)", tBus, True)
    Visum.Filters.InitAll()

def _desktop():
    home = Path.home()
    onedrive_path = Path(os.environ.get("OneDrive", ""))
    if (onedrive_path / "Desktop").exists():
        return onedrive_path / "Desktop"
    else:
        return home / "Desktop"  # Default desktop path