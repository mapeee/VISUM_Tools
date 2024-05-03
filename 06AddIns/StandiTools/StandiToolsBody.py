# -*- coding: utf-8 -*-

import numpy as np
import sys
import wx
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
from VisumPy.helpers import SetMulti
_ = AddIn.gettext

def Run(param):
    if param["CheckCon"]:
        if not CheckCon(): return
    if param["CheckTSys"]:
        if CheckTSys():
            Visum.Log(20480,_("TSys Check: OK"))
        else: return
    if param["CheckStops"]:
        if not CheckStops(): return
    if param["SetCon"]:
        if not SetCon(): return
    if param["SetPtt"]:
        if not SetPtt(): return
    if param["SetFareSystem"]:
        if not SetFareSystem(): return
    if param["SetVarious"]:
        if not SetVarious(): return

    Visum.Log(20480,_("Standi Tools: finished"))

def CheckCon():
    if Visum.Net.Zones.AttrExists("PARKEN") == False:
        addIn.ReportMessage(_("ZONE UDA 'PARKEN' is missing!"))
        return False
    
    PUTWALK_PRT = 0
    PUTStopArea = 0
    for con in Visum.Net.Connectors.GetAll:
        _TSysSet = con.AttValue("TSYSSET").split(",")
        if "OVF" in _TSysSet and len(_TSysSet)>1: 
            PUTWALK_PRT += 1
        if "OVF" in _TSysSet and con.AttValue("NODE\COUNT:STOPAREAS") == 0:
            PUTStopArea += 1
    if PUTWALK_PRT > 0:
        Visum.Log(20480,_("%s Connector(s) are open for PUTWALK and PrT!") %(str(PUTWALK_PRT)))
    if PUTStopArea > 0:
        addIn.ReportMessage(_("%s PUT-Connectors without StopArea!") %(str(PUTStopArea)))
        return False
    
    Visum.Filters.ConnectorFilter().Init()
    Visum.Filters.ConnectorFilter().AddCondition("OP_NONE",False,"TSYSSET",14,"OVF")
    nPuT = len(np.unique(np.array(Visum.Net.Connectors.GetMultiAttValues("ZONENO"))[:,1]))
    if nPuT != Visum.Net.Zones.Count:
        addIn.ReportMessage(_("%s Zones missing PuT-Connector!") %(str(Visum.Net.Zones.Count-nPuT)))
        return False
    Visum.Filters.ConnectorFilter().Init()
    Visum.Filters.ConnectorFilter().AddCondition("OP_NONE",False,"TSYSSET",14,"C,T")
    nPrT = len(np.unique(np.array(Visum.Net.Connectors.GetMultiAttValues("ZONENO"))[:,1]))
    if nPrT != Visum.Net.Zones.Count:
        addIn.ReportMessage(_("%s Zones missing PrT-Connector!") %(str(Visum.Net.Zones.Count-nPrT)))
        return False
    Visum.Filters.ConnectorFilter().Init()
    
    Visum.Log(20480,_("Connector Check: OK"))
    return True

def CheckStops():
    if Visum.Net.StopPoints.AttrExists("RAUSSTATTUNG") == False:
        addIn.ReportMessage(_("StopPoints UDA 'RAUSSTATTUNG' is missing!"))
        return False

    StopAreaWalkTimes = 0
    if CheckTSys(["PUTWALK"]) == False: return False
    for Stop in Visum.Net.Stops.GetAll:
        if Stop.StopAreas.Count <2:continue
        for StopAreafrom in Stop.StopAreas.GetAll:
            for StopAreato in Stop.StopAreas.GetAll:
                if Stop.StopAreaTransferWalkTime(StopAreafrom.AttValue("NO"), StopAreato.AttValue("NO"), "OVF") > 1800:
                    StopAreaWalkTimes += 1
    if StopAreaWalkTimes > 0:
        addIn.ReportMessage(_("%s StopAreaTransferWalkTime(s) > 30 Minutes!") %(str(StopAreaWalkTimes)))
        return False
    
    StopAreaN = 0
    StopPointNo = 0
    _State = True
    for stopArea in Visum.Net.StopAreas.GetAll:
        if stopArea.AttValue("NODENO") == 0:
            addIn.ReportMessage(_("At least on StopArea is not connected to Network (Node)!"))
            return False
        if stopArea.AttValue("NUMSTOPPOINTS") != 1:
            StopAreaN += 1
        elif stopArea.AttValue("NO") != stopArea.AttValue(r"MIN:STOPPOINTS\NO") or stopArea.AttValue("NO") != stopArea.AttValue("NODENO"):
            StopPointNo += 1
        else:pass
        
    if StopAreaN > 0:
        Visum.Log(20480,_("%s StopArea(s) contain 0 or more than 1 StopPoint") %(str(StopAreaN)))
    if StopPointNo > 0:
        Visum.Log(12288,_("%s StopArea(s) not sharing same NO with StopPoint and Node!") %(str(StopPointNo)))
        _State = False
    if _State: Visum.Log(20480,_("Stops Check: OK"))
    return True

def CheckTSys(Type = ["PUT","PrT","PUTWALK"]): 
    for TSys in Visum.Net.TSystems.GetAll:
        if "PUT" in Type and TSys.AttValue("TYPE") == "PUT" and TSys.AttValue("CODE") not in ["Bus", "S", "U", "RV", "FV", "W"]:
            addIn.ReportMessage(_("%s TSys %s (%s) is missing!") %("PUT", TSys.AttValue("CODE"), TSys.AttValue("NAME")))
            return False
        if "PRT" in Type and TSys.AttValue("TYPE") == "PRT" and TSys.AttValue("CODE") not in ["B", "C", "F", "T"]:
            addIn.ReportMessage(_("%s TSys %s (%s) is missing!") %("PRT", TSys.AttValue("CODE"), TSys.AttValue("NAME")))
            return False
        if "PUTWALK" in Type and TSys.AttValue("TYPE") == "PUTWALK" and TSys.AttValue("CODE") != "OVF":
            addIn.ReportMessage(_("%s TSys %s (%s) is missing!") %("PUTWALK", TSys.AttValue("CODE"), TSys.AttValue("NAME")))
            return False
    return True

def SetCon():
    if CheckTSys(["PUTWALK", "PRT"]) == False: return False
    if Visum.Net.Zones.AttrExists("LAGE") == False:
        addIn.ReportMessage(_("ZONE UDA 'LAGE' is missing!"))
        return False
    
    locations = Visum.Net.Connectors.GetMultiAttValues("ZONE\LAGE",False)
    
    newPuT = [[i[0],_conTimes(i[0], i[1])] for i in Visum.Net.Connectors.GetMultiAttValues("LENGTH",False)]
    Visum.Net.Connectors.SetMultiAttValues("T0_TSYS(OVF)",newPuT)
    
    newPrT = [[i[0],_conTimes(i[0], i[1], locations)] for i in Visum.Net.Connectors.GetMultiAttValues("LENGTH",False)]
    Visum.Net.Connectors.SetMultiAttValues("T0_TSYS(B)",newPrT)
    Visum.Net.Connectors.SetMultiAttValues("T0_TSYS(C)",newPrT)
    Visum.Net.Connectors.SetMultiAttValues("T0_TSYS(F)",newPrT)
    Visum.Net.Connectors.SetMultiAttValues("T0_TSYS(T)",newPrT)
    
    Visum.Log(20480,_("Set Connector Times: OK"))
    return True

def SetPtt():
    if CheckTSys(["PUT"]) == False: return False
    
    for i in ["S","U","FV","RV","W","Bus"]:
        if np.any(np.array(Visum.Net.TSystems.GetMultiAttValues("CODE"))[:,1] == i) == False:
            Visum.Log(12288,_("TSys %s is missing!") %(i))
    
    for i in ["KLASSE", "BUSSPUR", "GRUNDBELASTUNG"]:
        if Visum.Net.Links.AttrExists(i) == False:
            addIn.ReportMessage(_("Link UDA %s is missing!") %(i))
            return False
    for i in ["GRUNDBELASTUNG"]:
        if Visum.Net.Turns.AttrExists(i) == False:
            addIn.ReportMessage(_("Turn UDA %s is missing!") %(i))
            return False
    for i in ["ABSAUFSCHLAG", "RELAUFSCHLAG"]:
        if Visum.Net.LineRouteItems.AttrExists(i) == False:
            addIn.ReportMessage(_("LineRouteItems UDA %s is missing!") %(i))
            return False
    
    Visum.Filters.InitAll()
    Lines = Visum.Filters.LineGroupFilter()
    Lines.UseFilterForLineRouteItems = True
    Line = Lines.LineFilter()
    LineRouteItem = Lines.LineRouteItemFilter()

    SetMulti(Visum.Net.LineRouteItems,"ABSAUFSCHLAG",[1.8]*Visum.Net.LineRouteItems.CountActive,True)
    SetMulti(Visum.Net.LineRouteItems,"RELAUFSCHLAG",[0.18]*Visum.Net.LineRouteItems.CountActive,True)

    #Rail/water separated
    Line.AddCondition("OP_NONE",False,"TSYSCODE",13,"S,U,W")
    SetMulti(Visum.Net.LineRouteItems,"ABSAUFSCHLAG",[0]*Visum.Net.LineRouteItems.CountActive,True)
    SetMulti(Visum.Net.LineRouteItems,"RELAUFSCHLAG",[0]*Visum.Net.LineRouteItems.CountActive,True)
    Line.Init()
    #Bus
    Line.AddCondition("OP_NONE",False,"TSYSCODE",13,"Bus")
    SetMulti(Visum.Net.LineRouteItems,"ABSAUFSCHLAG",[1.8]*Visum.Net.LineRouteItems.CountActive,True)
    SetMulti(Visum.Net.LineRouteItems,"RELAUFSCHLAG",[0.18]*Visum.Net.LineRouteItems.CountActive,True)
    Line.Init()
    #Bus lane
    Line.AddCondition("OP_NONE",False,"TSYSCODE",13,"Bus")
    LineRouteItem.AddCondition("OP_NONE",False,"OUTLINK\BUSSPUR",3,0)
    LineRouteItem.AddCondition("OP_AND",False,"OUTLINK\BUSSPUR",1,5)
    SetMulti(Visum.Net.LineRouteItems,"ABSAUFSCHLAG",[1.2]*Visum.Net.LineRouteItems.CountActive,True)
    SetMulti(Visum.Net.LineRouteItems,"RELAUFSCHLAG",[0.12]*Visum.Net.LineRouteItems.CountActive,True)
    Line.Init()
    LineRouteItem.Init()
    #S-Bahn Mischbetrieb
    Line.AddCondition("OP_NONE",False,"TSYSCODE",13,"S")
    LineRouteItem.AddCondition("OP_NONE",False,"OUTLINK\KLASSE",9,101)
    SetMulti(Visum.Net.LineRouteItems,"ABSAUFSCHLAG",[1.2]*Visum.Net.LineRouteItems.CountActive,True)
    SetMulti(Visum.Net.LineRouteItems,"RELAUFSCHLAG",[0.12]*Visum.Net.LineRouteItems.CountActive,True)
    Line.Init()
    LineRouteItem.Init()
    #FV
    Line.AddCondition("OP_NONE",False,"TSYSCODE",13,"FV")
    SetMulti(Visum.Net.LineRouteItems,"ABSAUFSCHLAG",[1.2]*Visum.Net.LineRouteItems.CountActive,True)
    SetMulti(Visum.Net.LineRouteItems,"RELAUFSCHLAG",[0.12]*Visum.Net.LineRouteItems.CountActive,True)
    Line.Init()
    #RV
    Line.AddCondition("OP_NONE",False,"TSYSCODE",13,"RV")
    SetMulti(Visum.Net.LineRouteItems,"ABSAUFSCHLAG",[1.2]*Visum.Net.LineRouteItems.CountActive,True)
    SetMulti(Visum.Net.LineRouteItems,"RELAUFSCHLAG",[0.12]*Visum.Net.LineRouteItems.CountActive,True)
    Line.Init()
    #AKN
    Line.AddCondition("OP_NONE",False,"Liniennummer",9,"A1")
    Line.AddCondition("OP_OR",False,"Liniennummer",9,"A2")
    Line.AddCondition("OP_OR",False,"Liniennummer",9,"A3")
    SetMulti(Visum.Net.LineRouteItems,"ABSAUFSCHLAG",[0]*Visum.Net.LineRouteItems.CountActive,True)
    SetMulti(Visum.Net.LineRouteItems,"RELAUFSCHLAG",[0]*Visum.Net.LineRouteItems.CountActive,True)
    Line.Init()
    
    Visum.Log(20480,_("Set Perceived travel times: OK"))
    return True
    
def SetFareSystem():
    if CheckTSys(["PUT"]) == False: return False

    FareList = Visum.Workbench.Lists.CreateFareSystemList
    FareList.AddColumn("No")
    if FareList.NumActiveElements != 2 or int(FareList.Sum(0)) != 3: 
        addIn.ReportMessage(_("TicketSystem (NO 1: Zuschlag or NO 2: Ohne) is missing!"))
        return False
        
    Visum.Filters.InitAll()
    Lines = Visum.Net.Lines.GetAll
    for Line in Lines:
        if Line.AttValue("FARESYSTEMSET") == "":Line.SetAttValue("FARESYSTEMSET",2)
        if Line.AttValue("TSYSCODE") == "FV" and int(Line.AttValue("FARESYSTEMSET")) !=1 :Line.SetAttValue("FARESYSTEMSET",1)
    Visum.Log(20480,_("Set FareSystem: OK"))
    return True

def SetVarious():
    if Visum.Net.Links.AttrExists("VMAX_SCHIENE") == False: return
    tRail = [[_RailTimes(i[0], i[1])][0] for i in tuple((x, y) for x, y in zip(Visum.Net.Links.GetMultiAttValues("LENGTH",False), Visum.Net.Links.GetMultiAttValues("VMAX_SCHIENE",False)))]
    SetMulti(Visum.Net.Links,r"T_PUTSYS(FV)",tRail,False)
    Visum.Log(20480,_("Set Various: OK"))
    return True
    
def _conTimes(_index,_length,_locations=False):
    tacegr = 0.46*pow((_length*1000*1.2),0.4) #1.2 == Umwegfaktor
    tacegr = max(tacegr,3)
    tacegr = (tacegr*(0.9+(0.15*tacegr)))*60
    tacegr = min(tacegr, 3600)
    
    if _locations != False:
        if _locations[_index-1][1] == 1:
            tacegr = min(tacegr,300)
        elif _locations[_index-1][1] == 3:
            tacegr = min(tacegr,900)
        else:
            tacegr = min(tacegr,3600)
    return tacegr

def _RailTimes(_length, _speed):
    l = _length[1]
    v = _speed[1]
    if v == 0: return 0
    return l / v * 60 * 60
       
if len(sys.argv) > 1:
    addIn = AddIn()
else:
    addIn = AddIn(Visum)

if addIn.IsInDebugMode:
    app = wx.PySimpleApp(0)
    Visum = addIn.VISUM
    addInParam = AddInParameter(addIn, None)
else:
    addInParam = AddInParameter(addIn, Parameter)

if addIn.State != AddInState.OK:
    addIn.ReportMessage(addIn.ErrorObjects[0].ErrorMessage)
else:
    try:
        defaultParam = {"CheckTSys" : False, "CheckCon" : False, "CheckStops" : False,
                        "SetCon" : False, "SetPtt" : False, "SetFareSystem" : False}
        param = addInParam.Check(True, defaultParam)
        Run(param)
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
