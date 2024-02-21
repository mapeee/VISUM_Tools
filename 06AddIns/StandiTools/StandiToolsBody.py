# -*- coding: utf-8 -*-

import numpy as np
import sys
import wx
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
from VisumPy.helpers import SetMulti
_ = AddIn.gettext

def Run(param):
    if param["CheckCon"]:
        CheckCon()
    if param["CheckTSys"]:
        if CheckTSys(): Visum.Log(20480,_("TSys Check: OK"))
    if param["CheckStops"]:
        CheckStops()
        
    if param["SetCon"]:
        SetCon()
    if param["SetPtt"]:
        SetPtt()
    if param["SetFareSystem"]:
        SetFareSystem()

    Visum.Log(20480,_("Standi Tools: finished"))

def CheckCon():
    if Visum.Net.Zones.AttrExists("PARKEN") == False:
        Visum.Log(12288,_("ZONE UDA 'PARKEN' is missing!"))
        return
    
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
        Visum.Log(12288,_("%s PUT-Connectors without StopArea!") %(str(PUTStopArea)))
        return False
    Visum.Log(20480,_("Connector Check: OK"))

def CheckStops():
    if Visum.Net.StopPoints.AttrExists("RAUSSTATTUNG") == False:
        Visum.Log(12288,_("StopPoints UDA 'RAUSSTATTUNG' is missing!"))

    StopAreaWalkTimes = 0
    if CheckTSys(["PUTWALK"]) == False: return
    for Stop in Visum.Net.Stops.GetAll:
        if Stop.StopAreas.Count <2:continue
        for StopAreafrom in Stop.StopAreas.GetAll:
            for StopAreato in Stop.StopAreas.GetAll:
                if Stop.StopAreaTransferWalkTime(StopAreafrom.AttValue("NO"), StopAreato.AttValue("NO"), "OVF") > 1800:
                    StopAreaWalkTimes += 1
    if StopAreaWalkTimes > 0:
        Visum.Log(12288,_("%s StopAreaTransferWalkTime(s) > 30 Minutes!") %(str(StopAreaWalkTimes)))
    
    StopAreaN = 0
    StopPointNo = 0
    _State = True
    for stopArea in Visum.Net.StopAreas.GetAll:
        if stopArea.AttValue("NODENO") == 0:
            Visum.Log(12288,_("At least on StopArea is not connected to Network (Node)!"))
            _State = False
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

def CheckTSys(Type = ["PUT","PrT","PUTWALK"]):
    if Visum.Net.TSystems.AttrExists("SCHIENENBONUS") == False:
        Visum.Log(12288,_("TSys UDA 'SCHIENENBONUS' is missing!"))
    
    for TSys in Visum.Net.TSystems.GetAll:
        if "PUT" in Type and TSys.AttValue("TYPE") == "PUT" and TSys.AttValue("CODE") not in ["Bus", "S", "U", "RV", "FV", "W"]:
            Visum.Log(12288,_("%s TSys %s (%s) is missing!") %("PUT", TSys.AttValue("CODE"), TSys.AttValue("NAME")))
            return False
        if "PRT" in Type and TSys.AttValue("TYPE") == "PRT" and TSys.AttValue("CODE") not in ["B", "C", "F", "T"]:
            Visum.Log(12288,_("%s TSys %s (%s) is missing!") %("PRT", TSys.AttValue("CODE"), TSys.AttValue("NAME")))
            return False
        if "PUTWALK" in Type and TSys.AttValue("TYPE") == "PUTWALK" and TSys.AttValue("CODE") != "OVF":
            Visum.Log(12288,_("%s TSys %s (%s) is missing!") %("PUTWALK", TSys.AttValue("CODE"), TSys.AttValue("NAME")))
            return False
    return True

def SetCon():
    if CheckTSys(["PUTWALK", "PRT"]) == False: return
    if Visum.Net.Zones.AttrExists("LAGE") == False:
        Visum.Log(12288,_("ZONE UDA 'LAGE' is missing!"))
        return
    
    locations = Visum.Net.Connectors.GetMultiAttValues("ZONE\LAGE",False)
    
    newPuT = [[i[0],_conTimes(i[0], i[1])] for i in Visum.Net.Connectors.GetMultiAttValues("LENGTH",False)]
    Visum.Net.Connectors.SetMultiAttValues("T0_TSYS(OVF)",newPuT)
    
    newPrT = [[i[0],_conTimes(i[0], i[1], locations)] for i in Visum.Net.Connectors.GetMultiAttValues("LENGTH",False)]
    Visum.Net.Connectors.SetMultiAttValues("T0_TSYS(B)",newPrT)
    Visum.Net.Connectors.SetMultiAttValues("T0_TSYS(C)",newPrT)
    Visum.Net.Connectors.SetMultiAttValues("T0_TSYS(F)",newPrT)
    Visum.Net.Connectors.SetMultiAttValues("T0_TSYS(T)",newPrT)
    
    Visum.Log(20480,_("Set Connector Times: OK"))

def SetPtt():
    if CheckTSys(["PUT"]) == False: return
    
    for i in ["S","U","FV","RV","W","Bus"]:
        if i in np.array(Visum.Net.TSystems.GetMultiAttValues("CODE"))[:,1] == False:
            Visum.Log(12288,_("TSys %s is missing!") %(i))
    
    for i in ["KLASSE", "BUSSPUR"]:
        if Visum.Net.Links.AttrExists(i) == False:
            Visum.Log(12288,_("Link UDA %s is missing!") %(i))
            return
    for i in ["ABSAUFSCHLAG", "RELAUFSCHLAG"]:
        if Visum.Net.LineRouteItems.AttrExists(i) == False:
            Visum.Log(12288,_("LineRouteItems UDA %s is missing!") %(i))
            return
    
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
    
def SetFareSystem():
    if CheckTSys(["PUT"]) == False: return

    FareList = Visum.Lists.CreateFareSystemList
    FareList.AddColumn("No")
    if FareList.NumActiveElements != 2 or int(FareList.Sum(0)) != 3: 
        Visum.Log(12288,_("TicketSystem (NO 1: Zuschlag or NO 2: Ohne) is missing!"))
        return
        
    Visum.Filters.InitAll()
    Lines = Visum.Net.Lines.GetAll
    for Line in Lines:
        if Line.AttValue("FARESYSTEMSET") == "":Line.SetAttValue("FARESYSTEMSET",2)
        if Line.AttValue("TSYSCODE") == "FV" and int(Line.AttValue("FARESYSTEMSET")) !=1 :Line.SetAttValue("FARESYSTEMSET",1)
    Visum.Log(20480,_("Set FareSystem: OK"))

def _conTimes(_index,_length,_locations=False):
    tacegr = 0.46*pow((_length*1000*1.2),0.4) #1.2 == Umwegfaktor
    tacegr = max(tacegr,3)
    tacegr = (tacegr*(0.9+(0.15*tacegr)))*60
    tacegr = min(tacegr, 3600)
    
    if _locations != False:
        if _locations[_index-1][1] == 1:
            tacegr = min(tacegr,600)
        elif _locations[_index-1][1] == 3:
            tacegr = min(tacegr,1200)
        else:
            tacegr = min(tacegr,3600)
    return tacegr
       
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
