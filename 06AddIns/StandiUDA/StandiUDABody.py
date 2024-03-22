# -*- coding: utf-8 -*-

import numpy as np
import sys
import wx
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
from VisumPy.helpers import SetMulti
_ = AddIn.gettext

def Run(param):
    V = _("Standi-Verfahren")
    
    comments = [["SCHIENENBONUS", _("Weighting of travel time for rail transport with 0.5")],
                ["LAGE", _("Location in model area (PR=1, UR=3, EUR=4, UML=5)")],
                ["PARKEN", _("Parking availability")],
                ["GRUNDBELASTUNG", _("Base volume based on incoming link")],
                ["BUSSPUR", _("0=no bus lane, 1=bus lane right, 3=bus lane mid, 4=Environmental route")],
                ["GRUNDBELASTUNG", _("Base volume based on Bus lane")],
                ["KLASSE", _("Track class (based on Ras N)")],
                ["RAUSSTATTUNG", _("Penalty for stop equipment (Standi 2016+)")],
                ["ABSAUFSCHLAG", _("absIncrease")],
                ["RELAUFSCHLAG", _("relIncrease")],
                
                ["BEL_FV_MIT", _("Volume Long-distance PT planning case")],
                ["BEL_FV_OHNE", _("Volume Long-distance PT base case")],
                ["BEL_MIV_MIT", _("Volume MT (car+truck+PT base Volume) planning case")],
                ["BEL_MIV_OHNE", _("Volume MT (car+truck+PT base Volume) base case")],
                ["BEL_OEV_MIT", _("Volume PT planning case")],
                ["BEL_OEV_OHNE", _("Volume PT base case")]]
    
    if Visum.Net.UserDefinedGroups.Count > 0:
        if np.any(np.array(Visum.Net.UserDefinedGroups.GetMultiAttValues("Name"))[:,1] == V) ==False: Visum.Net.AddUserDefinedGroup(V)
    else: Visum.Net.AddUserDefinedGroup(V)

    for i in param["UDA_a"]:  
        TYPE, ID, Name = i.split(" : ")
        if Visum.Net.Links.AttrExists(ID)==True:
            if param["Replace"]==True: SetMulti(Visum.Net.Links,ID,[0]*Visum.Net.Links.CountActive,True)
            continue
        Visum.Net.Links.AddUserDefinedAttribute(ID, Name, Name,1)
        UDA = Visum.Net.Links.Attributes.ItemByKey(ID)
        UDA.UserDefinedGroup = V
        UDA.Comment = comments[[x[0] for x in comments].index(ID)][1]

    for i in param["UDA_m"]:
        TYPE, ID, Name = i.split(" : ")
        if TYPE in ["TSYS","Verkehrssysteme"]:
            if Visum.Net.TSystems.AttrExists(ID)==True: continue
            Visum.Net.TSystems.AddUserDefinedAttribute(ID, Name, Name,2,defval=1.0)
            UDA = Visum.Net.TSystems.Attributes.ItemByKey(ID)
        if TYPE in ["ZONE","Bezirke"]:
            if Visum.Net.Zones.AttrExists(ID)==True: continue
            if ID =="LAGE": Visum.Net.Zones.AddUserDefinedAttribute(ID, Name, Name,1)
            else: Visum.Net.Zones.AddUserDefinedAttribute(ID, Name, Name,2,defval=1.0)
            UDA = Visum.Net.Zones.Attributes.ItemByKey(ID)
        if TYPE in ["TURN", "Abbieger"]:
            if Visum.Net.Turns.AttrExists(ID)==True: continue
            Visum.Net.Turns.AddUserDefinedAttribute(ID, Name, Name,1,formula="if([VONSTRECKE\BUSSPUR]>0&[VONSTRECKE\BUSSPUR]<4;0;[ANZSERVICEFAHRT-VSYS(BUS,AP)])")
            UDA = Visum.Net.Turns.Attributes.ItemByKey(ID)
        if TYPE in ["LINK","Strecken"]:
            if Visum.Net.Links.AttrExists(ID)==True: continue
            if ID == "GRUNDBELASTUNG":
                Visum.Net.Links.AddUserDefinedAttribute(ID, Name, Name,1,formula="if([BUSSPUR]>0&[BUSSPUR]<4;0;[ANZSERVICEFAHRT-VSYS(BUS,AP)])")
            else:
                Visum.Net.Links.AddUserDefinedAttribute(ID, Name, Name,1)
            UDA = Visum.Net.Links.Attributes.ItemByKey(ID)
        if TYPE in ["STOPPOINT","Haltepunkte"]:
            if Visum.Net.StopPoints.AttrExists(ID)==True: continue
            Visum.Net.StopPoints.AddUserDefinedAttribute(ID, Name, Name,2)
            UDA = Visum.Net.StopPoints.Attributes.ItemByKey(ID)
        if TYPE in ["LINEROUTEITEM","LinienRoutenElemente"]:
            if Visum.Net.LineRouteItems.AttrExists(ID)==True: continue
            if ID == "ABSAUFSCHLAG": Visum.Net.LineRouteItems.AddUserDefinedAttribute(ID, Name, Name,2,defval=1.8)
            if ID == "RELAUFSCHLAG": Visum.Net.LineRouteItems.AddUserDefinedAttribute(ID, Name, Name,2,defval=0.18) 
            UDA = Visum.Net.LineRouteItems.Attributes.ItemByKey(ID)
        
        UDA.UserDefinedGroup = V
        UDA.Comment = comments[[x[0] for x in comments].index(ID)][1]

    Visum.Log(20480,_("Standi-UDA: created"))

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
        defaultParam = {"UDA_m" : "...", "Items_m" : "...", "UDA_a" : "...",
                        "Items_a" : "...", "Replace" : False}
        param = addInParam.Check(True, defaultParam)
        Run(param)
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
