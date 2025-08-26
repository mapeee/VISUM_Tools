# -*- coding: utf-8 -*-

import sys
import wx
from VisumPy.helpers import SetMulti
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext


def Run():
    SetMulti(Visum.Net.Nodes,"AddVal1",[0]*Visum.Net.Nodes.Count,False)
    SetMulti(Visum.Net.Nodes,"AddVal2",[0]*Visum.Net.Nodes.Count,False)
    SetMulti(Visum.Net.Nodes,"AddVal3",[0]*Visum.Net.Nodes.Count,False)
    
    if Visum.Net.MainNodes.Count > 0:
        SetMulti(Visum.Net.MainNodes,"AddVal1",[0]*Visum.Net.MainNodes.Count,False)
        SetMulti(Visum.Net.MainNodes,"AddVal2",[0]*Visum.Net.MainNodes.Count,False)
        SetMulti(Visum.Net.MainNodes,"AddVal3",[0]*Visum.Net.MainNodes.Count,False)
    
    SetMulti(Visum.Net.Links,"AddVal1",[0]*Visum.Net.Links.Count,False)
    SetMulti(Visum.Net.Links,"AddVal2",[0]*Visum.Net.Links.Count,False)
    SetMulti(Visum.Net.Links,"AddVal3",[0]*Visum.Net.Links.Count,False)
    
    SetMulti(Visum.Net.Turns,"AddVal1",[0]*Visum.Net.Turns.Count,False)
    SetMulti(Visum.Net.Turns,"AddVal2",[0]*Visum.Net.Turns.Count,False)
    SetMulti(Visum.Net.Turns,"AddVal3",[0]*Visum.Net.Turns.Count,False)
    
    if Visum.Net.MainTurns.Count > 0:
        SetMulti(Visum.Net.MainTurns,"AddVal1",[0]*Visum.Net.MainTurns.Count,False)
        SetMulti(Visum.Net.MainTurns,"AddVal2",[0]*Visum.Net.MainTurns.Count,False)
        SetMulti(Visum.Net.MainTurns,"AddVal3",[0]*Visum.Net.MainTurns.Count,False)
        
    if Visum.Net.Zones.Count > 0:
        SetMulti(Visum.Net.Zones,"AddVal1",[0]*Visum.Net.Zones.Count,False)
        SetMulti(Visum.Net.Zones,"AddVal2",[0]*Visum.Net.Zones.Count,False)
        SetMulti(Visum.Net.Zones,"AddVal3",[0]*Visum.Net.Zones.Count,False)
        
    if Visum.Net.MainZones.Count > 0:
        SetMulti(Visum.Net.MainZones,"AddVal1",[0]*Visum.Net.MainZones.Count,False)
        SetMulti(Visum.Net.MainZones,"AddVal2",[0]*Visum.Net.MainZones.Count,False)
        SetMulti(Visum.Net.MainZones,"AddVal3",[0]*Visum.Net.MainZones.Count,False)
    
    if Visum.Net.Connectors.Count > 0:
        SetMulti(Visum.Net.Connectors,"AddVal1",[0]*Visum.Net.Connectors.Count,False)
        SetMulti(Visum.Net.Connectors,"AddVal2",[0]*Visum.Net.Connectors.Count,False)
        SetMulti(Visum.Net.Connectors,"AddVal3",[0]*Visum.Net.Connectors.Count,False)
    
    if Visum.Net.Territories.Count > 0:
        SetMulti(Visum.Net.Territories,"AddVal1",[0]*Visum.Net.Territories.Count,False)
        SetMulti(Visum.Net.Territories,"AddVal2",[0]*Visum.Net.Territories.Count,False)
        SetMulti(Visum.Net.Territories,"AddVal3",[0]*Visum.Net.Territories.Count,False)
    
    SetMulti(Visum.Net.StopPoints,"AddVal1",[0]*Visum.Net.StopPoints.Count,False)
    SetMulti(Visum.Net.StopPoints,"AddVal2",[0]*Visum.Net.StopPoints.Count,False)
    SetMulti(Visum.Net.StopPoints,"AddVal3",[0]*Visum.Net.StopPoints.Count,False)
    
    SetMulti(Visum.Net.StopAreas,"AddVal1",[0]*Visum.Net.StopAreas.Count,False)
    SetMulti(Visum.Net.StopAreas,"AddVal2",[0]*Visum.Net.StopAreas.Count,False)
    SetMulti(Visum.Net.StopAreas,"AddVal3",[0]*Visum.Net.StopAreas.Count,False)
    
    SetMulti(Visum.Net.Stops,"AddVal1",[0]*Visum.Net.Stops.Count,False)
    SetMulti(Visum.Net.Stops,"AddVal2",[0]*Visum.Net.Stops.Count,False)
    SetMulti(Visum.Net.Stops,"AddVal3",[0]*Visum.Net.Stops.Count,False)
    
    SetMulti(Visum.Net.Lines,"AddVal1",[0]*Visum.Net.Lines.Count,False)
    SetMulti(Visum.Net.Lines,"AddVal2",[0]*Visum.Net.Lines.Count,False)
    SetMulti(Visum.Net.Lines,"AddVal3",[0]*Visum.Net.Lines.Count,False)
    
    SetMulti(Visum.Net.LineRoutes,"AddVal1",[0]*Visum.Net.LineRoutes.Count,False)
    SetMulti(Visum.Net.LineRoutes,"AddVal2",[0]*Visum.Net.LineRoutes.Count,False)
    SetMulti(Visum.Net.LineRoutes,"AddVal3",[0]*Visum.Net.LineRoutes.Count,False)
    
    SetMulti(Visum.Net.LineRouteItems,"AddVal",[0]*Visum.Net.LineRouteItems.Count,False)
    SetMulti(Visum.Net.TimeProfileItems,"AddVal",[0]*Visum.Net.TimeProfileItems.Count,False)
    
    SetMulti(Visum.Net.VehicleJourneys,"AddVal1",[0]*Visum.Net.VehicleJourneys.Count,False)
    SetMulti(Visum.Net.VehicleJourneys,"AddVal2",[0]*Visum.Net.VehicleJourneys.Count,False)
    SetMulti(Visum.Net.VehicleJourneys,"AddVal3",[0]*Visum.Net.VehicleJourneys.Count,False)
    
    if Visum.Net.CountLocations.Count > 0:
        SetMulti(Visum.Net.CountLocations,"AddVal1",[0]*Visum.Net.CountLocations.Count,False)
        SetMulti(Visum.Net.CountLocations,"AddVal2",[0]*Visum.Net.CountLocations.Count,False)
        SetMulti(Visum.Net.CountLocations,"AddVal3",[0]*Visum.Net.CountLocations.Count,False)
        SetMulti(Visum.Net.CountLocations,"AddVal4",[0]*Visum.Net.CountLocations.Count,False)
        SetMulti(Visum.Net.CountLocations,"AddVal5",[0]*Visum.Net.CountLocations.Count,False)
    
    Visum.Log(20480,_("All AddVals set to 0!"))

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
        Run()
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
