# -*- coding: utf-8 -*-

import numpy as np
import re
import sys
import wx
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext


def Run(param):  
    try:
        checkNetwork(param["NetworkType"], param["POICat"], param["POIAttr"])
    except Exception as e:
        addIn.ReportMessage(_("%s %s (POI %s) not existing!") %(_(param["NetworkType"])[:-1],str(e.args[0]),str(e.args[1])))
        return
    
    OP = Visum.Procedures.Operations.AddOperation(1)
    OP.SetAttValue("OPERATIONTYPE",65)
    CODE = \
f'import re\n\
elements = [val for _, val in Visum.Net.{param["NetworkType"]}s.GetMultiAttValues("NO")]\n\
for poi in Visum.Net.POICategories.ItemByKey({param["POICat"]}).POIs.GetAllActive:\n\
    if {param["Replace"]}:\n\
        for element in poi.POITo{param["NetworkType"]}Items.GetAll:\n\
            poi.RemovePOITo{param["NetworkType"]}Item(element.AttValue("{param["NetworkType"]}NO"))\n\
    for i in re.split(r",|;",str(poi.AttValue("{param["POIAttr"]}"))):\n\
        try: i = int(float(i))\n\
        except:continue\n\
        if i not in elements: continue\n\
        if i == 0: continue\n\
        for element in poi.POITo{param["NetworkType"]}Items.GetAll:\n\
            if i == element.AttValue("{param["NetworkType"]}NO"):\n\
                poi.RemovePOITo{param["NetworkType"]}Item(i)\n\
        poi.AddPOITo{param["NetworkType"]}Item(Visum.Net.{param["NetworkType"]}s.ItemByKey(i))\n\
Visum.Log(20480,"POI2Network: finished")'
    OP.ExecuteScriptParameters.SetAttValue("INTERNALSCRIPTCODE",CODE)
    OP.ExecuteScriptParameters.SetAttValue("USEINTERNALSCRIPTCODE",True)
    Visum.Procedures.OperationExecutor.SetCurrentOperation(OP)
    Visum.Procedures.OperationExecutor.ExecuteCurrentOperation()
    Visum.Procedures.Operations.RemoveOperation(1)

def checkNetwork(NetworkType, POICat, POIAttr):
    if NetworkType == "Link":
        arr = np.array(Visum.Net.Links.GetMultiAttValues("No")).astype("int")[:,1]
    if NetworkType == "Node":
        arr = np.array(Visum.Net.Nodes.GetMultiAttValues("No")).astype("int")[:,1]
    if NetworkType == "StopPoint":
        arr = np.array(Visum.Net.StopPoints.GetMultiAttValues("No")).astype("int")[:,1]
    if NetworkType == "StopArea":
        arr = np.array(Visum.Net.StopAreas.GetMultiAttValues("No")).astype("int")[:,1]
    if NetworkType == "Zone":
        arr = np.array(Visum.Net.Zones.GetMultiAttValues("No")).astype("int")[:,1]
    
    for poi in Visum.Net.POICategories.ItemByKey(POICat).POIs.GetAllActive:
        for i in re.split(r',|;',str(poi.AttValue(POIAttr))):
            try: int(i)
            except: continue
            if np.any(arr==int(i)) == False and int(i) >0:
                raise Exception(i, int(poi.AttValue("NO")))

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
        defaultParam = {"POICat" : 0, "POIAttr" : "...", "NetworkType" : None}
        param = addInParam.Check(True, defaultParam)
        Run(param)
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
