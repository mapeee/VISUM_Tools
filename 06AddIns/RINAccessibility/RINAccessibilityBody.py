# -*- coding: utf-8 -*-
import bisect
import numpy as np
from osgeo import ogr
import sys
import wx
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
from VisumPy.helpers import SetMulti
_ = AddIn.gettext

def Run(param):  
    if param["Proce"] == "PT_SAQ":
        if not CheckStopAreas(): return
        _PT_UDA()
        
        #Isochrone Parameter
        basePara = Visum.Analysis.Isochrones.CreatePuTIsochroneBaseParameters()
        basePara.SetAttValue("ONLYACTIVEVEHJOURNEYSECTIONS", 1)
        basePara.SetMode(_mode())
        requestPara = Visum.Analysis.Isochrones.CreatePuTIsochroneRequestParameters()
        requestPara.SetAttValue("SEARCHINTERVALSTARTTIME","12:00:00")
        requestPara.SetAttValue("SEARCHINTERVALENDTIME","16:00:00")
        requestPara.SetAttValue("EXTENSIONTIME", 900) #15 minutes
        requestPara.SetAttValue("TRANSFERFACTOR", 4)
    
        if Visum.Net.Marking.ObjectType == 14 and len(param["Marking"]) > 0:
            for from_stop in param["Marking"]:
                from_stop = Visum.Net.StopAreas.ItemByKey(from_stop)
                SAQ(from_stop, basePara, requestPara)
        else:
            for from_stop in Visum.Net.StopAreas.GetAllActive:
                SAQ(from_stop, basePara, requestPara)
    
    Visum.Log(20480,_("RIN Accessibility: finished"))
   
def CheckStopAreas():
    if Visum.Net.Marking.ObjectType == 14 and len(param["Marking"]) > 0:
        return True
    if len(Visum.Net.StopAreas.GetAllActive) > 50:
        addIn.ReportMessage(_("To many (> 50) active StopAreas!"))
        return False
    return True
   
def SAQ(_from_stop, _basePara, _requestPara):
    def _error():
        SAQ_OV.append(_stop.AttValue("SAQ_OV"))
        SAQ_RZ.append(_stop.AttValue("SAQ_RZ"))
        SAQ_UH.append(_stop.AttValue("SAQ_UH"))
        SAQ_VLUFT.append(_stop.AttValue("SAQ_VLUFT"))
        SAQ_LL.append(_stop.AttValue("SAQ_LL"))
        SAQ_HstBer.append(_stop.AttValue("SAQ_HstBer"))
    
    SAQ_OV, SAQ_RZ, SAQ_UH, SAQ_VLUFT, SAQ_LL, SAQ_HstBer = [], [], [], [], [], []
    
    _from_stopxy = ogr.CreateGeometryFromWkt(_from_stop.AttValue("WKTLOCWGS84"))
    
    NE = Visum.CreateNetElements()
    NE.Add(_from_stop)
    Visum.Analysis.Isochrones.ExecutePuTWithParameterObjects(NE, _basePara, _requestPara)

    for _stop in Visum.Net.StopAreas.GetAll:
        UH = _stop.AttValue("ISOCTRANSFERSPUT")
        RZ = _stop.AttValue("ISOCTIMEPUT") / 60
        RZ_gew = RZ + 11 #access and egress 4 min / initial waiting 3 min
        LLge = _from_stopxy.Distance(ogr.CreateGeometryFromWkt(_stop.AttValue("WKTLOCWGS84"))) * 111.111111
        try:
            VLuft = 60 / RZ_gew * LLge
        except:
            _error()
            continue

        if RZ > 500 or VLuft < _stop.AttValue("SAQ_VLUFT") or LLge == 0:
            _error()
            continue

        SAQ_v = [1/(0.40*pow(LLge,-0.5)+0.0063),
                 1/(0.32*pow(LLge,-0.5)+0.0052),
                 1/(0.26*pow(LLge,-0.5)+0.0044),
                 1/(0.22*pow(LLge,-0.5)+0.0037),
                 1/(0.19*pow(LLge,-0.5)+0.0031)]
        
        index = bisect.bisect_left(SAQ_v, VLuft)
        SAQ_OV.append(["F","E","D","C","B","A"][index])
        SAQ_RZ.append(RZ)
        SAQ_UH.append(int(UH))
        SAQ_VLUFT.append(VLuft)
        SAQ_LL.append(LLge)
        SAQ_HstBer.append(int(NE.GetAll[0].AttValue("NO")))
            
    SetMulti(Visum.Net.StopAreas,"SAQ_OV",SAQ_OV,False)
    SetMulti(Visum.Net.StopAreas,"SAQ_RZ",SAQ_RZ,False)
    SetMulti(Visum.Net.StopAreas,"SAQ_UH",SAQ_UH,False)
    SetMulti(Visum.Net.StopAreas,"SAQ_VLUFT",SAQ_VLUFT,False)
    SetMulti(Visum.Net.StopAreas,"SAQ_LL",SAQ_LL,False)
    SetMulti(Visum.Net.StopAreas,"SAQ_HstBer",SAQ_HstBer,False)

    Visum.Analysis.Isochrones.Clear() 

def _mode():
    for VSys in Visum.Net.TSystems.GetAll:
        if VSys.AttValue("TYPE") == "PUT":
            return VSys.AttValue(r"MIN:MODES\CODE")
        
def _PT_UDA():
    if np.any(np.array(Visum.Net.UserDefinedGroups.GetMultiAttValues("Name"))[:,1] == _("RIN Accessibility")) == False:
        Visum.Net.AddUserDefinedGroup(_("RIN Accessibility"))
    if Visum.Net.StopAreas.AttrExists("SAQ_OV") is False:
        Visum.Net.StopAreas.AddUserDefinedAttribute("SAQ_OV","SAQ_OV","SAQ_OV",5,DefVal="Z")
        Visum.Net.StopAreas.Attributes.ItemByKey("SAQ_OV").Comment = _("SAQ RIN Public Transport (Levels of quality)")
        Visum.Net.StopAreas.Attributes.ItemByKey("SAQ_OV").UserDefinedGroup = _("RIN Accessibility")
    if Visum.Net.StopAreas.AttrExists("SAQ_RZ") is False:
        Visum.Net.StopAreas.AddUserDefinedAttribute("SAQ_RZ","SAQ_RZ","SAQ_RZ",2,1,DefVal=0)
        Visum.Net.StopAreas.Attributes.ItemByKey("SAQ_RZ").Comment = _("SAQ RIN PT-TravelTime (Levels of quality)")
        Visum.Net.StopAreas.Attributes.ItemByKey("SAQ_RZ").UserDefinedGroup = _("RIN Accessibility")
    if Visum.Net.StopAreas.AttrExists("SAQ_UH") is False:
        Visum.Net.StopAreas.AddUserDefinedAttribute("SAQ_UH","SAQ_UH","SAQ_UH",2,0,DefVal=0)
        Visum.Net.StopAreas.Attributes.ItemByKey("SAQ_UH").Comment = _("SAQ RIN PT-Transers (Levels of quality)")
        Visum.Net.StopAreas.Attributes.ItemByKey("SAQ_UH").UserDefinedGroup = _("RIN Accessibility")
    if Visum.Net.StopAreas.AttrExists("SAQ_LL") is False:
        Visum.Net.StopAreas.AddUserDefinedAttribute("SAQ_LL","SAQ_LL","SAQ_LL",2,3,DefVal=0)
        Visum.Net.StopAreas.Attributes.ItemByKey("SAQ_LL").Comment = _("SAQ RIN DirectDistance (Levels of quality)")
        Visum.Net.StopAreas.Attributes.ItemByKey("SAQ_LL").UserDefinedGroup = _("RIN Accessibility")
    if Visum.Net.StopAreas.AttrExists("SAQ_VLUFT") is False:
        Visum.Net.StopAreas.AddUserDefinedAttribute("SAQ_VLUFT","SAQ_VLUFT","SAQ_VLUFT",2,3,DefVal=0)
        Visum.Net.StopAreas.Attributes.ItemByKey("SAQ_VLUFT").Comment = _("SAQ RIN DirectSpeed (Levels of quality)")
        Visum.Net.StopAreas.Attributes.ItemByKey("SAQ_VLUFT").UserDefinedGroup = _("RIN Accessibility")
    if Visum.Net.StopAreas.AttrExists("SAQ_HstBer") is False:
        Visum.Net.StopAreas.AddUserDefinedAttribute("SAQ_HstBer","SAQ_HstBer","SAQ_HstBer",2,0,DefVal=0)
        Visum.Net.StopAreas.Attributes.ItemByKey("SAQ_HstBer").Comment = _("SAQ RIN Start StopArea  (Levels of quality)")
        Visum.Net.StopAreas.Attributes.ItemByKey("SAQ_HstBer").UserDefinedGroup = _("RIN Accessibility")
    SetMulti(Visum.Net.StopAreas,"SAQ_OV",["Z"]*Visum.Net.StopAreas.Count,False)
    SetMulti(Visum.Net.StopAreas,"SAQ_RZ",[0]*Visum.Net.StopAreas.Count,False)
    SetMulti(Visum.Net.StopAreas,"SAQ_UH",[0]*Visum.Net.StopAreas.Count,False)
    SetMulti(Visum.Net.StopAreas,"SAQ_LL",[0]*Visum.Net.StopAreas.Count,False)
    SetMulti(Visum.Net.StopAreas,"SAQ_VLUFT",[0]*Visum.Net.StopAreas.Count,False)
    SetMulti(Visum.Net.StopAreas,"SAQ_HstBer",[0]*Visum.Net.StopAreas.Count,False)

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
        param = addInParam.Check(True,{"Proce" : "...", "Marking" : False})
        Run(param)
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
