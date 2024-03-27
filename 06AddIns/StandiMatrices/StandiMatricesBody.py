# -*- coding: utf-8 -*-

import numpy as np
import sys
import wx
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext


def Run(param):  
    for i in param["Matrices_b"]+param["Matrices_p"]+param["Matrices_s"]+param["Matrices_f"]:
        i = i.split(" : ")
        No = int(i[2])
        CODE = i[0]
        NAME = i[1]
        
        try: checkNoCode(No, CODE)
        except Exception as e:
            addIn.ReportMessage(_("%s already exists!") %e)
            return
        
        Add(No, CODE, NAME, param["Replace"])

    Visum.Log(20480,_("Standi Matrices: created!"))

def Add(_No, _CODE, _NAME, Replace=False):
    if Visum.Net.Matrices.Count > 0:
        matNo = np.array(Visum.Net.Matrices.GetMultiAttValues("No")).astype("int")[:,1]
        if Replace == False and _No in matNo:
                return
        if _No in matNo:
            mat = Visum.Net.Matrices.ItemByKey(_No)
            if mat.AttValue("DATASOURCETYPE") == "FORMULA":
                Visum.Net.RemoveMatrix(mat)
            else:
                mat.SetValuesToResultOfFormula("0")
                return
            
    if _No in [4,5,61,62]: 
        if _No == 4: m = Visum.Net.AddMatrixWithFormula(4,'Matrix([CODE] = "Pkw_E_O")*1,3',2,3)
        elif _No == 5: m = Visum.Net.AddMatrixWithFormula(5,'Matrix([CODE] = "Pkw_E_M")*1,3',2,3)
        elif _No == 61: m = Visum.Net.AddMatrixWithFormula(61,'(Matrix([CODE] = "OEV_Gesamtwiderstand_M")-Matrix([CODE] = "OEV_Gesamtwiderstand_O"))*((Matrix([CODE] = "OEV_E_O")+Matrix([CODE] = "OEV_E_M"))/2)/60',2,4)
        else: m = Visum.Net.AddMatrixWithFormula(62,'(Matrix([CODE] = "OEV_Gesamtwiderstand_M")-Matrix([CODE] = "OEV_Gesamtwiderstand_O"))*Matrix([CODE] = "OEV_S")/60',2,4)
    else:
        if _No < 30: m = Visum.Net.AddMatrix(_No,2,3)
        else: m = Visum.Net.AddMatrix(_No,2,4)

    m.SetAttValue("Code",_CODE)
    m.SetAttValue("Name",_NAME)

def checkNoCode(_No, _CODE):
    for i in Visum.Net.Matrices:
        if i.AttValue("Code")==_CODE and int(i.AttValue("NO"))!=_No:
            raise Exception(_("CODE %s (%s)") %(_CODE,int(i.AttValue("NO"))))
            
        if int(i.AttValue("NO"))==_No and i.AttValue("Code")!=_CODE:
            raise Exception(_("NO %s (%s)") %(_No,i.AttValue("Code")))

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
        defaultParam = {"Matrices_b" : "...", "Items_b" : "...", "Matrices_p" : "...",
                        "Items_p" : "...", "Matrices_s" : "...", "Items_s" : "...",
                        "Matrices_f" : "...", "Items_f" : "...", "Replace" : False}
        param = addInParam.Check(True, defaultParam)
        Run(param)
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
