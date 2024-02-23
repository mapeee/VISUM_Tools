# -*- coding: utf-8 -*-

import numpy as np
import sys
import wx
import CalVal
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext

def Run(param):
    
    operations = Visum.Procedures.Operations.GetAll
    operations_code = [o.AttValue("CODE") for o in operations]
    
    if True in [param["PrT_bc"], param["PuT_bc"], param["PrT_pc"], param["PrT_pcput"],
                param["PuT_pc"], param["PuT_pcprt"]]:
        for i in operations:
            i.SetAttValue("ACTIVE",False)
            if Visum.Procedures.Operations.GetParent(i) is not None: i.SetAttValue("ACTIVE",True)
            Visum.Procedures.Operations.ItemByKey(1).SetAttValue("ACTIVE",True)
            Visum.Procedures.Operations.ItemByKey(6).SetAttValue("ACTIVE",True)

    if param["PrT_bc"]:
        Visum.Procedures.Operations.ItemByKey(2).SetAttValue("ACTIVE",True)
        Visum.Procedures.Operations.ItemByKey(3).SetAttValue("ACTIVE",True)
        Visum.Procedures.Operations.ItemByKey(operations_code.index("1")+1).SetAttValue("ACTIVE",True)

    if param["PrT_pc"] or param["PrT_pcput"]:
        Visum.Procedures.Operations.ItemByKey(4).SetAttValue("ACTIVE",True)
        Visum.Procedures.Operations.ItemByKey(5).SetAttValue("ACTIVE",True)
        Visum.Procedures.Operations.ItemByKey(operations_code.index("2")+1).SetAttValue("ACTIVE",True)
        Visum.Procedures.Operations.ItemByKey(operations_code.index("3")+1).SetAttValue("ACTIVE",True)
        if param["PrT_pcput"]: Visum.Procedures.Operations.ItemByKey(operations_code.index("8")+1).SetAttValue("ACTIVE",True)
    
    if param["PuT_bc"]:
        Visum.Procedures.Operations.ItemByKey(2).SetAttValue("ACTIVE",True)
        Visum.Procedures.Operations.ItemByKey(3).SetAttValue("ACTIVE",True)
        Visum.Procedures.Operations.ItemByKey(operations_code.index("4")+1).SetAttValue("ACTIVE",True)

    if param["PuT_pc"] or param["PuT_pcprt"]:
        Visum.Procedures.Operations.ItemByKey(4).SetAttValue("ACTIVE",True)
        Visum.Procedures.Operations.ItemByKey(5).SetAttValue("ACTIVE",True)
        Visum.Procedures.Operations.ItemByKey(operations_code.index("5")+1).SetAttValue("ACTIVE",True)
        Visum.Procedures.Operations.ItemByKey(operations_code.index("6")+1).SetAttValue("ACTIVE",True)
        if param["PuT_pcprt"]: Visum.Procedures.Operations.ItemByKey(operations_code.index("7")+1).SetAttValue("ACTIVE",True)
        Visum.Procedures.Operations.ItemByKey(operations_code.index("8")+1).SetAttValue("ACTIVE",True)
        Visum.Procedures.Operations.ItemByKey(operations_code.index("9")+1).SetAttValue("ACTIVE",True)
    
    if param["PuTCon"]:
        for i in Visum.Procedures.Operations.GetChildren((Visum.Procedures.Operations.ItemByKey(operations_code.index("5")+1))):
            if i.AttValue("OPERATIONTYPE") == 102:
                i.TimetableBasedParameters.ConnExportParameters.SetAttValue("CONNEXPORTFILENAME",param["PuTConPath"])
                break
        for i in Visum.Procedures.Operations.GetChildren((Visum.Procedures.Operations.ItemByKey(operations_code.index("8")+1))):
            if i.AttValue("OPERATIONTYPE") == 100:
                i.TimetableBasedParameters.BaseParameters.SetAttValue("CONNIMPORTFILENAME",param["PuTConPath"])
                break
                
    if True in [param["PrT_bc"], param["PuT_bc"], param["PrT_pc"], param["PuT_pc"], param["PrT_pcput"], param["PuT_pcprt"]]:
        if param["ExeProc"]: Visum.Procedures.Execute()

    
    if param["Val_bc"] or param["Val_pc"]:
        if param["Val_bc"]:
            if _("Standi-Verfahren") not in np.array(Visum.Net.UserDefinedGroups.GetMultiAttValues("Name"))[:,1]:
                Visum.Net.AddUserDefinedGroup(_("Standi-Verfahren"))
            if "Ohnefall-Mitfall-Vergleich" not in [i.AttValue("NAME") for i in Visum.Net.TableDefinitions.GetAll]:
                Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\base_planning_case.accdb", True)
            for i in ["CODE","OHNEFALL","MITFALL","KENNWERT","DIFF_ABS","DIFF_REL"]:
                if i not in [e.ID for e in Visum.Net.TableDefinitions.ItemByKey("Ohnefall-Mitfall-Vergleich").TableEntries.Attributes.GetAll]:
                    Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\base_planning_case.accdb", True)
            if Visum.Net.TableDefinitions.ItemByKey("Ohnefall-Mitfall-Vergleich").TableEntries.Count != 22:
                Visum.Net.TableDefinitions.ItemByKey("Ohnefall-Mitfall-Vergleich").TableEntries.RemoveAll()
                Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\base_planning_caseEntries.accdb", True)
            
            if CalVal.calc(Visum,"Ohnefall"):
                Visum.Log(20480,_("Base case values: calculated!"))
        else:
            if "Ohnefall-Mitfall-Vergleich" not in [i.AttValue("NAME") for i in Visum.Net.TableDefinitions.GetAll]:
                Visum.Log(12288,_("Table 'Ohnefall-Mitfall-Vergleich' is missing!"))
                return
            if CalVal.calc(Visum,"Mitfall"):
                Visum.Log(20480,_("Planning case values: calculated!"))

    if param["Del"]:
        for CODE in ["TTC","ACT","EGT","TWT","NTR","JRT","PJT","WKT","XIMP","Rsv_O","Rsv_M",
                     "SV","SFQ","IPD","IVD","Gesamtnachfrage_O"]:
            Visum.Net.Matrices.ItemsByRef('Matrix([CODE] = \"%s")'%(CODE)).RemoveAll()   
        Visum.Log(20480,_("Delete interim steps: finished"))

    Visum.Log(20480,_("Standi procedures: finished"))

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
        defaultParam = {"Replace" : False}
        param = addInParam.Check(True, defaultParam)
        Run(param)
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
