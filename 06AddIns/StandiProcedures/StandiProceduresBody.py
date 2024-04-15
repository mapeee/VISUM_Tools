# -*- coding: utf-8 -*-

import numpy as np
import sys
import wx
import CalVal
import CalNKV
import CalPuTService
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext

def Run(param):
    Visum.Filters.InitAll()
    
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
        if param["PrT_pcput"]:
            Visum.Procedures.Operations.ItemByKey(operations_code.index("8")+1).SetAttValue("ACTIVE",True)
            for i in Visum.Procedures.Operations.GetChildren((Visum.Procedures.Operations.ItemByKey(operations_code.index("8")+1))):
                if i.AttValue("OPERATIONTYPE") == 100:
                    i.TimetableBasedParameters.BaseParameters.SetAttValue("CONNECTIONSOURCETYPE","ConnCalculated")
                    break
            
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
        for i in Visum.Procedures.Operations.GetChildren((Visum.Procedures.Operations.ItemByKey(operations_code.index("8")+1))):
            if i.AttValue("OPERATIONTYPE") == 100:
                i.TimetableBasedParameters.BaseParameters.SetAttValue("CONNECTIONSOURCETYPE","ConnFromFile")
                break
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
            if np.any(np.array(Visum.Net.UserDefinedGroups.GetMultiAttValues("Name"))[:,1] == _("Standi-Verfahren")) == False:
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
        Visum.Log(20480,_("Delete interim steps: finished!"))
        
    if param["NKV"] in ["GoPuT_op_bc", "GoPuT_op_pc"]:
        if Visum.Net.VehicleCombinations.Count == 0:
            addIn.ReportMessage(_("VehicleCombinations are missing!"))
            return False
        for op in Visum.Procedures.Operations.GetAll:
            if op.AttValue("OPERATIONTYPE") == 36:
                Visum.Procedures.OperationExecutor.SetCurrentOperation(op)
                Visum.Procedures.OperationExecutor.ExecuteCurrentOperation()
                break
        if not Put_op({"GoPuT_op_bc" : True, "GoPuT_op_pc" : False}[param["NKV"]]): return
        
    if param["NKV"] == "GoNKV":
        if not NKV(): return
        
    if param["NKV"] == "GoDel_NKV":
        Del_NKV() 
        Visum.Log(20480,_("Deleted all NKV components!"))
        
    if param["NKV"] == "GoImport_NKV":
        Import_NKV() 
        Visum.Log(20480,_("Imported all NKV components!"))
        
        
    Visum.Log(20480,_("Standi procedures: finished!"))

def Del_NKV():
    for i in ["Standi-NKV", "Standi-Szenario", "Standi-Energie", "Standi-Parameter"]:
        if np.any(np.array(Visum.Net.TableDefinitions.GetMultipleAttributes(["NAME"])) == i) == True:
            Visum.Net.RemoveTableDefinition(Visum.Net.TableDefinitions.ItemByKey(i))

    for c in ["_O", "_M"]:
        for i in ["FAHRTEN", "PERSONALK", "UKLL", "ENERGIEK", "LL"]:
            if Visum.Net.Lines.AttrExists(i+c): Visum.Net.Lines.DeleteUserDefinedAttribute(i+c)
        for i in ["ENERGIEV"]:
            if Visum.Net.VehicleCombinations.AttrExists(i+c): Visum.Net.VehicleCombinations.DeleteUserDefinedAttribute(i+c)
        for i in ["FZG", "FAHRZEUGK", "LL"]:
            if Visum.Net.VehicleUnits.AttrExists(i+c): Visum.Net.VehicleUnits.DeleteUserDefinedAttribute(i+c)

def Import_NKV():
    if np.any(np.array(Visum.Net.UserDefinedGroups.GetMultiAttValues("Name"))[:,1] == _("Standi-Verfahren")) == False:
        Visum.Net.AddUserDefinedGroup(_("Standi-Verfahren"))
    Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_NKV.accdb", True)
    Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_NKVEntries.accdb", True)
    Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_Energie.accdb", True)
    Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_EnergieEntries.accdb", True)
    Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_Parameter.accdb", True)
    Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_ParameterEntries.accdb", True)
    Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_Scenario.accdb", True)
    Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_ScenarioEntries.accdb", True)
    Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\PuTService_bc.accdb", True)
    Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\PuTService_pc.accdb", True)

def NKV():
    Visum.Log(20480,_("NKV Calculation: starting!"))
    if np.any(np.array(Visum.Net.UserDefinedGroups.GetMultiAttValues("Name"))[:,1] == _("Standi-Verfahren")) == False:
        Visum.Net.AddUserDefinedGroup(_("Standi-Verfahren"))
    if "Standi-NKV" not in [i.AttValue("NAME") for i in Visum.Net.TableDefinitions.GetAll]:
        Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_NKV.accdb", True)
        Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_NKVEntries.accdb", True)
    if "Standi-Energie" not in [i.AttValue("NAME") for i in Visum.Net.TableDefinitions.GetAll]:
        Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_Energie.accdb", True)
        Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_EnergieEntries.accdb", True)
    if "Standi-Parameter" not in [i.AttValue("NAME") for i in Visum.Net.TableDefinitions.GetAll]:
        Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_Parameter.accdb", True)
        Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_ParameterEntries.accdb", True)
    if "Standi-Szenario" not in [i.AttValue("NAME") for i in Visum.Net.TableDefinitions.GetAll]:
        Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_Scenario.accdb", True)
        Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_ScenarioEntries.accdb", True)
        Visum.Log(12288,_("Table 'Standi-Szenario' was created but is empty!"))
        return False

    if not CalNKV.calc(Visum, Visum.Net.TableDefinitions.ItemByKey("Standi-NKV")): return False
    
    Visum.Log(20480,_("NKV calculation: finished!"))
    return True

def Put_op(bc):
    Visum.Log(20480,_("PuT Service %s: Calculating!") %({True : _("Base case"), False : _("Planning case")}[bc]))
    for i in ["LEERMASSE", "UKLL", "UKT", "UKTKM", "EVS", "ENERGIE", "ANSCHAFFUNGSK", "THG"]:
        if not Visum.Net.VehicleUnits.AttrExists(i):
            Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\PuTVehAttributes.accdb", True)
            addIn.ReportMessage(_("UDA %s for VehicleUnits is empty/mising!")%(i))
            return False
        
    if "Standi-Energie" not in [i.AttValue("NAME") for i in Visum.Net.TableDefinitions.GetAll]:
        Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_Energie.accdb", True)
        Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_EnergieEntries.accdb", True)
    if "Standi-Parameter" not in [i.AttValue("NAME") for i in Visum.Net.TableDefinitions.GetAll]:
        Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_Parameter.accdb", True)
        Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\Standi_ParameterEntries.accdb", True)
        
    v, v_db = "O", "bc"
    if bc == False: v, v_db = "M", "pc"
    for i in ["FZG_", "FAHRZEUGK_", "LL_"]:
        if not Visum.Net.VehicleUnits.AttrExists(i+v):
            Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\PuTService_"+v_db+".accdb", True)
            break
    for i in ["ENERGIEV_"]:
        if not Visum.Net.VehicleCombinations.AttrExists(i+v):
            Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\PuTService_"+v_db+".accdb", True)
            break
    for i in ["FAHRTEN_", "PERSONALK_", "UKLL_", "ENERGIEK_", "LL_"]:
        if not Visum.Net.Lines.AttrExists(i+v):
            Visum.IO.LoadAccessDatabase(addIn.DirectoryPath + "Data\\PuTService_"+v_db+".accdb", True)
            break         

    CalPuTService.VehicleUnits(Visum,bc)
    
    if not CalPuTService.PuTCosts(Visum,bc):
        return False
    
    Visum.Log(20480,_("PuT operating values: calculated!"))
    return True

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
