# -*- coding: utf-8 -*-

import sys
import pandas as pd
import wx
import win32com.client
import pythoncom
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext


def Run(param):
    if param["Proc"] in ["setParameters", "PerformanceStatement"]:
        Visum.Net.SetAttValue("TN", param["TN"])
        Visum.Net.SetAttValue("S_WOCHEN", param["SW"])
        Visum.Net.SetAttValue("F_WOCHEN", param["FW"])
        Visum.Net.SetAttValue("EFW", param["EFW"])
        Visum.Net.SetAttValue("M_FAHRZEUGBEDARF", param["FB"])
        Visum.Net.SetAttValue("ANABOVERLAP", param["UEL"])
    if param["Proc"] == "setParameters":
        addIn.ReportMessage(_("Parameters: set"),2 )
    if param["Proc"] == "PerformanceStatement":
        # alle Verfahren deaktivieren
        for o in Visum.Procedures.Operations.GetAll:
            o.SetAttValue("ACTIVE", False)
        # generiere Start- und End-Index von LAR
        index_LAR_start = [o.AttValue("COMMENT") for o in Visum.Procedures.Operations.GetAll].index("TNM Leistungsabrechnung")
        index_LAR_end = len(Visum.Procedures.Operations.GetAll) # falls TNM Leistungsabrechnung am Ende steht.
        for o in Visum.Procedures.Operations.GetAll[index_LAR_start + 1:]:
            if not Visum.Procedures.Operations.GetParent(o):
                index_LAR_end = int(o.AttValue("NO") - 1)
        # nur Verfahren zur LAR aktivieren
        for o in Visum.Procedures.Operations.GetAll[index_LAR_start:index_LAR_end]:
            o.SetAttValue("ACTIVE", True)
        Visum.Procedures.Execute()
        # Schaue, ob in den Verfahren Fehler auftauchen (18 = 19 Verfahrensschritte (start bei Index 0))
        n_errors = int(sum([o.AttValue("ERRORCOUNT") or 0 for o in Visum.Procedures.Operations.GetAll][index_LAR_start:index_LAR_end]))
        if n_errors > 0: # wenn Fehler, dann kein Excel-Export, sondern Hinweis
            addIn.ReportMessage(_("%s errors occurred, please check the messages.") %(str(n_errors)))
        else:
            on_excel_write(param["TN"])
            addIn.ReportMessage(_("Service billing: completed"), 2)
    if param["Proc"] == "openLAR":
        on_excel_write(param["TN"])
        addIn.ReportMessage(_("Finished"), 2)
    
    
def on_excel_write(TN):
        pythoncom.CoInitialize()  # important inside GUI event
        xls = _get_excel()
        werte = _get_abrechnungs_werte(TN)
        top_left = xls.Range("A1")
        bottom_right = top_left.Offset(len(werte), len(werte[0]))
        write_range = xls.Range(top_left, bottom_right)
        write_range.Value = tuple(tuple(w) for w in werte)
    
        # Formatierung
        xls.Columns.AutoFit()
        xls.Rows(1).Font.Bold = True
        letter = chr(ord('A') + len(werte[0]))
        xls.Columns(f"E:{letter}").NumberFormat = "#.##0"
        

def _get_abrechnungs_werte(TN):
    attribute = ["NO", "CODE", "NAME", "RANG", f"{TN}_GESAMTKOSTEN"]
    for i in Visum.Net.VehicleCombinations.Attributes.GetAll:
        if TN in i.Name and "GESAMTKOSTEN" in i.Name:
            if i.Name == f"{TN}_GESAMTKOSTEN": continue
            attribute.append(i.Name)
    
    df_fahrzeugkombinationen = pd.DataFrame(Visum.Net.VehicleCombinations.GetMultipleAttributes(attribute, True),
                            columns = attribute)
    df_fahrzeugkombinationen = df_fahrzeugkombinationen.where(pd.notnull(df_fahrzeugkombinationen), None)
    rows_fahrzeugkombinationen = df_fahrzeugkombinationen.values.tolist()
    rows_fahrzeugkombinationen.insert(0, list(df_fahrzeugkombinationen.columns))
    return rows_fahrzeugkombinationen


def _get_excel():
        """Attach to running Excel or create new"""
        try:# running
            xl = win32com.client.GetActiveObject("Excel.Application")
        except:# new
            xl = win32com.client.Dispatch("Excel.Application")
            xl.Visible = True
            xl.UserControl = True
        wb = xl.Workbooks.Add()
        win = xl.Windows(wb.Name)
        win.Visible = True
        ws = wb.ActiveSheet
        return ws


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
        defaultParam = {"Proc" : False}
        param = addInParam.Check(True, defaultParam)
        Run(param)
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
