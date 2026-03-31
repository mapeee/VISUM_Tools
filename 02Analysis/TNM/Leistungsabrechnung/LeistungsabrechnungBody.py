# -*- coding: utf-8 -*-

import sys
import pandas as pd
import wx
import win32com.client
import pythoncom
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext


def Run(param):
    Visum.Net.SetAttValue("TN", param["TN"])
    if param["Proc"] != "setParameters" and not Visum.Net.VehicleJourneySections.CountActive:
        addIn.ReportMessage(_("No active VehicleJourneys"))
        return False
    if param["Proc"] in ["setParameters", "PerformanceStatement"]:
        Visum.Net.SetAttValue("S_WOCHEN", param["SW"])
        Visum.Net.SetAttValue("F_WOCHEN", param["FW"])
        Visum.Net.SetAttValue("EFW", param["EFW"])
        Visum.Net.SetAttValue("M_FAHRZEUGBEDARF", param["FB"])
        Visum.Net.SetAttValue("ANABOVERLAP", param["UEL"])
    if param["Proc"] == "setParameters":
        addIn.ReportMessage(_("Parameters: set"), 2)
    if param["Proc"] == "PerformanceStatement":
        # alle Verfahren deaktivieren
        for o in Visum.Procedures.Operations.GetAll:
            o.SetAttValue("ACTIVE", False)
            if "TNM" in o.AttValue("CODE"):
                o.SetAttValue("ACTIVE", True)
        Visum.Procedures.Execute()
        # Schaue, ob in den Verfahren Fehler auftauchen (18 = 19 Verfahrensschritte (start bei Index 0)) (über Start- und End-Index von LAR)
        index_LAR_start = [o.AttValue("COMMENT") for o in Visum.Procedures.Operations.GetAll].index("TNM Leistungsabrechnung")
        index_LAR_end = [o.AttValue("COMMENT") for o in Visum.Procedures.Operations.GetAll].index("TNM Endabrechnung")
        n_errors = int(sum([o.AttValue("ERRORCOUNT") or 0 for o in Visum.Procedures.Operations.GetAll][index_LAR_start:index_LAR_end+1]))
        if n_errors > 0: # wenn Fehler, dann kein Excel-Export, sondern Hinweis
            addIn.ReportMessage(_("%s errors occurred, please check the messages.") %(str(n_errors)))
        else:
            on_excel_LAR(param["TN"])
            addIn.ReportMessage(_("Service billing: completed"), 2)
    if param["Proc"] == "initEFW":
        # alle Verfahren deaktivieren
        for o in Visum.Procedures.Operations.GetAll:
            o.SetAttValue("ACTIVE", False)
            if o.AttValue("CODE") in ["TNM", "TNM_BEGINN", "TNM_FILTER"]:
                o.SetAttValue("ACTIVE", True)
        Visum.Procedures.Execute()
        addIn.ReportMessage(_("EFW for LineRouteItems of SN: initialized"), 2)
    if param["Proc"] == "transferEFW":
        Visum.Net.SetAttValue("EFW", True)
        # alle Verfahren deaktivieren
        for o in Visum.Procedures.Operations.GetAll:
            o.SetAttValue("ACTIVE", False)
            if o.AttValue("CODE") in ["TNM", "TNM_BEGINN", "TNM_EFW", "TNM_FILTER"]:
                o.SetAttValue("ACTIVE", True)
        Visum.Procedures.Execute()
        addIn.ReportMessage(_("EFW of SN: transfered"), 2)
    if param["Proc"] == "openLAR":
        on_excel_LAR(param["TN"])
        addIn.ReportMessage(_("Finished"), 2)
    if param["Proc"] == "exportMETN":
        # Erst Filtern
        filter_operation = Visum.Procedures.Operations.ItemByKey([o.AttValue("CODE") for o in Visum.Procedures.Operations.GetAll].index("TNM_FILTER") + 1) # +1, da Index in Verfahrensablauf
        Visum.Procedures.OperationExecutor.SetCurrentOperation(filter_operation)
        Visum.Procedures.OperationExecutor.ExecuteCurrentOperation()
        on_excel_METN()
        addIn.ReportMessage(_("Finished"), 2)
    
    
def on_excel_LAR(TN):
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
        
        
def on_excel_METN():
        pythoncom.CoInitialize()  # important inside GUI event
        xls = _get_excel()
        werte = _get_METN_werte()
        top_left = xls.Range("A1")
        bottom_right = top_left.Offset(len(werte), len(werte[0]))
        write_range = xls.Range(top_left, bottom_right)
        write_range.Value = tuple(tuple(w) for w in werte)
        xls.Range("A1").ClearContents()
        xls.Range("E1").ClearContents()
        xls.Range("F1").Value = _("LENGTH")
        

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


def _get_METN_werte():
    attribute = [[r"VEHJOURNEY\LINENAME", "VEHJOURNEYNO", "DAYVECTOR", "SAISON", "LENGTH", "VEHCOMBNO", "DEP", "DURATION", r"FROMSTOPPOINT\NAME", r"TOSTOPPOINT\NAME",
                  "ISVALID(MO)", "ISVALID(TU)", "ISVALID(WE)", "ISVALID(TH)", "ISVALID(FR)", "ISVALID(SA)", "ISVALID(SU)"],
                 ["Linie", "Fahrtnr", "Tage" ,"Saison", "LENGTH", "Fzg", "Ab", "Dauer", "Von", "Nach",
                  "MO", "DI", "MI", "DO", "FR", "SA", "SO"]]
    df_fahrplanfahrten = pd.DataFrame(Visum.Net.VehicleJourneySections.GetMultipleAttributes(attribute[0], True),
                            columns = attribute[1])
    df_fahrplanfahrten.insert(0, "first", 4)
    df_fahrplanfahrten['LENGTH'] = df_fahrplanfahrten['LENGTH'].round(2)
    df_fahrplanfahrten['Von - Nach'] = df_fahrplanfahrten['Von'] + ' - ' + df_fahrplanfahrten['Nach']
    df_fahrplanfahrten = df_fahrplanfahrten.drop(columns=['Von', 'Nach'])
    df_fahrplanfahrten[["Dauer", "Ab"]] = df_fahrplanfahrten[["Dauer", "Ab"]].astype("float64")
    
    # Fahrtbeginn nach Mitternacht
    mask = df_fahrplanfahrten['Ab'] >= 86400
    df_fahrplanfahrten.loc[mask, 'Ab'] -= 86400
    cols = ["MO", "DI", "MI", "DO", "FR", "SA", "SO"]
    df_fahrplanfahrten.loc[mask, cols] = df_fahrplanfahrten.loc[mask, cols].to_numpy()[:, [6,0,1,2,3,4,5]]
    
    df_fahrplanfahrten["Ab"] = df_fahrplanfahrten["Ab"].apply(__seconds_to_hhmm)
    df_fahrplanfahrten["Dauer"] = df_fahrplanfahrten["Dauer"].apply(__seconds_to_hhmm)
    
    # Wochentage
    wochentage = ["MO", "DI", "MI", "DO", "FR"]
    for i, e in [[wochentage+["SO"], "SA"], [wochentage+["SA"], "SO"]]:
        mask = df_fahrplanfahrten[i].any(axis=1) & df_fahrplanfahrten[e]
        df_copy = df_fahrplanfahrten[mask].copy()
        df_fahrplanfahrten.loc[mask, e] = 0
        df_copy[i] = 0
        df_copy[e] = 1
        df_fahrplanfahrten = pd.concat([df_fahrplanfahrten, df_copy], ignore_index=True)
    df_fahrplanfahrten["MO"] = df_fahrplanfahrten["MO"].eq(0).map({True: ".", False: "1"})
    df_fahrplanfahrten["DI"] = df_fahrplanfahrten["DI"].eq(0).map({True: ".", False: "2"})
    df_fahrplanfahrten["MI"] = df_fahrplanfahrten["MI"].eq(0).map({True: ".", False: "3"})
    df_fahrplanfahrten["DO"] = df_fahrplanfahrten["DO"].eq(0).map({True: ".", False: "4"})
    df_fahrplanfahrten["FR"] = df_fahrplanfahrten["FR"].eq(0).map({True: ".", False: "5"})
    df_fahrplanfahrten["SA"] = df_fahrplanfahrten["SA"].eq(0).map({True: ".", False: "6"})
    df_fahrplanfahrten["SO"] = df_fahrplanfahrten["SO"].eq(0).map({True: ".", False: "7"})
    cols = ["MO", "DI", "MI", "DO", "FR", "SA", "SO"]
    df_fahrplanfahrten["Tage"] = df_fahrplanfahrten[cols].agg("".join, axis=1)
    df_fahrplanfahrten.drop(columns=cols, inplace=True)
    
    df_fahrplanfahrten.sort_values(by=["Linie", "Ab", "Tage"], ascending=[True, True, False], inplace=True) # Sortierung wie in METN
    rows_fahrplanfahrten = df_fahrplanfahrten.values.tolist()
    rows_fahrplanfahrten.insert(0, list(df_fahrplanfahrten.columns))
    return rows_fahrplanfahrten


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
    
    
def __seconds_to_hhmm(seconds):
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    return f"{h:02d}:{m:02d}"


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
