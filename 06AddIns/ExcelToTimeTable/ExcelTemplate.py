# -*- coding: utf-8 -*-
import pandas as pd
import win32com.client
import win32gui
from VisumPy.AddIn import AddIn
_ = AddIn.gettext


def open(Visum):
    '''
    1. Checke, ob überhaupt HaltePunkte aktiv sind.
    2. Stelle Verbindung zu Excel-Com her.
    3. Ergänze Werte in Excel.
    4. Öffne Excel im Vordergrund.

    Parameters
    ----------
    Visum : TYPE
        Laufende Visum Instanz

    Returns
    -------
    None
    '''
    addIn = AddIn(Visum)
    if not Visum.Net.StopPoints.CountActive:
        addIn.ReportMessage(_("No active Stops to add to ExcelTemplate"))
    pathExcel = addIn.DirectoryPath + _("ExcelTemplate.xlsx")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    wb = excel.Workbooks.Open(pathExcel, ReadOnly=True)
    AddValues(Visum, wb)
    
    win32gui.SetForegroundWindow(excel.Hwnd)
    
def AddValues(Visum, wb):
    '''
    Ergänze die Excel-Vorlage mit notwendigen Inhalten
    - HaltePunkte
    - Verkehrstage
    - Oberlinien
    - Verkehrssysteme
    - FahrzeugKombinationen

    Parameters
    ----------
    Visum : TYPE
        Laufende Visum Instanz.
    wb : TYPE
        Geöffnetes / aktives Excel-File.

    Returns
    -------
    None
    '''
    
    # StopPoints
    ws = wb.Worksheets("Halte")
    StopPoints = pd.DataFrame(Visum.Net.StopPoints.GetMultipleAttributes(["NAME", "NO", "TSYSSET"], True))
    data = tuple(map(tuple, StopPoints.to_numpy()))
    ws.Range(ws.Cells(1, 1),ws.Cells(len(StopPoints), 3)).Value = data
    
    # Days
    ws = wb.Worksheets("Verkehrstage")
    vd = [n for _,n in Visum.Net.ValidDaysCont.GetMultiAttValues("CODE")]
    ws.Range(ws.Cells(1, 1),ws.Cells(len(vd), 1)).Value = tuple((v,) for v in vd)
    
    # Mainlines
    ws = wb.Worksheets("Oberlinien")
    ml = [n for _,n in Visum.Net.MainLines.GetMultiAttValues("NAME")]
    ws.Range(ws.Cells(1, 1),ws.Cells(len(ml), 1)).Value = tuple((v,) for v in ml)
    
    # TSys
    ws = wb.Worksheets("Verkehrssysteme")
    ts = [c for c,t in Visum.Net.TSystems.GetMultipleAttributes(["CODE", "TYPE"]) if t == "PUT"]
    ws.Range(ws.Cells(1, 1),ws.Cells(len(ts), 1)).Value = tuple((v,) for v in ts)
    
    # VehicleCombinations
    ws = wb.Worksheets("Fahrzeuge")
    VehicleCombinations = pd.DataFrame(Visum.Net.VehicleCombinations.GetMultipleAttributes(["CODE", "TSYSSET"]))
    data = tuple(map(tuple, VehicleCombinations.to_numpy()))
    ws.Range(ws.Cells(1, 1),ws.Cells(len(VehicleCombinations), 2)).Value = data