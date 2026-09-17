# -*- coding: utf-8 -*-
import sys
import wx
import re
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / "libs"))
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
from VisumPy.helpers import SetMulti
_ = AddIn.gettext

import pandas as pd
import warnings ## regarding pandas messages
warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
warnings.filterwarnings(
    "ignore",
    message="Dtype inference on a pandas object.*",
    category=FutureWarning
)

'''
- Hash zu LR für Bündelung
- Dokumentation
    - Kleine Word-Hilfe und Verknüpfung mit Tool
'''

def Run(param):
    '''
    0. Checke den Input auf Fehler / Fehleingaben
    1. Trage notwendige DataFrames zusammen
    2. Erstelle oder wähle existierende Linie
    3. Filter auf diese Linie
    4. Loop über alle Fahrplanfahrten
    4.1 Erstellen der Linienroute
    4.2 Erstellen eines Fahrzeitprofils
    4.3 Erstellung Fahrplanfahrten mit Attrbiten
    5. Aggregation der neuen LinienRouten und Fahrten

    Parameters
    ----------
    param : Dictionary
        Parameters from Visum dialogue

    Returns
    -------
    None
    '''
    dfExcel = pd.read_excel(param["ExcelFolder"], header=None)
    Validate_Excel(dfExcel)
    dfLineAttr, dfStopsTimes, dfVJAttr = GetTables(dfExcel) # creates all DataFrames
    Validate_Line(dfLineAttr)
    Validate_VJAttr(dfVJAttr, dfLineAttr) # validDays and vehicleCombinations
    Validate_Stops(dfStopsTimes, dfLineAttr)
    Validate_Times(dfStopsTimes)
    Visum.Filters.LineGroupFilter().Init() # Only init Line Filter
    SetMulti(Visum.Net.LineRoutes,"AddVal3",[0]*Visum.Net.LineRoutes.Count,False) # to identify new ones by 99 later
    Line, Direction = AddLine(dfLineAttr) # Creating or selecting Line
    FilterLine(Line) # Filter line for better handling
    
    for VJIndex in range(2, dfStopsTimes.shape[1]):
        dfVJStopsTimes = dfStopsTimes.iloc[:, [0, 1, VJIndex]] # first two cols are StopName und StopNo
        LRNetElements, VJTimes, DepVJ = GetStopsTimes(dfVJStopsTimes)
        LR = AddLineRoute(Line, Direction, LRNetElements)
        VJAttr = dfVJAttr[["VJAttr", VJIndex]]
        VJAttr = VJAttr.set_index("VJAttr").iloc[:, -1]
        TP = AddTimeProfile(LR, VJTimes)
        AddVehicleJourney(TP, DepVJ, VJAttr)
        
    LineRouteFilter = Visum.Filters.LineGroupFilter().LineRouteFilter()
    LineRouteFilter.AddCondition("OP_NONE", False, "ADDVAL3", "EqualVal", 99)
    
    AggrPT() # Aggregation of all new PT-Elements
    
    if not Visum.Workbench.IsTabularTimetableRunning(): # open tabular timetable with new VJ if not active yet
        tt = Visum.Workbench.CreateTabularTimetable
        tt.Show(True)
        tt.ShowOnlyActiveVehJourneys = True
        NELines = Visum.CreateNetElements()
        NELines.Add(Line)
        ttss = tt.LineSelectionAndStopSequence
        ttss.ClearLineSelection()
        ttss.SetLineSelection(NELines)
        ttss.CalculateStopSequenceFromLineSelection()
        Visum.Workbench.ActivateNetworkEditor()
    Visum.Graphic.Autozoom(Line)


def AddLine(dfLineAttr):
    '''
    1. Attribute der Linie aus dfLineAttr zusammentragen
    2. Lege neue Linie an
    2.1 Lege Oberlinie an, falls noch nicht vorhanden

    Parameters
    ----------
    dfLineAttr : pandas series
        Linienattribute (Name, VSys, Richtung etc.)

    Returns
    -------
    Line : Visum Object
        Linienobjet der Linie
    Direction : TYPE
        Code der Richtung

    Raises
    ------
    ValueError
        Rückgabe Text bei fehlenden oder nicht plausiblen Eingaben
    '''
    LineName = str(dfLineAttr["Linie"])
    TSys = dfLineAttr["Verkehrssystem"]
    Direction = dfLineAttr["Richtung"]
    ML = dfLineAttr["Oberlinie"]
    ML = None if pd.isna(ML) else str(ML)
    TN = dfLineAttr["Teilnetz"]
    TN = None if pd.isna(TN) else str(TN)
    if any(l == LineName for _, l in Visum.Net.Lines.GetMultiAttValues("NAME", False)):
        Line = Visum.Net.Lines.ItemByKey(LineName)
    else: # create new line if not existing
        Line = Visum.Net.AddLine(LineName, TSys)
        if ML:
            Line.SetAttValue("MAINLINENAME", ML)
        if TN and Visum.Net.Lines.AttrExists("TN"):
            Line.SetAttValue("TN", TN)
    return Line, Direction

def AddLineRoute(_Line, _Direction, LRNetElements):
    '''
    1. Erstelle die Routingparameter
    2. Erzeuge den Namen der neuen LinienRoute (darf noch nicht vergeben sein)
    3. Erstelle die neue LinienRoute

    Parameters
    ----------
    _Line : Visum-Objekt
        Linie als Visum-Objekt
    _Direction : String
        Richtungscode der Fahrplanfahrt
    LRNetElements : Visum-Container
        Abfolge der HaltePunkte in Visum

    Returns
    -------
    LR : Visum-Objekt
        Neue Linienroute als Visum-Objekt
    '''
    RouteSearchTSys = _getRouteSearchTSys()
    nameLR = _getNameLR(_Line)
    LR = Visum.Net.AddLineRoute(nameLR, _Line, _Direction, LRNetElements, RouteSearchTSys) #create the line route
    LR.SetAttValue("ADDVAL3", 99)
    return LR

def AddTimeProfile(_LR, _VJTimes):
    '''
    Erzeuge FahrzeitProfil und setze Ankunfts- und Abfahrtszeiten (bei Halt)

    Parameters
    ----------
    _LR : Visum-Objekt
        LinienRoute als Visum-Objekt
    _VJTimes : Liste
        Ankunfgszeiten und Haltezeiten als int in Tages-Sekunden

    Returns
    -------
    _TP : Visum-Objekt
        FahrzeitProfil als Visum-Objekt
    '''
    _TP = Visum.Net.AddTimeProfile("1", _LR)
    _ArrTimes, _StopTimes = _VJTimes
    for i, TPI in enumerate(_TP.TimeProfileItems.GetAll):
        if i == 0: # first Item
            continue
        if i == len(_TP.TimeProfileItems.GetAll) - 1: # last Item
            TPI.SetAttValue("ARR", _ArrTimes[i])
            continue
        TPI.SetAttValue("ARR", _ArrTimes[i])
        if _StopTimes[i]: # set only if not 0
            TPI.SetAttValue("DEP", _StopTimes[i])
    return _TP

def AddVehicleJourney(_TP, _DepVJ, _VJAttr):
    '''
    1. Suche erste freie Nummer für eine Fahrplanfahrt (aus allen Fahrten)
    2. Erzeuge Fahrplanfahrt
    3. Setze Abfahrtszeit
    4. Setze Verkehrstage
    5. Setze Fahrzeug
    6. Setze Fahrplanfaht Nummer (nicht ID)
    7. Setze Saison (wenn vorhanden)

    Parameters
    ----------
    _TP : Visum-Objekt
        Fahrzeitprofil als Visum-Objekt
    _DepVJ : Int
        Abfahrtszeit in Tages-Sekunden
    _VJAttr : Pandas DataFrame
        Unterschiedliche Attribute zur FahrplanFahrt

    Returns
    -------
    None
    '''
    VJno = [_no for _, _no in Visum.Net.VehicleJourneys.GetMultiAttValues("NO", False)]
    freeVJno = next(i for i in range(1, int(max(VJno, default=0)) + 2) if i not in VJno)
    VJ = Visum.Net.AddVehicleJourney(freeVJno, _TP)
    VJ.SetAttValue("DEP", _DepVJ)
    if pd.notna(_VJAttr["Bezeichnung"]):
        VJ.SetAttValue("NAME", _VJAttr["Bezeichnung"])
    
    # VehicleJourneySections Attributes
    VJS = VJ.VehicleJourneySections.GetAll[0]
    if pd.notna(_VJAttr["Fahrzeug"]):
        VehCode = _VJAttr["Fahrzeug"]
        Veh = Visum.Net.VehicleCombinations.GetMultipleAttributes(["NO", "CODE"])
        VehNo = next(no for no, name in Veh if name == VehCode)
        VJS.SetAttValue("VEHCOMBNO", int(VehNo))
    if pd.notna(_VJAttr["Saison"]) and Visum.Net.VehicleJourneySections.AttrExists("SAISON"):
        VJS.SetAttValue("SAISON", _VJAttr["Saison"])
    if Visum.Net.CalendarPeriod.AttValue("TYPE") != 'CALENDARPERIODWEEK':
        return
    if pd.notna(_VJAttr["Verkehrstage"]):
        DayCode = _VJAttr["Verkehrstage"]
        Day = Visum.Net.ValidDaysCont.GetMultipleAttributes(["NO","CODE"])
        DayNo = next(no for no, name in Day if name == DayCode)
        if not int(DayNo) == 1: # change only if not standard (daily)
            VJS.SetAttValue("VALIDDAYSNO", int(DayNo))
        
def AggrPT():
    '''
    Aggregiert die neuen Linienrouten und Fahrzeitprofile in Visum zur Reduktion der Menge.
    '''
    LRAggrPara = Visum.Net.LineRoutes.GetLineRouteAggregationParameters
    LRAggrPara.SetAttValue("AggregateVehJourneys", False)
    LRAggrPara.SetAttValue("AggregationMode", 1)
    LRAggrPara.SetAttValue("LineRouteSameStartStopPoint", True)
    LRAggrPara.SetAttValue("LineRouteSameEndStopPoint", True)
    LRAggrPara.SetAttValue("LineRouteSameLinksOverlapShare", 1.0)
    Visum.Net.LineRoutes.Aggregate(LRAggrPara, True)
    
    
def FilterLine(Line):
    '''
    Filter auf Linie der neuen Fahrplanfahrten

    Parameters
    ----------
    Line : Visum-Objekt
        Linie als Visum-Objekt

    Returns
    -------
    None
    '''
    LineName = Line.AttValue("NAME")
    Lines = Visum.Filters.LineGroupFilter()
    Lines.Init()
    LineFilter = Lines.LineFilter()
    LineFilter.AddCondition("OP_NONE", False, "NAME", "EqualVal", LineName)   
    Lines.UseFilterForLines = True
    Lines.UseFilterForLineRoutes = True
    Lines.UseFilterForLineRouteItems = True
    Lines.UseFilterForTimeProfiles = True
    Lines.UseFilterForTimeProfileItems = True
    Lines.UseFilterForVehJourneys = True
    Lines.UseFilterForVehJourneySections = True
    Lines.UseFilterForVehJourneyItems = True

def GetStopsTimes(_dfVJStopsTimes):
    '''
    Loop über alle Zielen der Reihenfolge der Haltepunkte
    1. Wenn Zeit leer oder |, dann kein Halte
    2. Wenn erste Fahrzeit, dann Starthalt und Abfhartszeit
    3. Wenn gleiche HaltepunktNo wie zuvor, dann Haltezeit = Zeit aktueller Loop - Zeit vorheriger loop
    4. Ankunftszeit = maximum aus Ankunft- und Abfahrtszeit zuvor minus Differenz aus Zeit aktueller Loop - Zeit vorheriger loop

    Parameters
    ----------
    _dfVJStopsTimes : pandas DataFrame
        Haltestellenfolge und Zeiten der Fahrplanfahrt

    Returns
    -------
    LRE : Visum-Objekt
        Visum-Container mit HaltePunkten als Visum-Objekte
    VJTimes list
        Listen mit Ankunftszeiten und Abfahrtszeiten (Haltezeiten) je HaltePunkt
    DepVJ : Int
        Abfahrtszeit in Tages-Sekunden
    '''
    StopNo = 0
    RunTime = -1 # -1 to identify first DepTime and previous loop
    ArrTimes = [] # list with arrival times
    StopTimes = [] # list with StopTimes (if same StopPointNo one after another)
    LRE = Visum.CreateNetElements()
    
    for _, row in _dfVJStopsTimes.iterrows():
        if pd.isna(row.iloc[2]): # no time for StopNo
            continue
        # RunTimeNew = row.iloc[2].hour * 3600 + row.iloc[2].minute * 60 + row.iloc[2].second
        RunTimeNew = row.iloc[2].total_seconds()
        if RunTime == -1: # first loop
            DepVJ = RunTimeNew
            ArrTimes.append(0)
            StopTimes.append(0)
            RunTime = RunTimeNew
            StopNo = row[1]
            LRE.Add(Visum.Net.StopPoints.ItemByKey(row[1]))
            continue
        if StopNo == row[1]: # if same stop one after another (to handle waiting times)
            StopTimes[-1] = ArrTimes[-1] + (RunTimeNew - RunTime) # StopTime for Stop before is ArrTime before + (new ArrTime and ArrTime before)
            RunTime = RunTimeNew
            continue
        ArrTimes.append(max(ArrTimes[-1], StopTimes[-1]) + (RunTimeNew - RunTime))
        StopTimes.append(0)
        RunTime = RunTimeNew
        StopNo = row[1]
        LRE.Add(Visum.Net.StopPoints.ItemByKey(StopNo))
    VJTimes = [ArrTimes, StopTimes]
    return LRE, VJTimes, DepVJ    

def GetTables(dfExcel):
    '''
    Aufbereitung aller Auswertungs DataFrames

    Parameters
    ----------
    dfExcel : pandas df
        Import Excel ohne Bearbeitung

    Returns
    -------
    dfLineAttr : pandas series
        Linienattribute (Name, VSys, Richtung etc.)
    dfStopsTimes : pandas df
        Abfolge der angefahrenen Haltepunkte (Name+Nummer) und Zeiten der Fahrplanfahrten
    dfVJAttr : pandas df
        Attribute der Fahrplanfahrten (Tage, Fahrzeug, Zeiten etc.)
    '''
    dfExcel = dfExcel.dropna(how="all") # drops all rows where all values are nan
    rowsplit = dfExcel[1].eq("Verkehrstage").idxmax()-1 # index-1 of row in col1 where value is 'Verkehrstage'
    dfLineAttr = dfExcel.loc[:rowsplit-1, :1]
    dfLineAttr = dfLineAttr.set_index(0)[1]
    dfLineRoute = dfExcel.loc[rowsplit:]
    dfLineRoute = dfLineRoute.dropna(axis=1, how="all") # drops all cols where all values are nan
    rowSP = dfExcel[0].eq("Haltepunkt").idxmax()
    dfStopsTimes = dfLineRoute.loc[rowSP+1:] # deletes first rows until 'Haltepunkt'
    dfStopsTimes = dfStopsTimes.mask(dfStopsTimes.isin(["|", ""]))
    dfVJAttr = dfLineRoute.loc[:rowSP-1, 1:]
    dfVJAttr.rename(columns={dfVJAttr.columns[0]: "VJAttr"}, inplace=True)
    return dfLineAttr, dfStopsTimes, dfVJAttr

def Validate_Excel(dfExcel):
    '''
    Prüft, ob die Einträge 'Verkehrstage' und 'Haltepunkt' zur Strukturierung in Excel vorhanden.

    Parameters
    ----------
    dfExcel : pandas df
        Excle-Import ohne Anpassungen

    Returns
    -------
    None

    Raises
    ------
    ValueError
        Fehler, wenn Einträge 'Verkehrstage' oder 'Haltepunkt' in Excel fehlen.
    '''
    for i in ["Verkehrstage", "Fahrzeug", "Saison", "Bezeichnung"]:
        if not dfExcel[1].eq(i).any():
            raise ValueError(_("Entry '%s' is missing in excel") % i)
    if not dfExcel[0].eq("Haltepunkt").any():
        raise ValueError(_("Entry 'Haltepunkt' is missing in excel"))
        
def Validate_Line(dfLineAttr):
    '''
    1. Check, ob VSys im Netz vorhanden
    2. Check, ob RichtungsCode im Netz vorhanden
    3. Falls Linie schon vorhanden: Check, ob VSys zur Linie passt

    Parameters
    ----------
    dfLineAttr : pandas series
        Linienattribute (Name, VSys, Richtung etc.)

    Raises
    ------
    ValueError
        Rückgabe Text bei fehlenden oder nicht plausiblen Eingaben
    '''
    LineName = dfLineAttr["Linie"]
    TSys = dfLineAttr["Verkehrssystem"]
    Direction = dfLineAttr["Richtung"]
    TN = dfLineAttr["Teilnetz"]

    if not any(t == TSys for _, t in Visum.Net.TSystems.GetMultiAttValues("CODE")):
        raise ValueError(_("Transport system %s did not exist") % TSys)
    if not any(d == Direction for _, d in Visum.Net.Directions.GetMultiAttValues("CODE")):
        raise ValueError(_("Direction %s did not exist") % Direction)
    if any(l == LineName for _, l in Visum.Net.Lines.GetMultiAttValues("NAME", False)):
        Line = Visum.Net.Lines.ItemByKey(LineName)
        if Line.AttValue("TSYSCODE") != TSys:
            raise ValueError(_("Line already exists with a different transit system"))
    if pd.notna(TN) and re.search(r'[ÖöÄäÜüß#\- ]', TN):
        raise ValueError(_(r"Line-UDA 'TN' contains special characters (e.g. []#\- )"))
            
def Validate_Stops(dfStopsTimes, dfLineAttr):
    '''
    1. Prüft, ob alle HaltePunktNummern im Netz vorhanden und für VSys zugelassen.
    2. Prüft, ob HaltePunktNummern am Beginn oder Ende von Fahrplanfahrt identisch

    Parameters
    ----------
    dfStopsTimes : pandas df
        Abfolge der angefahrenen Haltepunkte (Name+Nummer) und Zeiten der Fahrplanfahrten
    dfLineAttr : pandas series
        Linienattribute (Name, VSys, Richtung etc.)

    Returns
    -------
    None

    Raises
    ------
    ValueError
        Fehler, wenn HaltePunktNummern nicht vorhanden oder für VSys zugelassen.
        Fehler, Wenn HaltePunktNummern am Beginn oder Ende doppelt
    '''
    # closed or missing stops
    TSys = dfLineAttr["Verkehrssystem"]
    Stops = Visum.Filters.StopGroupFilter()
    Stops.Init()
    StopPointFilter = Stops.StopPointFilter()
    StopPointFilter.AddCondition("OP_NONE", False, "TSYSSET", 14, TSys)
    Stops.UseFilterForStopPoints = True
    StopNOs = [n for _, n in Visum.Net.StopPoints.GetMultiAttValues("NO", True)]
    if not dfStopsTimes[1].isin(StopNOs).all(): # col 1 has StopPointNo
        missingStopNOs = dfStopsTimes.loc[~dfStopsTimes[1].isin(StopNOs), 1].tolist()
        raise ValueError(_("Some StopPoints are closed for TSys or missing: %s") % missingStopNOs)
    Stops.Init()
    # identical stops at beginning or end of vj
    for vj in dfStopsTimes.columns[2:]:
        stopsvj = dfStopsTimes[[dfStopsTimes.columns[1], vj]]
        stopsvj = stopsvj.dropna(subset=[vj]).iloc[:, 0]
        if stopsvj.iloc[0] == stopsvj.iloc[1] or stopsvj.iloc[-1] == stopsvj.iloc[-2]:
            raise ValueError(_("Identical stops at beginning or end in Journey %s") % (vj-1))

    
def Validate_Times(dfStopsTimes):
    '''
    Prüfe, ob es bei den Fahrzeiten absteigende Werte gibt (negative Fahr- oder Haltezeiten)

    Parameters
    ----------
    dfStopsTimes : pandas df
        Abfolge der angefahrenen Haltepunkte (Name+Nummer) und Zeiten der Fahrplanfahrten

    Raises
    ------
    ValueError
        Fehler bei absteigenden Fahrzeiten, identische sind okay
    '''
    dfVJTimes = dfStopsTimes.iloc[:, 2:] # all cols but not first two with StopName and StopNo
    for col in dfVJTimes.columns:
        s = dfVJTimes[col].dropna()
        if not s.is_monotonic_increasing:
            raise ValueError(_("Decreasing times in Journey %s") % (col-1))

def Validate_VJAttr(dfVJAttr, dfLineAttr):
    '''
    Prüfe, ob alle Verkehrstage und FahrzeugKombinationen im Netz vorhanden sind.

    Parameters
    ----------
    dfVJAttr : pandas df
        Attribute der Fahrplanfahrten (Tage, Fahrzeug, Zeiten etc.)
    dfLineAttr : pandas series
        Linienattribute (Name, VSys, Richtung etc.)

    Raises
    ------
    ValueError
        Fehler, wenn Verkehrstage oder FahrzeugKombinationen nicht vorhanden.
    '''
    TSys = dfLineAttr["Verkehrssystem"]
    
    # Valid days
    vd_Excel = dfVJAttr.loc[dfVJAttr["VJAttr"] == "Verkehrstage"].iloc[0, 1:].dropna().tolist()
    vd_Visum = [n for _,n in Visum.Net.ValidDaysCont.GetMultiAttValues("CODE")]
    missing_days = [day for day in vd_Excel if day not in vd_Visum]
    if missing_days:
        raise ValueError(_("Valid Days are missing in Network: %s") % (", ".join(missing_days)))
        
    # VehicleCombinations
    vc_Excel = dfVJAttr.loc[dfVJAttr["VJAttr"] == "Fahrzeug"].iloc[0, 1:].dropna().tolist()
    vc_Visum = [code for code,t in Visum.Net.VehicleCombinations.GetMultipleAttributes(["CODE", "TSYSSET"]) if TSys in t]
    missing_vcs = [vc for vc in vc_Excel if vc not in vc_Visum]
    if missing_vcs:
        raise ValueError(_("VehicleCombinations are missing in Network: %s") % (", ".join(missing_vcs)))
    

def _getNameLR(_Line):
    '''
    Sucht nach dem ersten freien Namen der LinienRoute nach folgendem Muster:
        {LinienName}_{erste freie Nummer}

    Parameters
    ----------
    _Line : Visum-Objekt
        Linie als Visum-Objekt

    Returns
    -------
    str
        Name der neuen Linienroute
    '''
    _LineName = _Line.AttValue("NAME")
    LRnames = [_name for _, _name in Visum.Net.LineRoutes.GetMultiAttValues("Name", True)]
    prefix = f"{_LineName}_"
    used = {int(x[len(prefix):]) for x in LRnames
        if x.startswith(prefix) and x[len(prefix):].isdigit()}
    free = next(i for i in range(1, max(used, default=0) + 2) if i not in used)
    return f"{_LineName}_{free}"

def _getRouteSearchTSys():
    '''
    SUMMARY.

    Returns
    -------
    Visum-Objekt
        Routingparameter als Visum-Objekt
    '''
    RS = Visum.IO.CreateNetReadRouteSearchTSys()
    RS.SetAttValue("ChangeLinkTypeOfOpenedLinks", False)
    RS.SetAttValue("DeleteOldLineRoutes", False)
    RS.SetAttValue("IncludeBlockedTurns", False)
    RS.SetAttValue("HowToHandleIncompleteRoute", 2) # search shortest path
    RS.SetAttValue("ShortestPathCriterion", 1) # 1 = TSys travel time; 2 = travel time from linktype; 3 = link length
    RS.SetAttValue("MaxDeviationFactor", 50)
    RS.SetAttValue("WhatToDoIfShortestPathNotFound", 2) # insert link if necessary
    RS.SetAttValue("LinkTypeForInsertedLinksReplacingMissingShortestPaths", 1)
    RS.SetAttValue("WhatToDoIfStopPointIsBlocked", 2) # Open the StopPoint
    RS.SetAttValue("WhatToDoIfStopPointNotFound", 0) # Dont Read Line Route
    return RS


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
        defaultParam = {"" : False}
        param = addInParam.Check(True, defaultParam)
        Run(param)
        NumVJ = Visum.Net.VehicleJourneys.CountActive
        addIn.ReportMessage(_("Import of %s VehicleJournes finished") % NumVJ, 2)
        NumLinks = int(Visum.Net.Links.GetFilteredSet("[TYPENO]=1").Count/2)
        if NumLinks:
            addIn.ReportMessage(_("%s link(s) of TYPNO = 2 created due to unsuccessful routing") % NumLinks)
            
    except ValueError as e:
        addIn.ReportMessage(str(e))
    except Exception:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
