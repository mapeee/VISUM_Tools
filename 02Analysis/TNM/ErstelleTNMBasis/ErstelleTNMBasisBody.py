# -*- coding: utf-8 -*-

import sys
import wx
from VisumPy.helpers import SetMulti
from VisumPy.AddIn import AddIn, AddInState, AddInParameter
_ = AddIn.gettext


def Run(param):
    '''
    -Umlaute (Kommentare der BDA)
    -Koordinatensystem einstellen (wird benötigt?)

    Hinweise ergänzen, was ggf noch zu tun ist:
        - Fahrzeuge hinzufügen und Fahrten zuweisen
    '''
    _erstelle_bdg(Visum)
    _erstelle_bda(Visum)
    if param["Gebiete"]:
        _import_gebiete(Visum)
    if param["Kalender"]:
        _kalender(Visum)
    if param["Fahrzeuge"]:
        _fahrzeuge(Visum)
    _verfahrensablauf(Visum)
    _zeitintervallmengen(Visum)
    addIn.ReportMessage(_("TNM-Basis: created"), 2)
    

def _erstelle_bda(Visum):
    def __attribute_bda(_bda, _comment, _bdg, _n, _string = None):
        _bda.Comment = _comment
        _bda.UserDefinedGroup = _bdg
        if _string:
            _bda.MaxStringLen = _string
        return n+1

    # Netz-Parameter
    n = 0
    if not Visum.Net.AttrExists("ANABOVERLAP"):
        bda = Visum.Net.AddUserDefinedAttribute("ANABOVERLAP", "ANABOVERLAP", "ANABOVERLAP", 9, DefVal = False)
        bda = Visum.Net.Attributes.ItemByKey("ANABOVERLAP")
        n = __attribute_bda(bda, r"Identische Ankunfts- und Abfahrtszeit gilt als Überlappung", "TNM_Parameter", n)
    if not Visum.Net.AttrExists("EFW"):
        bda = Visum.Net.AddUserDefinedAttribute("EFW", "EFW", "EFW", 9, DefVal = True)
        bda = Visum.Net.Attributes.ItemByKey("EFW")
        n = __attribute_bda(bda, "Verwendung eines Entfernungswerkes bei Distanzberechnung", "TNM_Parameter", n)
    if not Visum.Net.AttrExists("F_WOCHEN"):
        bda = Visum.Net.AddUserDefinedAttribute("F_WOCHEN", "F_WOCHEN", "F_WOCHEN", 1, 0, False, 0, 52, 12)
        bda = Visum.Net.Attributes.ItemByKey("F_WOCHEN")
        n = __attribute_bda(bda, "Ferienwochen je Jahr", "TNM_Parameter", n)
    if not Visum.Net.AttrExists("S_WOCHEN"):
        bda = Visum.Net.AddUserDefinedAttribute("S_WOCHEN", "S_WOCHEN", "S_WOCHEN", 1, 0, False, 0, 52, 40)
        bda = Visum.Net.Attributes.ItemByKey("S_WOCHEN")
        n = __attribute_bda(bda, "Normalwochen je Jahr", "TNM_Parameter", n)
    if not Visum.Net.AttrExists("M_FAHRZEUGBEDARF"):
        bda = Visum.Net.AddUserDefinedAttribute("M_FAHRZEUGBEDARF", "M_FAHRZEUGBEDARF", "M_FAHRZEUGBEDARF", 5, DefVal = "M1")
        bda = Visum.Net.Attributes.ItemByKey("M_FAHRZEUGBEDARF")
        n = __attribute_bda(bda, "Fahrzeugbedarf je TN nach Methode 1 oder Methode 2", "TNM_Parameter", n, 2)
    if not Visum.Net.AttrExists("TN"):
        bda = Visum.Net.AddUserDefinedAttribute("TN", "TN", "TN", 5, DefVal = "ohne")
        bda = Visum.Net.Attributes.ItemByKey("TN")
        n = __attribute_bda(bda, "Kürzel des aktiven Teilnetzes", "TNM", n, 10)
        
    # Fahrzeugkombinationen
    if not Visum.Net.VehicleCombinations.AttrExists("RANG"):
        bda = Visum.Net.VehicleCombinations.AddUserDefinedAttribute("RANG", "RANG", "RANG", 1, 0, False, 0, 100, 0)
        bda = Visum.Net.VehicleCombinations.Attributes.ItemByKey("RANG")
        n = __attribute_bda(bda, r"Reihenfolge nach der Fahrzuege im TNM ersetzt werden dürfen.", "TNM", n)
        
    # Linien
    if not Visum.Net.Lines.AttrExists("TN"):
        bda = Visum.Net.Lines.AddUserDefinedAttribute("TN", "TN", "TN", 5, DefVal = "ohne")
        bda = Visum.Net.Lines.Attributes.ItemByKey("TN")
        n = __attribute_bda(bda, r"Kürzel des aktiven Teilnetzes", "TNM", n, 10)
        
    # Linienrouten-Verläufe
    if not Visum.Net.LineRouteItems.AttrExists("EFW_METER"):
        bda = Visum.Net.LineRouteItems.AddUserDefinedAttribute("EFW_METER", "EFW_METER", "EFW_METER", 1, 0, False, 0, None, 0)
        bda = Visum.Net.LineRouteItems.Attributes.ItemByKey("EFW_METER")
        n = __attribute_bda(bda, "Meter laut Entfernungswerk", "EFW", n)
    
    # Fahrplanfahrtabschnitte
    if not Visum.Net.VehicleJourneySections.AttrExists("SAISON"):
        bda = Visum.Net.VehicleJourneySections.AddUserDefinedAttribute("SAISON", "SAISON", "SAISON", 5, DefVal = "S+F")
        bda = Visum.Net.VehicleJourneySections.Attributes.ItemByKey("SAISON")
        n = __attribute_bda(bda, r"Normalwochen (S), Ferien (F) oder ganzjährig (S+F)", "TNM", n, 3)
    if not Visum.Net.VehicleJourneySections.AttrExists("S"):
        bda = Visum.Net.VehicleJourneySections.AddUserDefinedAttribute("S", "S", "S", 9, Formula = 'IF([SAISON]="S"|[SAISON]="S+F";1;0)')
        bda = Visum.Net.VehicleJourneySections.Attributes.ItemByKey("S")
        n = __attribute_bda(bda, "Ist Normalwoche", "TNM", n)
    if not Visum.Net.VehicleJourneySections.AttrExists("F"):
        bda = Visum.Net.VehicleJourneySections.AddUserDefinedAttribute("F", "F", "F", 9, Formula = 'IF([SAISON]="F"|[SAISON]="S+F";1;0)')
        bda = Visum.Net.VehicleJourneySections.Attributes.ItemByKey("F")
        bda.Comment = "Ist Ferienwoche"
        bda.UserDefinedGroup = "TNM"
        n = __attribute_bda(bda, "Ist Ferienwoche", "TNM", n)
    if n > 0:
        Visum.Log(20480, _("%s UDA: created") %(str(n)))
        

def _erstelle_bdg(Visum):
    '''
    Erstellt alle notwendigen BDG.
    '''
    bdg = [["TNM", "Teilnetzmanagement"],
           ["TNM_Parameter", "TNM Parameter"],
           ["EFW", "Entfernungswerk"]]
    BDGs = Visum.Net.UserDefinedGroups.GetMultiAttValues("NAME")
    n = 0
    for i, iname in bdg:
        if not any(i in BDG for _, BDG in BDGs):
            Visum.Net.AddUserDefinedGroup(i, iname)
            n+=1
    if n > 0:
        Visum.Log(20480, _("%s UDG: created") %(str(n)))
    
def _fahrzeuge(Visum):
    n = Visum.Net.VehicleCombinations.Count
    Visum.IO.LoadSQLiteDatabase(addIn.DirectoryPath+"data\Fahrzeuge.sqlite3", True)
    diff_n = Visum.Net.VehicleCombinations.Count - n
    if diff_n > 0:
        Visum.Log(20480, _("%s VehicleCombinations: imported") %(str(diff_n)))

def _import_gebiete(Visum):
    n = Visum.Net.Territories.Count
    JSONImport = Visum.IO.CreateImportGeoJSONPara()
    JSONImport.AddAttributeAllocation("NAME", "NAME")
    JSONImport.AddAttributeAllocation("CODE", "CODE")
    JSONImport.AddAttributeAllocation("TYPENO", "TYPENO")
    JSONImport.ObjectType = 8 # Gebiete als Fläche
    Visum.IO.ImportGeoJSON(addIn.DirectoryPath+"data\TNM_Gebiete.geojson", JSONImport)
    diff_n = Visum.Net.Territories.Count - n
    if diff_n > 0:
        Visum.Log(20480, _("%s Territories: imported") %(str(diff_n)))

def _kalender(Visum):
    if Visum.Net.CalendarPeriod.AttValue("TYPE") != "CALENDARPERIODWEEK":
        Visum.Net.CalendarPeriod.SetAttValue("TYPE", "CALENDARPERIODWEEK")
        Visum.Log(20480, _("Calendar: changed to weekly calendar"))
    if Visum.Net.CalendarPeriod.AttValue("ANALYSISPERIODENDDAYINDEX") != 7:
        Visum.Net.CalendarPeriod.SetAttValue("ANALYSISPERIODENDDAYINDEX",7)
        Visum.Log(20480, _("Calendar: Analysis period changed to Mon-Sun"))
        
def _verfahrensablauf(Visum):
    '''
    Importiere Verfahrensablauf (keine Verfahrensparameter und Verfahrensvariablen)
    Setze den neuen Verfahrensablauf an den Anfang aber nur, wenn es die Gruppe 'TNM Leistungsabrechnung' noch nicht gibt.
    '''
    if not any(o.AttValue("COMMENT") == "TNM Leistungsabrechnung" for o in Visum.Procedures.Operations.GetAll):
        Visum.Procedures.OpenXmlWithOptions(addIn.DirectoryPath+"data\TNM_Verfahrensablauf.xml", True, False, 2, 0, False)
        Visum.Log(20480, _("Procedure sequence: imported"))
        
def _zeitintervallmengen(Visum):
    '''
    - Stelle Analyseperiode von Montag bis Sonntag
    - ermittlere erste freie Nummer (ID) für neue Zeitintervallmenge
    - Lege neue Zeitintervallmenge mit erster freier Nummer an, falls 'TNM_Tage' noch nicht vorhanden.
    - Füge Zeitintervalle für Mo-So hinzu
    - Setze die Zeitintervallmenge aktiv
    '''
    if Visum.Net.CalendarPeriod.AttValue("ANALYSISPERIODENDDAYINDEX") != 7:
        Visum.Net.CalendarPeriod.SetAttValue("ANALYSISPERIODENDDAYINDEX",7)
        Visum.Log(20480, _("Calendar: Analysis period changed to Mon-Sun"))
    _zi = Visum.Net.TimeIntervalSets.GetMultiAttValues("NO")
    _ziNO = {int(v) for _, v in _zi} 
    _ersteNO = 1
    while _ersteNO in _ziNO:
        _ersteNO += 1
    if not any(ts.AttValue("CODE") == "TNM_Tage" for ts in Visum.Net.TimeIntervalSets.GetAll):
        ts = Visum.Net.AddTimeIntervalSet(_ersteNO)
        ts.SetAttValue("CODE", "TNM_Tage")
        ts.SetAttValue("NAME", "TNM Wochentage")
        tage = [("Mo", "Montag"),
                ("Di", "Dienstag"),
                ("Mi", "Mittwoch"),
                ("Do", "Donnerstag"),
                ("Fr", "Freitag"),
                ("Sa", "Samstag"),
                ("So", "Sonntag"),]
        for i, (tag1, tag2) in enumerate(tage):
            ti = ts.AddTimeInterval(tag1, 0, 86400, DayIndex=i+1)
            ti.SetAttValue("NAME", tag2)
        Visum.Net.CalendarPeriod.SetAttValue("ANALYSISTIMEINTERVALSETNO", _ersteNO)
        Visum.Log(20480, _("Time intervals: created and activated"))


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
    except:
        addIn.HandleException(addIn.TemplateText.MainApplicationError)
