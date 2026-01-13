#!/usr/bin/env python3
"""
Filtert alle relevanten Netzelemente für das aktive Teilnetz.
Ebenen:
    - Gebiete
    - Linien

Erstellt: 25.12.2025
@author: mape
Version: 0.9
"""

from TNM_Checks import check_bda


def tnm_filter(Visum):
    '''Filtert die relevanten Netzelemente und führt vorab checks auf fehlende BDAs durch.
    Filtert Linien, die dem im Netz-BDA 'TN' hinterlegten Teilnetzkürzel zugeordnet sind.
    Filtert Gebiete (Territorien), die als Kreise für die Teilnetzabrechnung gekennzeichnet sind (TYPENO) == 1.
    Teste, ob der filter aktive Linien und Gebiete zurückgibt.
    '''
    if not check_bda(Visum, Visum.Net, "Network", "TN"):
        # Check, ob Netz-BDA 'TN' existiert.
        return False
    if not check_bda(Visum, Visum.Net.Lines, "Lines", "TN"):
        # Check, ob Linien-BDA 'TN' existiert.
        return False
    TN = Visum.Net.AttValue("TN")
    Visum.Filters.InitAll()
    Lines = Visum.Filters.LineGroupFilter()
    LineFilter = Lines.LineFilter()
    LineFilter.AddCondition("OP_NONE", False, "TN", "EqualVal", TN)
    Lines.UseFilterForLines = True
    Lines.UseFilterForLineRoutes = True
    Lines.UseFilterForLineRoutes = True
    Lines.UseFilterForLineRouteItems = True
    Lines.UseFilterForTimeProfiles = True
    Lines.UseFilterForTimeProfileItems = True
    Lines.UseFilterForVehJourneys = True
    Lines.UseFilterForVehJourneySections = True
    Lines.UseFilterForVehJourneyItems = True
    if Visum.Net.Lines.CountActive == 0:
        Visum.Log(12288, f"{TN} gefiltert: keine Linie vorhanden")
        return False
    Territories = Visum.Filters.TerritoryFilter()
    Territories.AddCondition("OP_NONE", False, "TYPENO", "ContainedIn", 1)
    if Visum.Net.Territories.CountActive == 0:
        Visum.Log(12288, f"{TN} gefiltert: keine Kreise (Gebiete) vorhanden")
        return False
    Visum.Log(20480, f"Filter gesetzt auf TN: {TN}")
    return True

if __name__ == "__main__":           
    tnm_filter(Visum)
