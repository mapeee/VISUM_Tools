#!/usr/bin/env python3
"""
Filter alle relevanten Netzelemente auf das aktive Teilnetz.
Linien und Gebiete.

Erstellt: 25.12.2025
@author: mape
"""

from TNM_Checks import check_BDA


def NEFilter(Visum):
    if not check_BDA(Visum, Visum.Net, "Network", "TN"):
        return False
    if not check_BDA(Visum, Visum.Net.Lines, "Lines", "TN"):
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
    else:
        Territories = Visum.Filters.TerritoryFilter()
        Territories.AddCondition("OP_NONE", False, "TYPENO", "ContainedIn", 1)
        Visum.Log(20480, f"Filter gesetzt auf TN: {TN}")
    return True

NEFilter(Visum)
