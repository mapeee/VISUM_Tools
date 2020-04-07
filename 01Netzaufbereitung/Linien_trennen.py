# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        Skript zum selektieren und trennen doppelter Linien innerhalb einer Linie nach dem HAFAS-Import
# Purpose:
#
# Author:      mape
#
# Created:     01/03/2017
# Copyright:   (c) mape 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------

#---Vorbereitung---#
import win32com.client.dynamic
import time
import numpy as np
start_time = time.clock() ##Ziel: am Ende die Berechnungsdauer ausgeben

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(r'C:'+path),'VISUM_Tools','Linien_trennen.txt')
f = path.read_text()
f = f.split('\n')

"""#--Eingabe-Parameter--#"""
Netz = 'C:'+f[0]
VSys = ["BUS","AST"] ##Für die folgenden Verkehrssysteme nicht



#--VISUM öffnen--#
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.17")
VISUM.loadversion(Netz)
VISUM.Filters.InitAll()

"""Linien-Schleife"""
"""while: Es sollen zwoelf Durchläufe gerechnet werden, da nicht davon auszugehen ist, dass es mehr als sieben Linien in einer Linie vereint sind."""
Durchlauf = 1
while Durchlauf < 12:
    print "--- Beginne mit Durchlauf: "+str(Durchlauf)+" ---"
    for Line in VISUM.Net.Lines:
        if Line.AttValue("TSysCode") not in VSys: continue
        if Durchlauf == 1: pass ##Der erste Durchlauf findet immer statt.
        else:
            if "("+str(Durchlauf)+")" not in Line.AttValue("Name"): continue ##Es sollen nur die Linien angeschaut werden, die im Durchlauf zuvor neu angelegt wurden (daher -1).

        #--Linien-Filter--#
        ##Der Vergleich wird immer nur für alle Fpf. einer Linie durchgeführt
        Name = Line.AttValue("Name")
        print "Ueberpruefung fuer die folgende Linie: "+Name
        VISUM.Filters.InitAll()
        Linien = VISUM.Filters.LineGroupFilter()
        Linien.UseFilterForVehJourneys = True ##Wirksam für Fahrplanfahrt-Abschnitte
        Linien.UseFilterForLines = True ##Wirksam für Fahrplanfahrt-Abschnitte
        Linien.UseFilterForLineRoutes = True ##Wirksam für Fahrplanfahrt-Abschnitte
        Linien = Linien.LineFilter()
        Linien.AddCondition("OP_NONE",False,"Name","EqualVal",Name)##Kein weiterer Filter,nicht komplementär,Szenario,GleichWert,Wert=1

        """Alle Routen dieser Linie werden nur mit der laengsten Linienroute dieser Linie verglichen"""
        #Liste aller Linienrouten dieser Linie mit den zugehoerigen Knoten
        Anzahl_Knoten = np.array(VISUM.Net.LineRoutes.GetMultiAttValues("NumStopPoints",True))[:,1]
        Routen_ID = np.array(VISUM.Net.LineRoutes.GetMultiAttValues("ID",True))[:,1]
        Route_max = VISUM.Net.LineRoutes.ItemByID(np.array(VISUM.Net.LineRoutes.GetMultiAttValues("ID",True))[:,1][np.argmax(Anzahl_Knoten)])
        Routen_Knoten = []
        for i in Route_max.LineRouteItems: ##Muss so ausgewählt werden, da der String über Verketten nicht alle Nummern erfasst.
            Routen_Knoten.append(i.AttValue("StopPoint\StopArea\StopNo"))


        #Waehle die laengste aus
        z = 0
        while z < len(Routen_ID): ##nur solange noch neue Linien in die Wolke aufgenommen werden. Wenn nicht, ist z am Ende der Schleife gleich der Anzahl der Linienrouten
            z = 0
            for i in Routen_ID:
                Route = VISUM.Net.LineRoutes.ItemByID(i)
                Knoten_alle = []
                for e in Route.LineRouteItems: ##Muss so ausgewählt werden, da der String über Verketten nicht alle Nummern erfasst.
                    Knoten_alle.append(e.AttValue("StopPoint\StopArea\StopNo"))
                if Route.AttValue("AddVal1") == 1:
                    z+=1
                    continue
                for Knoten in Knoten_alle:
                    if Knoten in Routen_Knoten: ##Sobald ein KLnoten in der längsten Route gefunden wurde, wird die schleife verlassen
                        Route.SetAttValue("AddVal1",1)
                        Routen_Knoten+= Knoten_alle
                        break
                if Route.AttValue("AddVal1") == 1:continue
                z+=1

        #Neue Linie für alle Routen, die nicht zu der bestehenden Wolke gehören.
        for i in Routen_ID:
            Route = VISUM.Net.LineRoutes.ItemByID(i)
            if Route.AttValue("AddVal1") == 1: continue
            #Route hat keine Kontaktpunkte zur längsten Route
            try:
                if Durchlauf == 1: ##Nur im ersten Durchgang wird die Klammer amgehaengt, dann nur noch ersetzt.
                    Text = Name+"("+str(Durchlauf+1)+")"
                    VISUM.Net.AddLine(Text,Line.AttValue("TSysCode"))
                else:
                    Text = Name.replace("("+str(Durchlauf)+")","("+str(Durchlauf+1)+")")
                    VISUM.Net.AddLine(Text,Line.AttValue("TSysCode"))
                print "Neue Linie "+Text+" eingefuegt."
                angelegt = "ja"
                Route.Line = Text
            except: #'Falls es diese Linie schon gibt.
                Route.Line = Text

    Durchlauf+=1


"""Abschluss"""
Sekunden = int(time.clock() - start_time)

print "--Scriptdurchlauf erfolgreich nach",Sekunden,"Sekunden!--"