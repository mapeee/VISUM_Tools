# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        VISUM neue Hst-Nummer
# Purpose:
#
# Author:      mape
#
# Created:     18/09/2014
# Copyright:   (c) mape 2014
# Licence:     <your licence>
#-------------------------------------------------------------------------------

#---Vorbereitung---#
import sys
for i in Sys_Pfad:
    sys.path.append(i)
import win32com.client.dynamic
import time
start_time = time.clock() ##Ziel: am Ende die Berechnungsdauer ausgeben

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(r'C:'+path),'VISUM_Tools','Knotennummern.txt')
f = path.read_text()
f = f.split('\n')

#--Eingabe-Parameter--#
Netz = 'V:'+f[0]

#--VISUM �ffnen--#
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.14")
VISUM.loadversion(Netz)

#--Linien--#
##Welche gibt es?
for Linien in VISUM.Net.Lines: ##Schleife �ber alle Linien
    print Linien.AttValue("Name") ##Ausgabe der Liniennamen

    for Routen in Linien.LineRoutes: ##Schleife �ber alle Linienrouten dieser Linie
        print Routen.AttValue("Name")

        ##for Knoten in Routen.LineRouteItems:
            ##print Knoten.AttValue("NodeNo")

#--Add-Methods--#
VISUM.Net.AddLine("test","S") ##F�gt Linie "test" im Vsys S-Bahn hinzu.

HP = VISUM.CreateNetElements() ##Erstellt Container f�r die notwendigen Nummern der haltepunkte der Linienroute
HP.Add(VISUM.Net.StopPoints.ItemByKey(19994)) ##Hbf
HP.Add(VISUM.Net.StopPoints.ItemByKey(39992)) ##Altona

Read = VISUM.CreateNetReadRouteSearchTSys() ##Erstellt Object f�r Regeln bei Rputensuche
Read.SearchShortestPath(3,True,True,8,0,3)
VISUM.Net.AddLineRoute("testroute","test","<",HP,Read) ##erstellt die Linienroute und ber�cksichtigt Regel

VISUM.Net.AddTimeProfile("f",VISUM.Net.LineRoutes.ItemByKey("test","<","testroute"))

TP = VISUM.Net.TimeProfiles.ItemByKey("test","<","testroute","f") ##Objekt Zeitprofil zum ver�ndern
for Arr in TP.TimeProfileItems:
    if int(Arr.AttValue("Arr")) == 90:
        Arr.SetAttValue("Arr",int(120))

VISUM.Net.AddVehicleJourney(3899,TP) ##F�ge Fahrplanfahrt ein
FF = VISUM.Net.VehicleJourneys.ItemByKey(3899) ##W�hle diese Fahrplanfahrt aus
FF.SetAttValue("Dep",10*60*60) ##Abfahrt um 10 Uhr



##Ende
Sekunden = int(time.clock() - start_time)

print "--Scriptdurchlauf erfolgreich nach",Sekunden,"Sekunden!--"