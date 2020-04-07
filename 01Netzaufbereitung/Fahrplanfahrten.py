# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        Bearbeitung von Fahrplanfahrten, die sich ueberlappen
# Purpose:
#
# Author:      mape
#
# Created:     17/02/2017
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
path = Path.joinpath(Path(r'C:'+path),'VISUM_Tools','Fahrplanfahrten.txt')
f = path.read_text()
f = f.split('\n')

"""#--Eingabe-Parameter--#"""
Netz = 'C:'+f[0]
VSys = [""] ##Nur für die folgenden Verkehrssysteme



#--VISUM öffnen--#
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.17")
VISUM.loadversion(Netz)
VISUM.Filters.InitAll()


"""#--Funktionen--#"""
'''Diese Funktion erstellt einen Array mit den notwendigen Angaben einer Fpf'''
def Fpf_daten (Fpf): ##Übergabe: VISUM.Fahrplanfahrt
    Nr = Fpf.AttValue("No")
    t_Start = Fpf.AttValue("Dep")
    t_Ziel = Fpf.AttValue("Arr")
    von = Fpf.AttValue("FromTProfItemIdentifier").split(" ")[1]
    nach = Fpf.AttValue("ToTProfItemIdentifier").split(" ")[1]
    Stop = Fpf.VehicleJourneyItems.GetAll ##Der vierte Stop dieser Fahrplanfahrt.
    Stopps = len(Stop) ##Anzahl Stopps einer Fahrplanfahrten
    try: ##Falls die Fahrt sehr kurz ist, kann die Länge nicht verglichen werden
        Stop_a4 = Stop[3].AttValue("TimeProfileItem\LineRouteItem\NodeNo")
        Stop_end4 = Stop[Stopps-4].AttValue("TimeProfileItem\LineRouteItem\NodeNo") ##Nimm den viertletzten Stop
    except:
        Stop_a4 = 999
        Stop_end4 = 999

    return Nr, t_Start, t_Ziel, von, nach, Stopps, Stop_a4, Stop_end4


'''Dieser Funktion vergleicht zwei Fahrplanfahrten'''
def Vergleich(Fpf1,Fpf2):
    Fpf1_daten = Fpf_daten(Fpf1)
    Fpf2_daten = Fpf_daten(Fpf2)

    #vom Start aus
    ###Gleiche Abfahrtszeit, Gleiche Abfahrtsknoten, ungleiche Nummer
    if Fpf1_daten[1] == Fpf2_daten[1] and Fpf1_daten[3] == Fpf2_daten[3] and Fpf1_daten[0] != Fpf2_daten[0]:

        ##Und gleiche Ankunftszeit und gleicher Zielknoten
        if Fpf1_daten[2] == Fpf2_daten[2] and Fpf1_daten[4] == Fpf2_daten[4]: ##Gleiche Linienführung
            I_gleich = 1
        ##Ankunftszeit und oder Zielnoten nicht gleich, aber ersten vier Knoten gleich
        elif Fpf1_daten[6] == Fpf2_daten[6] and Fpf1_daten[6] != 999: ##Gleicher Startpunkt, ersten vier Haltestellen gleich.; ungleich 999 damit kein Fehler bei ganz kurzen Fahrten
            if Fpf1_daten[5] >= Fpf2_daten[5]:
                I_gleich = 11 ##Fpf1 ist länger oder gleich lang!

            else:I_gleich = 12 ##Fpf2 ist länger oder gleich lang!
        ##Gleiche Abfahrtszeit und gleicher Knoten, aber nicht gleiches Ziel und auch nicht auf mind. vier Teilwegen gleich.
        else:I_gleich = 9

    ##Bei manchen Fahrplanfahrten ist die Ankunftszeit gleich, nicht aber die Abfahrtszeit, da diese sich am Ziel überlappen aber an unterschiedl. Punkten Starten.
    #vom Ziel aus
    elif Fpf1_daten[2] == Fpf2_daten[2] and Fpf1_daten[4] == Fpf2_daten[4] and Fpf1_daten[0] != Fpf2_daten[0]: ##Gleiches Ziel

        if Fpf1_daten[7] == Fpf2_daten[7] and Fpf1_daten[6] != 999: ##viert letzter  Halt ist gleich
            if Fpf1_daten[5] >= Fpf2_daten[5]:
                I_gleich = 11 ##Fpf1 ist länger oder gleich lang!

            else:I_gleich = 12

        else:I_gleich = 9 ##also nur auf kurzem Teilweg gleich!!
    else:I_gleich = 9


    return I_gleich


'''Diese Funktion erstellt neue Fahrplantage oder loescht Fpf'''
"""
Fahrplanfahrten die gelöscht werden sollen, bekommen den ZWert2 = 1. Anschließend löschen per Hand, da das im Skript zu lange dauert.
"""
def VTage (Vjour1,Vjour2,Ident,Nr_VT,Vector_VT):  ##*args: beliebig viele weitere Parameter
    loesch_fahrten = 0
    loesch_tage = 0
    for section in Vjour1.VehicleJourneySections:VT1 = section.AttValue("ValidDays\DayVector")
    for section in Vjour2.VehicleJourneySections:VT2 = section.AttValue("ValidDays\DayVector")
    try:
        VT1 = list(VT1)
        VT2 = list(VT2)
    except:
        return loesch_tage, loesch_fahrten
    if Ident == 1:
        for i in range(7):
            VT1[i] = (VT1[i] if VT1[i] == "1" else VT2[i]) ##Nimm den gültigen Tag von VT1 sonst VT2
        ##VISUM.Net.RemoveVehicleJourney(Vjour2) ##VT2 kann dann weg, alle Verkehrstage sind in VT1 enthalten
        Vjour2.SetAttValue("AddVal2",1)
        loesch_fahrten+=1
        VT1 = "".join(VT1)
        try:
            VT_neu = Nr_VT[np.where(VT1 == Vector_VT)[0]][0]
        except:
            return loesch_tage, loesch_fahrten
        for section in Vjour1.VehicleJourneySections: section.SetAttValue("ValidDaysNo",VT_neu)
    if Ident in [11,12]:
        for i in range(7):
            if (VT1[i] == VT2[i]) and (VT1[i] != "0"):
                VT1[i] = (VT1[i] if Ident == 11 else "0")
                VT2[i] = ("0" if Ident == 11 else VT2[i])
                loesch_tage+=1 ##Wenn Bedingung erfüllt, fällt immer ein Tag weg.
        VT1 = "".join(VT1)
        VT2 = "".join(VT2)
        if VT1 != "0000000" and VT1 not in Vector_VT:
            print "Verkehrstage gibt es noch nicht!!!!!: "+VT1+", "+"Fahrplanfahrt: "+str(Vjour1.AttValue("No"))
            return loesch_tage, loesch_fahrten
        if VT2 != "0000000" and VT2 not in Vector_VT:
            print "Verkehrstage gibt es noch nicht!!!!!: "+VT2+", "+"Fahrplanfahrt: "+str(Vjour2.AttValue("No"))
            return loesch_tage, loesch_fahrten
        if VT1 == "0000000":
            ##VISUM.Net.RemoveVehicleJourney(Vjour1)
            Vjour1.SetAttValue("AddVal2",1)
            loesch_fahrten+=1
        else:
            VT_neu = Nr_VT[np.where(VT1 == Vector_VT)[0]][0]
            for section in Vjour1.VehicleJourneySections: section.SetAttValue("ValidDaysNo",VT_neu)
        if VT2 == "0000000":
            ##VISUM.Net.RemoveVehicleJourney(Vjour2)
            Vjour2.SetAttValue("AddVal2",1)
            loesch_fahrten+=1
        else:
            VT_neu = Nr_VT[np.where(VT2 == Vector_VT)[0]][0]
            for section in Vjour2.VehicleJourneySections: section.SetAttValue("ValidDaysNo",VT_neu)

    if loesch_fahrten == 1: ##Entweder ganze Fahrt wird gelöscht oder einzelne Tage
        loesch_tag = 0
    return loesch_tage, loesch_fahrten


"""Linien-Schleife"""
for Line in VISUM.Net.Lines:
    loesch_tage = 0
    loesch_fahrten = 0

    ##if Line.AttValue("TSysCode") in VSys: continue
    ##if Line.AttValue("AddVal1") == 0: continue

    #--Linien-Filter--#
    ##Der Vergleich wird immer nur für alle Fpf. einer Linie durchgeführt
    Name = Line.AttValue("Name")
    print "Bereinigung fuer die folgende Linie: "+Name
    VISUM.Filters.InitAll()
    Linien = VISUM.Filters.LineGroupFilter()
    Linien.UseFilterForVehJourneys = True ##Wirksam für Fahrplanfahrt-Abschnitte
    Linien.UseFilterForLines = True ##Wirksam für Fahrplanfahrt-Abschnitte
    Linien = Linien.LineFilter()
    Linien.AddCondition("OP_NONE",False,"Name","EqualVal",Name)##Kein weiterer Filter,nicht komplementär,Szenario,GleichWert,Wert=1

    """Fahrplanfahrten-Schleife"""
    """Es wird jede Fpf einer Linie mit allen Fpf dieser Linie verglichen."""
    #Liste aller aktiven Fahrplanfahrten dieser Linie
    FPF_Linie = np.array(VISUM.Net.VehicleJourneys.GetMultiAttValues("No",True))[:,1]
    t_Start_alle = np.array(VISUM.Net.VehicleJourneys.GetMultiAttValues("Dep",True))[:,1]
    t_Ziel_alle = np.array(VISUM.Net.VehicleJourneys.GetMultiAttValues("Arr",True))[:,1]
    Nr_VT = np.array(VISUM.Net.ValidDaysCont.GetMultiAttValues("No"))[:,1]
    Vector_VT = np.array(VISUM.Net.ValidDaysCont.GetMultiAttValues("DayVector"))[:,1]
    for FPF in FPF_Linie:
        try: ##Wenn die Fahrt bereits gelöscht wurde, dann kann man diese nicht mehr abrufen.
            Fpf_Bezug = VISUM.Net.VehicleJourneys.ItemByKey(FPF)
            Fpf_Bezug_Daten = Fpf_daten(Fpf_Bezug)
        except:continue
        if Fpf_Bezug.AttValue("AddVal2") == 1: continue ##Wenn diese Fpf nicht mehr gültig ist.

        #Suche in allen Fpf dieser Linie die Fpf mit der gleichen Ankunfts.- bzw. Abfahrtszeit
        nt_start = np.where(t_Start_alle == Fpf_Bezug_Daten[1])[0]
        nt_ziel = np.where(t_Ziel_alle == Fpf_Bezug_Daten[2])[0]
        nt_startziel = np.concatenate([nt_start,nt_ziel])
        if len(nt_startziel) == 2: continue ##wenn zwei, dann nur die Fahrten gefunden, die gerade betrachtet wird.

        for i in nt_startziel:
            Nr_i = FPF_Linie[i] ##Gibt die Nummer der zu vergleichenden Fahrplanfahrt aus.
            if Nr_i == Fpf_Bezug.AttValue("No"): continue ##Führe Berechnung nur dann durch, wenn nicht die gleiche Fpf miteinander vergleichen wird!
            try:Fpf_alle = VISUM.Net.VehicleJourneys.ItemByKey(Nr_i)
            except: continue ##Fals Fpf schon gelöscht wurde
            if Fpf_alle.AttValue("AddVal2") == 1: continue ##Wenn diese Fpf nicht mehr gültig ist.
            Fpf_aenderung = Vergleich(Fpf_Bezug,Fpf_alle) ##hier werden die beiden Fpf miteinander verglichen

            if Fpf_aenderung==9: continue##wenn keine Gemeinsamkeit, dann mach einfach weiter, sonst ändern

            else:
                loeschen = VTage(Fpf_Bezug,Fpf_alle,Fpf_aenderung,Nr_VT,Vector_VT)
                loesch_fahrten+= loeschen[1]
                loesch_tage+= loeschen[0]

    print "geloeschte Fahrten: "+str(loesch_fahrten)
    print "geloeschte Tage: "+str(loesch_tage)


"""Abschluss"""
Sekunden = int(time.clock() - start_time)

print "--Scriptdurchlauf erfolgreich nach",Sekunden,"Sekunden!--"