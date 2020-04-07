# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        VISUM neue Hst-Nummer und neue Haltestellen
# Purpose:
#
# Author:      mape
#
# Created:     18/09/2014
# Copyright:   (c) mape 2014
# Licence:     <your licence>
#-------------------------------------------------------------------------------

#---Vorbereitung---#
import win32com.client.dynamic
import xlrd ##nur zum lesen von Excel-Daten
import time
start_time = time.clock() ##Ziel: am Ende die Berechnungsdauer ausgeben

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(r'C:'+path),'VISUM_Tools','Haltestellen.txt')
f = path.read_text()
f = f.split('\n')

#--Eingabe-Parameter--#
Netz = 'C:'+f[0]
XLS = 'C:'+[1]

#--VISUM öffnen--#
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.17")
VISUM.loadversion(Netz)
VISUM.Filters.InitAll()

#--Hst-Nummer ändern--#
##Spalten aus xls auswählen
Dokument = xlrd.open_workbook(XLS)
Blatt = Dokument.sheet_by_index(0) ##greift auf das erste (0) Tabellenblatt zu
Zeilen = Blatt.nrows-1 ##-1 wegen Überschriften
print("Es gibt ",Zeilen,"Zeilen!")

##for i in range(Zeilen):
##
##    #--Nummern von Haltestellen ändern--#
##    alt = int(Blatt.cell_value(rowx=1+i,colx=0)) ##1+i damit überschriften ignoriert werden
##    neu = int(Blatt.cell_value(rowx=1+i,colx=1)) ##1+i damit überschriften ignoriert werden
##    ##Anzahl = int(Blatt.cell_value(rowx=1+i,colx=3))
##    print str(alt)+" ersetzt durch "+str(neu)
##    if Anzahl > 1:
##        print "ist doppelt!!"
##        continue
##    try:
##        VISUM.Net.RemoveNode(VISUM.Net.Nodes.ItemByKey(neu))
##        VISUM.Net.Nodes.ItemByKey(alt).SetAttValue("No",neu)
##
##    VISUM.Net.Stops.ItemByKey(alt).SetAttValue("No",neu)
##        VISUM.Net.StopAreas.ItemByKey(alt).SetAttValue("No",neu)
##        VISUM.Net.StopPoints.ItemByKey(alt).SetAttValue("No",neu)
##    except:
##        print "geht nicht"
##        VISUM.Net.Nodes.ItemByKey(neu).SetAttValue("No",neu+54887839)
##        VISUM.Net.Nodes.ItemByKey(alt).SetAttValue("No",neu)



#--Haltestellen einfuegen--#
##    X = float(Blatt.cell_value(rowx=1+i,colx=0)) ##1+i damit überschriften ignoriert werden
##    Y = float(Blatt.cell_value(rowx=1+i,colx=1))
##    Nr = int(Blatt.cell_value(rowx=1+i,colx=2))
##    Name = Blatt.cell_value(rowx=1+i,colx=3)
##
##    #Knoten
##    print "Nummer: "+str(Nr)
##    print "Name: "+unicode(Name)
##
##    try:
##        VISUM.Net.AddNode(Nr,X,Y)
##    except:
##        print "Knoten konnte nicht eingefügt werden "+str(Nr)
##        X = VISUM.Net.Nodes.ItemByKey(Nr).AttValue("XCoord") ##Damit Alles weitere auf bereits bestehende Knoten gesetzt wird.
##        Y = VISUM.Net.Nodes.ItemByKey(Nr).AttValue("YCoord")
##
##    #Haltestelle
##
##    try:
##        VISUM.Net.AddStop(Nr,X,Y)
##        VISUM.Net.Stops.ItemByKey(Nr).SetAttValue("Name",Name)
##    except:
##        print "Hst konnte nicht eingefügt werden "+str(Nr)
##
##    #Haltestellenbereich
##
##    try:
##        VISUM.Net.AddStopArea(Nr,Nr,Nr,X,Y)
##        VISUM.Net.StopAreas.ItemByKey(Nr).SetAttValue("Name",Name)
##    except:
##        print "Haltestellenbereich konnte nicht eingefügt werden "+str(Nr)
##
##    #Haltepunkt
##
##    try:
##        VISUM.Net.AddStopPointOnNode(Nr,Nr,Nr)
##        VISUM.Net.StopPoints.ItemByKey(Nr).SetAttValue("Name",Name)
##    except:
##        print "Haltestellenbereich konnte nicht eingefügt werden "+str(Nr)


#--Strecken--#
##    vonKnoten = int(Blatt.cell_value(rowx=1+i,colx=0)) ##1+i damit überschriften ignoriert werden
##    nachKnoten = int(Blatt.cell_value(rowx=1+i,colx=1)) ##1+i damit überschriften ignoriert werden
##    neuStreckNr = int(Blatt.cell_value(rowx=1+i,colx=3))
##
##    try:
##        print "Aendern"
##        Link = VISUM.Net.Links.ItemByKey(vonKnoten,nachKnoten)
##        print neuStreckNr
##
##        try:
##            Link.SetNo(neuStreckNr)
##        except:
##            print "error "+neuStreckNr
##    except:
##        continue


#--Haltestellenbereiche neu zuordnen
for i in range(Zeilen):

    #--Nummern von Haltestellen ändern--#
    nummer = int(Blatt.cell_value(rowx=1+i,colx=0)) ##1+i damit überschriften ignoriert werden
    alt = int(Blatt.cell_value(rowx=1+i,colx=1)) ##1+i damit überschriften ignoriert werden
    neu = int(Blatt.cell_value(rowx=1+i,colx=2)) ##1+i damit überschriften ignoriert werden
    ##Anzahl = int(Blatt.cell_value(rowx=1+i,colx=3))
    print(str(alt)+" ersetzt durch "+str(neu))
##    print str(i)

    hstbereich = VISUM.Net.StopAreas.ItemByKey(nummer)
    try: hstbereich.SetAttValue("StopNo",neu)
    except: "Geht nicht"


#Ende
Sekunden = int(time.clock() - start_time)

print("--Scriptdurchlauf erfolgreich nach",Sekunden,"Sekunden!--")