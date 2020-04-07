# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        VISUM Strecken Teiler
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
XLS = 'C:'+f[1]

#--VISUM öffnen--#
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.15")
VISUM.loadversion(Netz)
VISUM.Filters.InitAll()

#--Hst-Nummer ändern--#
##Spalten aus xls auswählen
Dokument = xlrd.open_workbook(XLS)
Blatt = Dokument.sheet_by_index(0) ##greift auf das erste (0) Tabellenblatt zu
Zeilen = Blatt.nrows-1 ##-1 wegen Überschriften
print "Es gibt ",Zeilen,"Zeilen!"

for i in range(Zeilen):

#--Knoten--#
    VonKnoten = int(Blatt.cell_value(rowx=1+i,colx=1)) ##1+i damit überschriften ignoriert werden
    ZuKnoten = int(Blatt.cell_value(rowx=1+i,colx=2))
    KnotenNr = int(Blatt.cell_value(rowx=1+i,colx=0))

    ##Haltestelle auswählen
    print "Neuer Knoten angebunden: "+str(int(KnotenNr))

    try:
##        if VISUM.Net.Nodes.ItemByKey(KnotenNr).AttValue("NumLinks") >0:
##            if VISUM.Net.Nodes.ItemByKey(KnotenNr).AttValue("Count:StopPoints") >0:
##                print "gab es schon!!!"
##                continue
##            else:
##                VISUM.Net.Nodes.ItemByKey(KnotenNr).SetAttValue("No",KnotenNr+54698321)
##                print "neue Nummer vergeben!!!"
##                continue
        VISUM.Net.Links.SplitViaNode(VonKnoten,ZuKnoten,KnotenNr)
    except:
        print "Geht nicht!!"
        continue




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

##Ende
Sekunden = int(time.clock() - start_time)

print "--Scriptdurchlauf erfolgreich nach",Sekunden,"Sekunden!--"