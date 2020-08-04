# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        Different changes in VISUM networks
# Purpose:
#
# Author:      mape
#
# Created:     03/08/2015
# Copyright:   (c) mape 2014
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import win32com.client.dynamic
import xlrd
import time
start_time = time.time()

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(path),'VISUM_Tools','Nummern.txt')
f = path.read_text()
f = f.split('\n')

#--Parameter--#
method = "links" #nodes, stoppoints

Netz = f[0]
XLS = f[1]

print("excel "+XLS)

#--open VISUM--#
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.20")
VISUM.loadversion(Netz)

#Spalten aus xls auswählen
Dokument = xlrd.open_workbook(XLS)
Blatt = Dokument.sheet_by_index(0) ##greift auf das erste (0) Tabellenblatt zu
Zeilen = Blatt.nrows-1 ##-1 wegen Überschriften
print("Es gibt ",Zeilen,"Zeilen!")


#--Knoten--#
if method == "nodes":
    for i in range(Zeilen):
        Alt = int(Blatt.cell_value(rowx=1+i,colx=0)) ##1+i damit überschriften ignoriert werden
        Neu = Blatt.cell_value(rowx=1+i,colx=2)
        
        if Alt != Neu:
            VISUM.Net.Nodes.ItemByKey(int(Alt)).SetAttValue("No",int(Neu))
            
#--Strecken--#
if method == "links":
    for i in range(Zeilen):
        vonKnoten = int(Blatt.cell_value(rowx=1+i,colx=1)) ##1+i damit überschriften ignoriert werden
        nachKnoten = int(Blatt.cell_value(rowx=1+i,colx=2)) ##1+i damit überschriften ignoriert werden
        altStreckNr = int(Blatt.cell_value(rowx=1+i,colx=0))
        neuStreckNr = int(Blatt.cell_value(rowx=1+i,colx=3))
        if altStreckNr == neuStreckNr: continue

        try:
            Link = VISUM.Net.Links.ItemByKey(vonKnoten,nachKnoten)
            Link.SetNo(neuStreckNr)
        except:
            print ("error "+neuStreckNr)
            continue

    ##Haltestelle auswählen
    # print "alt: "+str(int(Alt))
    # print "neu: "+str(int(Neu))
    ##VISUM.Net.Nodes.ItemByKey(int(Alt)).SetAttValue("No",int(Neu)) ##222501 ist nur Zufallszahl

##    if Alt != Neu:
##        try:
##            VISUM.Net.Stops.ItemByKey(int(Alt)).SetAttValue("No",int(Neu)) ##222501 ist nur Zufallszahl
##        except:
##            print "gibt es schon, neuer Versuch"
##            VISUM.Net.Nodes.ItemByKey(int(Neu)).SetAttValue("No",int(Neu)+5814789) ##5814789 ist nur Zufallszahl
##            VISUM.Net.Nodes.ItemByKey(int(Alt)).SetAttValue("No",int(Neu))
##            continue

    # try:
    #     VISUM.Net.Stops.ItemByKey(int(Alt)).SetAttValue("No",int(Neu))
    #     VISUM.Net.StopAreas.ItemByKey(int(Alt)).SetAttValue("No",int(Neu)) ##222501 ist nur Zufallszahl
    #     VISUM.Net.StopPoints.ItemByKey(int(Alt)).SetAttValue("No",int(Neu)) ##222501 ist nur Zufallszahl
    #     VISUM.Net.Nodes.ItemByKey(int(Alt)).SetAttValue("No",int(Neu)) ##222501 ist nur Zufallszahl
    # except:
    #     print "keine Haltestelle vorhanden!"

##for i in range(Zeilen):
##
###--Knoten--#
##    ##vorhanden = Blatt.cell_value(rowx=1+i,colx=5) ##1+i damit überschriften ignoriert werden
##    Alt = int(Blatt.cell_value(rowx=1+i,colx=0)) ##1+i damit überschriften ignoriert werden
##    Neu = Blatt.cell_value(rowx=1+i,colx=1)
##    ##Zahl = int(Blatt.cell_value(rowx=1+i,colx=3))
##
##    ##Haltestelle auswählen
##    ##print "alt: "+str(int(Alt))
##    ##print "neu: "+str(int(Neu))
##    try:
##        try:
##            VISUM.Net.Nodes.ItemByKey(int(Alt)).SetAttValue("No",int(Neu)) ##222501 ist nur Zufallszahl
##        except:
##            VISUM.Net.Nodes.ItemByKey(int(Neu)).SetAttValue("No",int(Neu)+954128) ##222501 ist nur Zufallszahl
##            VISUM.Net.Nodes.ItemByKey(int(Alt)).SetAttValue("No",int(Neu)) ##222501 ist nur Zufallszahl
####        if Zahl < 2: ##wenn die Hst mehr als einen HstBer hat,dann keine Änderung der Hst-Nr.
####            VISUM.Net.Stops.ItemByKey(int(Alt)).SetAttValue("No",int(Neu)) ##222501 ist nur Zufallszahl
##        VISUM.Net.StopAreas.ItemByKey(int(Alt)).SetAttValue("No",int(Neu)) ##222501 ist nur Zufallszahl
##        VISUM.Net.StopPoints.ItemByKey(int(Alt)).SetAttValue("No",int(Neu)) ##222501 ist nur Zufallszahl
##    except:
##        print "Fehler bei "+str(int(Alt))+"!!!!!!!!!!!!!!"

##    try:
##        VISUM.Net.Stops.ItemByKey(int(Alt)).SetAttValue("No",int(Neu)) ##222501 ist nur Zufallszahl
##    except:
##        print "Fehler bei "+str(int(Alt))+" (Hst)"


##for Knoten in VISUM.Net.Nodes:
##    try:
##        Knoten.SetAttValue("No",Knoten.AttValue("AddVal1"))
##    except:
##        pass
##
##
##for HPunkte in VISUM.Net.StopPoints:
##    HPunkte.SetAttValue("No",HPunkte.AttValue("NodeNo"))
##
##for HstBer in VISUM.Net.StopAreas:
##    if HstBer.AttValue("No") != HstBer.AttValue("Min:StopPoints\No"):
##        HstBer.SetAttValue("No",HstBer.AttValue("Min:StopPoints\No"))
##    if HstBer.AttValue("AddVal1") !=0:
##        HstBer.SetAttValue("No",HstBer.AttValue("AddVal1"))

##end
Sekunden = int(time.time() - start_time)

print("--finished after ",Sekunden,"seconds--")