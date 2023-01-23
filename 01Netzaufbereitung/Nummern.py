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
from datetime import datetime

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(path),'VISUM_Tools','Nummern.txt')
f = path.read_text()
f = f.split('\n')

#--Parameter--#
nodes = 1
merge_connect = 0
links = 0


Netz = f[0]
XLS = f[1]

print("excel "+XLS)

#--outputs--#
date = datetime.now()
txt_name = 'Nummern_'+date.strftime('%m%d%Y')+'.txt'
f = open(Path.home() / 'Desktop' / txt_name,'w+')
f.write(time.ctime()+"\n")
f.write("Input Network: "+Netz+"\n")
f.write("Change Excel: "+XLS+"\n")

f.write("\n\n\n\n")

#--open VISUM--#
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.22")
VISUM.loadversion(Netz)

#Spalten aus xls auswählen
Dokument = xlrd.open_workbook(XLS)
Blatt = Dokument.sheet_by_index(0) ##getting first sheet (0)
Zeilen = Blatt.nrows-1 ##-1 due to headlines
print(" ",Zeilen,"rows!")


#--Nodes--#
if nodes == 1:
    f.write("Node changes \n")
    for i in range(Zeilen):
        old = int(Blatt.cell_value(rowx=1+i,colx=2)) #+1 to ignore headlines
        new = Blatt.cell_value(rowx=1+i,colx=0)
        
        if old != new:
            try:
                VISUM.Net.Nodes.ItemByKey(int(old)).SetAttValue("No",int(new))
                f.write("node: old: "+str(old)+" new: "+str(int(new))+"\n")
            except: 
                f.write("!!Error: node: old: "+str(old)+"\n")
                print("!!Error: "+str(old)+"; neu: "+str(new))
    f.write("\n\n\n")

#--Connections--#
if merge_connect == 1:
    f.write("Merge nodes of connections \n")
    for i in range(Zeilen):
        keepingnode = int(Blatt.cell_value(rowx=1+i,colx=0))
        dyingnode = Blatt.cell_value(rowx=1+i,colx=1)
        try:VISUM.Net.Nodes.Merge(int(keepingnode),int(dyingnode))
        except:pass
        f.write("keeping node: "+str(keepingnode)+" dying node: "+str(dyingnode)+"\n")
    f.write("\n\n\n")

#--Links--#
if links == 1:
    f.write("Link changes \n")
    for i in range(Zeilen):
        fromNode = int(Blatt.cell_value(rowx=1+i,colx=1))
        toNode = int(Blatt.cell_value(rowx=1+i,colx=2))
        oldLinkNo = int(Blatt.cell_value(rowx=1+i,colx=0))
        newLinkNo = int(Blatt.cell_value(rowx=1+i,colx=3))
        if oldLinkNo == newLinkNo: continue

        try:
            Link = VISUM.Net.Links.ItemByKey(fromNode,toNode)
            Link.SetNo(newLinkNo)
            f.write("link: old: "+str(oldLinkNo)+" new: "+str(int(newLinkNo))+"\n")
        except:
            f.write("!!Error: link: old: "+str(oldLinkNo)+" new: "+str(int(newLinkNo))+"\n")
            print ("error "+newLinkNo)
            continue
    f.write("\n\n\n")

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

"""
#--Haltestellenbereiche--#
"""
##for hstbereich in VISUM.Net.Stops:
##    neu = int(hstbereich.AttValue("Min:StopAreas\No"))
##    alt = int(hstbereich.AttValue("No"))
##    if neu != alt:
##        hstbereich.SetAttValue("No",neu)
##        print "alt :"+str(alt)+" neu: "+str(neu)

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
VISUM.SaveVersion(Netz)

Sekunden = int(time.time() - start_time)
print("--finished after ",Sekunden,"seconds--")

f.write("--finished after "+str(Sekunden)+" seconds--")
f.close()