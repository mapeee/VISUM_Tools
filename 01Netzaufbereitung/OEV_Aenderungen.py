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
import win32com.client.dynamic
import time
import random
start_time = time.clock() ##Ziel: am Ende die Berechnungsdauer ausgeben

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(r'C:'+path),'VISUM_Tools','Knotennummern.txt')
f = path.read_text()
f = f.split('\n')

#--Eingabe-Parameter--#
Netz = 'C:'+f[0]

#--VISUM öffnen--#
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.17")
VISUM.loadversion(Netz)
VISUM.Filters.InitAll()


"""
#Linienzuordnung (Vereinigung Linienrouten eigentlich gleicher Linien)
"""
#--Zuordnugn LinienRouten zu Linien--#
##l = [] ##erstelle Liste mit allen Linien, damit keine Doppelschleifen nach dem Ändern des Namens entstehen
##
##for Linien in VISUM.Net.Lines: ##Schleife über alle Linien
####    if Linien.AttValue("TSysCode") not in ["BUS"]: continue
##    Name = Linien.AttValue("Name")
##    l.append((Name))
##
##for i in l:
##    Line = VISUM.Net.Lines.ItemByKey(i)
####    if Line.AttValue("AddVal1") ==0:
####        continue
##    Name = Line.AttValue("Name")
##    Name_neu = Line.AttValue("Test")
##    try:
##        Line.SetAttValue("Name",Name_neu)
##    except:
##        for Route in Line.LineRoutes:
##            try: ##Fals es diesen Linienroutennamen schon gibt.
##                Route.Line = Name_neu
##            except:
##                Routenname = Route.AttValue("Name")
##                try:
##                    Route.SetAttValue("Name",str(int(Routenname)+random.randrange(10,800)))
##                    Route.Line = Name_neu
##                except:
##                    Route.SetAttValue("Name",str(int(Routenname)+random.randrange(10,800)))
##                    Route.Line = Name_neu
"""
#Linienrouten
"""
#--Namen Linienrouten ändern--#
##l = [] ##erstelle Liste mit allen Linien, damit keine Doppelschleifen nach dem Ändern des Namens entstehen
##for Linien in VISUM.Net.Lines: ##Schleife über alle Linien
##    if Linien.AttValue("TSysCode") not in ["BUS","AST","Fuss"]: continue
##    Name = Linien.AttValue("Name")
##    l.append((Name))
##
##for i in l:
##    Linien = VISUM.Net.Lines.ItemByKey(i)
##
##    for Routen in Linien.LineRoutes: ##Schleife über alle Linienrouten dieser Linie
##        Start = Routen.AttValue("StartLineRouteItem\StopPoint\Name")
##        Ende = Routen.AttValue("EndLineRouteItem\StopPoint\Name")
##        Nr = Routen.AttValue("Name") ##Nummer aus den Hafas-Daten
##        Neu = Start+"--"+Ende+"("+Nr+")"
##        Routen.SetAttValue("Name",Neu)

##        for Knoten in Routen.LineRouteItems:
##            print Knoten.AttValue("NodeNo")


"""
#Linien
"""
##VISUM.Filters.InitAll() ##Um jeweils die längste Linienroute zu erhalten!
##Linien = VISUM.Filters.LineGroupFilter()
##Linien.UseFilterForLineRoutes = True ##Wirksam für Fahrplanfahrt-Abschnitte
##Linien = Linien.LineRouteFilter()
##Linien.AddCondition("OP_NONE",False,"Length","EqualAtt","Line\Max:LineRoutes\Length")##Kein weiterer Filter,nicht komplementär,Szenario,GleichWert,Wert=1
###--Liniennamen--#
##l = [] ##erstelle Liste mit allen Linien, damit keine Doppelschleifen nach dem Ändern des Namens entstehen
##for Linien in VISUM.Net.Lines: ##Schleife über alle Linien
##    if Linien.AttValue("TSysCode") in ["BUS","AST","Fuss"]: continue
##    Name = Linien.AttValue("Name")
##    l.append((Name))
##
##for i in l:
##    Line = VISUM.Net.Lines.ItemByKey(i)
##    Name = Line.AttValue("Name")
##    Vsys = Line.AttValue("TSysCode")
##    if Vsys in ("U","S","RV","W","FV"):
##        continue
##
##    print Name
##    Betr = Line.AttValue("Operator\Name")
##    Betr = Betr.split(" ")[0] ##Um den ganzen Namen abzutrennen.
##    if "STB" == Vsys:
##        Betr = Betr+"(STB)"
##    if "AST" == Vsys:
##        Betr = Betr+"(AST)"
##    if "ALT" == Vsys:
##        Betr = Betr+"(ALT)"
##    try:
##        Neu = Betr+"--"+Name.split("_")[1]
##    except:
##        Neu = Betr+"--"+Name.split("_")[0]
##    Neu = Name.split(" ")[0]
##    Neu = Line.AttValue("Test")
##
##
##
##    Neu += " ("+Line.AttValue("MinActive:LineRoutes\StartLineRouteItem\StopPoint\Name")+"--"+Line.AttValue("MaxActive:LineRoutes\EndLineRouteItem\StopPoint\Name")+")"
##    print Neu
##
##    try:
##        Line.SetAttValue("Name",Neu)
##    except:
##        print "Fehler!!!"
##
##VISUM.Filters.InitAll() ##Damit dieser Filter später keine Relevanz mehr hat

"""
#--Fahrplanfahrten--#
"""
###Verkehrstage#
##l = [] ##erstelle Liste mit allen Tagen, damit keine Doppelschleifen nach dem Ändern des Namens entstehen
##for Tag in VISUM.Net.ValidDaysCont: ##Schleife über alle Linien
##    Tage = Tag.AttValue("No")
##    l.append((Tage))
##
##for i in l:
##    Tag = VISUM.Net.ValidDaysCont.ItemByKey(i)
##    Nummer = int(Tag.AttValue("No"))
##    if Nummer==1:continue
##    Nummer_neu = Nummer+29000
##    try:
##        Tag.SetAttValue("No",Nummer_neu)
##        print "neu" +str(Nummer_neu)
##    except:
##        print "alt" +str(Nummer)
##        pass


##for Fahrten in VISUM.Net.VehicleJourneys:
##    VTneu = Fahrten.AttValue("AddVal1")
##    if int(VTneu) != 2:
##        continue
##    for section in Fahrten.VehicleJourneySections:
##        if section.AttValue("ValidDaysNo") != VTneu:
##            section.SetAttValue("ValidDaysNo",VTneu)



#Namen#
##for Fahrten in VISUM.Net.VehicleJourneys:
##    von = Fahrten.AttValue("FromStopPoint\Name")
##    nach = Fahrten.AttValue("ToStopPoint\Name")
##    Nr = int(Fahrten.AttValue("No")) ##Nummer aus Hafas-Daten
##    Neu = von+"--"+nach+"("+str(Nr)+")"
##    Fahrten.SetAttValue("Name",Neu)

#Ende

"""
#--Haltestellenbereiche--#
"""
##for hstbereich in VISUM.Net.Stops:
##    neu = int(hstbereich.AttValue("Min:StopAreas\No"))
##    alt = int(hstbereich.AttValue("No"))
##    if neu != alt:
##        hstbereich.SetAttValue("No",neu)
##        print "alt :"+str(alt)+" neu: "+str(neu)


Sekunden = int(time.clock() - start_time)

print "--Scriptdurchlauf erfolgreich nach",Sekunden,"Sekunden!--"