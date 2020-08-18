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
start_time = time.time()
from datetime import datetime
import random

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(path),'VISUM_Tools','Nummern.txt')
f = path.read_text()
f = f.split('\n')

#--Eingabe-Parameter--#
Netz = f[0]
Name_LineRoute = 1
Name_Line = 1
ValidDay = 0
VehicleJourney = 1


#--outputs--#
date = datetime.now()
txt_name = 'OEV_Aenderungen_'+date.strftime('%m%d%Y')+'.txt'
f = open(Path.home() / 'Desktop' / txt_name,'w+')
f.write(time.ctime()+"\n")
f.write("Input Network: "+Netz+"\n")

f.write("\n\n\n\n")


#--VISUM öffnen--#
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.20")
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
#Lineroutes
"""
if Name_LineRoute == 1:
    f.write("LineRoute changes \n")
    l = [] ##list of lines to avoid references to already changed attributes
    for Linien in VISUM.Net.Lines: 
        if Linien.AttValue("TSysCode") not in ["BUS","AST"]: continue
        Name = Linien.AttValue("Name")
        l.append((Name))
    
    for i in l:
        Linien = VISUM.Net.Lines.ItemByKey(i)
    
        for Routen in Linien.LineRoutes: ##loop of lineroutes
            Start = Routen.AttValue(r"StartLineRouteItem\StopPoint\Name")
            Ende = Routen.AttValue(r"EndLineRouteItem\StopPoint\Name")
            Nr = Routen.AttValue("Name")
            Neu = Start+"--"+Ende+"("+Nr+")"
            Routen.SetAttValue("Name",Neu)
            f.write("lineroute: old: "+str(Nr)+" new: "+str(Neu)+"\n")
    f.write("\n\n\n")

"""
#Lines
"""
if Name_Line == 1:
    f.write("Line changes \n")
    VISUM.Filters.InitAll() 
    Linien = VISUM.Filters.LineGroupFilter() ##To get the longest lineroute
    Linien.UseFilterForLineRoutes = True ##for all timetable elements
    Linien = Linien.LineRouteFilter()
    Linien.AddCondition("OP_NONE",False,"Length","EqualAtt","Line\Max:LineRoutes\Length")##no more filter,not complementary,szenario,isValue,Vaklue=1
    #--Linename--#
    l = [] ##list of lines to avoid references to already changed attributes
    for Linien in VISUM.Net.Lines: ##loop of lines
        if Linien.AttValue("TSysCode") not in ["BUS","AST"]: continue
        Name = Linien.AttValue("Name")
        l.append((Name))
    
    for i in l:
        Line = VISUM.Net.Lines.ItemByKey(i)
        Name = Line.AttValue("Name")
        Vsys = Line.AttValue("TSysCode")
        if Vsys == "AST": Name = "AST"+Name
        # Betr = Line.AttValue("Operator\Name")
        # Betr = Betr.split(" ")[0]
        try:
            Neu = Name.split(",")[1]
        except:
            Neu = Name.split(",")[0]       
        Neu += " ("+Line.AttValue(r"MinActive:LineRoutes\StartLineRouteItem\StopPoint\Name")+"--"+Line.AttValue(r"MaxActive:LineRoutes\EndLineRouteItem\StopPoint\Name")+")"
 
        try:
            Line.SetAttValue("Name",Neu)
            f.write("line: old: "+str(Name)+" new: "+str(Neu)+"\n")  
        except:
            f.write("!!Error: line: old: "+str(Name)+" new: "+str(Neu)+"\n")
            print ("error "+Neu)
    
    VISUM.Filters.InitAll() ##Damit dieser Filter später keine Relevanz mehr hat
    f.write("\n\n\n")

"""
#--ValidDays--#
"""
if ValidDay == 1:
    f.write("ValidDay changes \n")
    l = [] ##list of lines to avoid references to already changed attributes
    for Tag in VISUM.Net.ValidDaysCont: ##loop of lines
        Tage = Tag.AttValue("No")
        l.append((Tage))
    
    for i in l:
        Tag = VISUM.Net.ValidDaysCont.ItemByKey(i)
        Nummer = int(Tag.AttValue("No"))
        if Nummer==1:continue
        Nummer_neu = Nummer+29000
        try:
            Tag.SetAttValue("No",Nummer_neu)
            f.write("ValidDay: old: "+str(Nummer)+" new: "+str(Nummer_neu)+"\n")  
        except:pass   
    
    for Fahrten in VISUM.Net.VehicleJourneys:
        VTneu = Fahrten.AttValue("AddVal1")
        for section in Fahrten.VehicleJourneySections:
            if section.AttValue("ValidDaysNo") != VTneu:
                section.SetAttValue("ValidDaysNo",VTneu)          
    f.write("\n\n\n")

"""
#--VehicleJourneys--#
"""
if VehicleJourney == 1:
    f.write("VehhicleJourney changes \n")
    for Fahrten in VISUM.Net.VehicleJourneys:
        von = Fahrten.AttValue(r"FromStopPoint\Name")
        nach = Fahrten.AttValue(r"ToStopPoint\Name")
        Nr = int(Fahrten.AttValue("No")) ##Number from HAFAS-Data
        Neu = von+"--"+nach+"("+str(Nr)+")"
        f.write("VehicleJourney: old: "+str(Fahrten.AttValue("Name")+" new: "+str(Neu)+"\n"))      
        Fahrten.SetAttValue("Name",Neu)
    f.write("\n\n\n")

##end
Sekunden = int(time.time() - start_time)
print("--finished after ",Sekunden,"seconds--")

f.write("--finished after "+str(Sekunden)+" seconds--")
f.close()