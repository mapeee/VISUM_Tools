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
Name_Line = 1
Name_LineRoute = 0
Lineroute_Line = 0
ValidDay = 0
VehicleJourney = 0

Vsys_c = ["Bus","AST"]


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
#Lineroute to Line
"""
if Lineroute_Line == 1:
    f.write("Lineroute to Line changes \n")
    l = [] ##list of lines to avoid references to already changed attributes
    for Linien in VISUM.Net.Lines: ##Schleife über alle Linien
        if Linien.AttValue("TSysCode") not in Vsys_c: continue
        Name = Linien.AttValue("Name")
        l.append((Name))
    
    for i in l:
        Linien = VISUM.Net.Lines.ItemByKey(i)
        for Route in Linien.LineRoutes: ##loop of lineroutes
            if Route.AttValue("Line_neu") == "no": continue
            Name_new = Route.AttValue("Line_neu")
            try: Route.Line = Name_new
            except:
                Routenname = Route.AttValue("Name")                    
                Route.SetAttValue("Name",str(Routenname)+" ("+str(random.randrange(100,800))+")")
                Route.Line = Name_new
               
            f.write("line of lineroute: old: "+str(Route.AttValue("LineName"))+" new: "+str(Name_new)+"\n") 
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
        if Linien.AttValue("TSysCode") not in Vsys_c: continue
        Name = Linien.AttValue("Name")
        l.append((Name))
    
    for i in l:
        Line = VISUM.Net.Lines.ItemByKey(i)
        Name = Line.AttValue("Name")
        Vsys = Line.AttValue("TSysCode")
        # if Vsys == "AST": Name = "AST"+Name
        # Betr = Line.AttValue("Operator\Name")
        # Betr = Betr.split(" ")[0]
        try: Neu = Name.split(",")[1]
        except: Neu = Name.split(",")[0]       
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
#Lineroutes
"""
if Name_LineRoute == 1:
    f.write("LineRoute changes \n")
    l = [] ##list of lines to avoid references to already changed attributes
    for Linien in VISUM.Net.Lines: 
        if Linien.AttValue("TSysCode") not in Vsys_c: continue
        Name = Linien.AttValue("Name")
        l.append((Name))
    
    for i in l:
        Linien = VISUM.Net.Lines.ItemByKey(i)
        n = 1
        for Routen in Linien.LineRoutes: ##loop of lineroutes
            Start = Routen.AttValue(r"StartLineRouteItem\StopPoint\Name")
            Ende = Routen.AttValue(r"EndLineRouteItem\StopPoint\Name")
            Nr = Routen.AttValue("Name")
            Neu = Start+"--"+Ende+"("+str(n)+")"
            try: Routen.SetAttValue("Name",Neu)
            except: pass
            f.write("lineroute: old: "+str(Nr)+" new: "+str(Neu)+"\n")
            n+=1
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
        if Fahrten.AttValue("TSysCode") not in Vsys_c: continue
        von = Fahrten.AttValue(r"FromStopPoint\Name")
        nach = Fahrten.AttValue(r"ToStopPoint\Name")
        Neu = von+"--"+nach
        f.write("VehicleJourney: old: "+str(Fahrten.AttValue("Name")+" new: "+str(Neu)+"\n"))      
        Fahrten.SetAttValue("Name",Neu)
    f.write("\n\n\n")

##end
Sekunden = int(time.time() - start_time)
print("--finished after ",Sekunden,"seconds--")

f.write("--finished after "+str(Sekunden)+" seconds--")
f.close()