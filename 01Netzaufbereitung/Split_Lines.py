# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        select and split different lineroutes of different lines within one line after HAFAS-Import
# Author:      mape
# Created:     01/03/2017
# Copyright:   (c) mape 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import win32com.client.dynamic
import time
start_time = time.time()
from datetime import datetime

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(path),'VISUM_Tools','Linien_trennen.txt')
f = path.read_text()
f = f.split('\n')

#--Parameter--#
Netz = f[0]
new_line = 0

#--outputs--#
date = datetime.now()
txt_name = 'line_split_'+date.strftime('%m%d%Y')+'.txt'
f = open(Path.home() / 'Desktop' / txt_name,'w+')
f.write(time.ctime()+"\n")
f.write("Input Network: "+Netz+"\n")

f.write("\n\n\n\n")

#--VISUM--#
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.20")
VISUM.loadversion(Netz)
VISUM.Filters.InitAll()

#--cheing for all lines--#
f.write("Splitting lines \n")
for Line in VISUM.Net.Lines:
    # if "M" != Line.AttValue("Name")[0]: continue

    ##check for all lineroutes within one line
    Name = Line.AttValue("Name")
    print ("checking line: "+Name)
    VISUM.Filters.InitAll()
    Linien = VISUM.Filters.LineGroupFilter()
    Linien.UseFilterForVehJourneys = True
    Linien.UseFilterForLines = True
    Linien.UseFilterForLineRoutes = True
    Linien = Linien.LineFilter()
    Linien.AddCondition("OP_NONE",False,"Name","EqualVal",Name)
    #--build table--#
    Stops = []
    i = 1
    for Route in Line.LineRoutes:
        HP = Route.AttValue("Concatenate:StopPoints\\No")
        HP = HP.replace(",...","")
        HP = HP.replace("...","")
        HP = [int(i) for i in HP.split(",")]
        Stops.append([HP,int(Route.AttValue("ID")),i])
        i+=1
    for i in enumerate(Stops):
        for e in enumerate(Stops):
            if any(x in i[1][0] for x in e[1][0]) == True:
                if i[1][2] <= e[1][2]: Stops[e[0]][2] = i[1][2]
                else: Stops[i[0]][2] = e[1][2]
    for i in Stops:
        Route = VISUM.Net.LineRoutes.ItemByID(i[1])
        Route.SetAttValue("AddVal1",i[2])
        if i[2]>1:
            print("splitting in line: "+Name)
            f.write("splitting in line: "+Name+"\n")    

    if new_line == 0:continue
    for Route in Line.LineRoutes:
        if Route.AttValue("AddVal1") == 1: continue
        Text = Name+"("+str(Route.AttValue("AddVal1"))+")"
        try:            
            VISUM.Net.AddLine(Text,Line.AttValue("TSysCode"))
            print ("added new line: "+Text)
            f.write("added new line: "+Text+"\n")   
            Route.Line = Text
        except: #if line already exists
            Route.Line = Text

f.write("\n\n\n")

##end
Sekunden = int(time.time() - start_time)
print("--finished after ",Sekunden,"seconds--")

f.write("--finished after "+str(Sekunden)+" seconds--")
f.close()