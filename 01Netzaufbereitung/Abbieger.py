# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        editing turns
# Purpose:
#
# Author:      mape
#
# Created:     25/09/2020
# Copyright:   (c) mape 2020
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import win32com.client.dynamic
import xlsxwriter
import time
import xlrd
start_time = time.time()
from datetime import datetime

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(path),'VISUM_Tools','Turns.txt')
f = path.read_text()
f = f.split('\n')

#--Parameter--#
Netz = f[0]
Sharp = 0 #writing number of all sharp turns to textfile
direct = 1 #writing value directly into VISUM network

#--excel--#
# wb = xlwt.Workbook()
# ws = wb.add_sheet(f[1]+"x")
# results = 'C:'+f[2]

path_out = 'C:'+f[2]
wb = xlsxwriter.Workbook(path_out)
ws = wb.add_worksheet(f[1])


RowNo = 0
#head#
column = 0
ws.write(RowNo, column, "FromNodeNo")
ws.write(RowNo, column+1, "ViaNodeNo")
ws.write(RowNo, column+2, "ToNodeNo")
ws.write(RowNo, column+3, "TypeNo")
ws.write(RowNo, column+4, "Kapazitaet_Nachm")
ws.write(RowNo, column+5, "Kapazitaet_Nacht")
ws.write(RowNo, column+6, "Kapazitaet_Rest")
ws.write(RowNo, column+7, "Kapazitaet_Tag")
ws.write(RowNo, column+8, "Kapazitaet_Vorm")
ws.write(RowNo, column+9, "SPURIGKEIT_abb")
ws.write(RowNo, column+10, "t0_Nachm")
ws.write(RowNo, column+11, "t0_Nacht")
ws.write(RowNo, column+12, "t0_Rest")
ws.write(RowNo, column+13, "t0_Tag")
ws.write(RowNo, column+14, "t0_Vorm")
ws.write(RowNo, column+15, "TSys")

#--Functions--#
def turnInfo(node):
    fromType = node.ViaNodeTurns.GetMultiAttValues("FromLink\DISPLAYTYPE")
    toType = node.ViaNodeTurns.GetMultiAttValues("ToLink\DISPLAYTYPE")
    angle = node.ViaNodeTurns.GetMultiAttValues("Angle")
    fromTSys = node.ViaNodeTurns.GetMultiAttValues("FromLink\TSysSet")
    toTSys = node.ViaNodeTurns.GetMultiAttValues("ToLink\TSysSet")
    fromNode = node.ViaNodeTurns.GetMultiAttValues("FromNodeNo")
    viaNode = node.ViaNodeTurns.GetMultiAttValues("ViaNodeNo")
    toNode = node.ViaNodeTurns.GetMultiAttValues("ToNodeNo")
    fromNodeType = node.ViaNodeTurns.GetMultiAttValues("FromNode\Kreuzungstyp")
    NodeType = node.ViaNodeTurns.GetMultiAttValues("ViaNode\Kreuzungstyp")
    toNodeType = node.ViaNodeTurns.GetMultiAttValues("ToNode\Kreuzungstyp")
    turnType = node.ViaNodeTurns.GetMultiAttValues("TypeNo")
    infos=[]
    for i in enumerate(fromType):
        infos.append([fromType[i[0]][1],toType[i[0]][1],angle[i[0]][1],fromTSys[i[0]][1],toTSys[i[0]][1],
                  fromNode[i[0]][1],viaNode[i[0]][1],toNode[i[0]][1],fromNodeType[i[0]][1],NodeType[i[0]][1],
                  toNodeType[i[0]][1],turnType[i[0]][1]])
    return infos

def TSysIdent(turn, fromTSys, toTSys): ##for valid transport systems
    fromT = fromTSys.split(",")
    toT = toTSys.split(",")
    TSys = []
    for T in fromT:
        if T in toT:TSys.append(T)
    TSys = ','.join(TSys)
    if TSys == "": return False, TSys
    else: return True, TSys

def setValues(turn,values,rowNo):
    column = 0
    ws.write(rowNo, column, turn[5])#"FromNodeNo")
    ws.write(rowNo, column+1, turn[6])#"ViaNodeNo")
    ws.write(rowNo, column+2, turn[7])# "ToNodeNo")
    ws.write(rowNo, column+3, values[0])#"TypeNo")
    ws.write(rowNo, column+4, values[1])#"Kapazitaet_Nachm")
    ws.write(rowNo, column+5, values[1])#"Kapazitaet_Nacht")
    ws.write(rowNo, column+6, values[1])#"Kapazitaet_Rest")
    ws.write(rowNo, column+7, values[1])#"Kapazitaet_Tag")
    ws.write(rowNo, column+8, values[1])#"Kapazitaet_Vorm")
    ws.write(rowNo, column+9, values[2])#"SPURIGKEIT_abb")
    ws.write(rowNo, column+10, values[3])#"t0_Nachm")
    ws.write(rowNo, column+11, values[3])#"t0_Nacht")
    ws.write(rowNo, column+12, values[3])#"t0_Rest")
    ws.write(rowNo, column+13, values[3])#"t0_Tag")
    ws.write(rowNo, column+14, values[3])#"t0_Vorm")  
    ws.write(rowNo, column+15, values[4])#"TSys")
    
def UTurn(turn, RowNo):
    Ident = TSysIdent(turn,turn[3],turn[4])
    if turn[5] == turn[7]:
        if turn[9] == 4:setValues(turn,[0,0,0,0,""],RowNo) ##roundabout
        if turn[9] == 10:setValues(turn,[0,0,0,0,""],RowNo) ##connection
        if turn[9] == 12:setValues(turn,[0,0,0,0,""],RowNo) ##connector
        if turn[9] == 11:setValues(turn,[0,0,0,0,""],RowNo) ##Remaining
        if turn[9] == 13:setValues(turn,[0,0,0,0,""],RowNo) ##Ramp-only
        if turn[9] == 14:setValues(turn,[1,2000,0,0,Ident[1]],RowNo) ##Rail
        if turn[9] == 15:setValues(turn,[0,0,0,0,""],RowNo) ##connection
        if turn[9] < 4 or turn[9] == 5:setValues(turn,[3,20,0.2,22,Ident[1]],RowNo) ##priority to the right or traffic light
        if 6 <= turn[9] <= 7:setValues(turn,[3,20,0.2,22,Ident[1]],RowNo) ##Main-Priority-Lane
        if 8 <= turn[9] <= 9:setValues(turn,[0,0,0,0,""],RowNo) ##Ramp
        return True
    else: return False
    
#--outputs--#
date = datetime.now()
txt_name = 'SharpAngle_'+date.strftime('%m%d%Y')+'.txt'
if Sharp == 1:
    f = open(Path.home() / 'Desktop' / txt_name,'w+')
    f.write(time.ctime()+"\n")
    f.write("Input Network: "+Netz+"\n")
    f.write("\n\n\n\n")

#--open VISUM--#
VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.20")
VISUM.loadversion(Netz)
VISUM.Filters.InitAll()

#--nodes--#
nodes = VISUM.Net.Nodes.GetMultiAttValues("No")
area = VISUM.Net.Nodes.GetMultiAttValues("Lage_Untersuchungsraum")
nodeType = VISUM.Net.Nodes.GetMultiAttValues("Kreuzungstyp")
nLinks = VISUM.Net.Nodes.GetMultiAttValues("NumLinks")
minType = VISUM.Net.Nodes.GetMultiAttValues("Min:InLinks\DISPLAYTYPE")
maxType = VISUM.Net.Nodes.GetMultiAttValues("Max:InLinks\DISPLAYTYPE")
Val1 = VISUM.Net.Nodes.GetMultiAttValues("AddVal1")

for i in enumerate(nodes):
    # if area[i[0]][1] == 1:continue  ##not in focus area
    if Val1[i[0]][1] != 1:continue  ##only some nodes to change
    # if minType[i[0]][1] != 1:continue  ##highway only
    # if minType[i[0]][1] != 2 and minType[i[0]][1] != 3:continue  ##trunk only
    # if nodeType[i[0]][1] != 1 and nodeType[i[0]][1] != 6 and nodeType[i[0]][1] != 7: continue
    # if nodeType[i[0]][1] != 12 and nodeType[i[0]][1] != 5: continue
    
    if Sharp == 1:
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            if (turn[2] <= 5 or turn[2] > 355) and turn[2] != 360:
                f.write(str(int(i[1][1]))+"\n")
                break
        continue
                
    #LSA#
    if nodeType[i[0]][1] == 1:        
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            RowNo+=1
            Ident = TSysIdent(turn,turn[3],turn[4])
            ##UTurns
            if UTurn(turn, RowNo) is True: continue
            ##closed link for car and walk/bike
            if Ident[0] is False:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ## highway  only
            if maxType[i[0]][1] <= 1:
                setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                continue
            ##sharp angles
            if turn[2] <= 15 or turn[2] > 345:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##right
            if turn[2] > 15 and turn[2] <= 135:
                if turn[0] <= 3: ##from trunk road
                    setValues(turn,[1,450,0.6,18,Ident[1]],RowNo)
                    continue
                if turn[0] <= 4: ##from main
                    setValues(turn,[1,350,0.5,18,Ident[1]],RowNo)
                    continue
                else:
                    setValues(turn,[1,210,0.3,18,Ident[1]],RowNo) ##from minor
                    continue                
            ##straight
            if turn[2] > 135 and turn[2] <= 225:
                if turn[0] <= 2: ##from trunk road
                    setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                    continue
                if turn[0] <= 3: ##from main road
                    setValues(turn,[2,800,1.0,16,Ident[1]],RowNo)
                    continue
                if turn[0] <= 4: ##from main
                    setValues(turn,[2,640,0.8,16,Ident[1]],RowNo)
                    continue
                else:
                    setValues(turn,[2,400,0.5,16,Ident[1]],RowNo) ##from minor
                    continue
            ##left
            if turn[2] > 225 and turn[2] <= 345:
                if turn[0] <= 3: ##from main
                    setValues(turn,[3,450,1.0,22,Ident[1]],RowNo)
                    continue
                if turn[0] <= 4: ##from main
                    setValues(turn,[3,225,0.5,22,Ident[1]],RowNo)
                    continue
                else:
                    setValues(turn,[3,180,0.4,22,Ident[1]],RowNo) ##from minor
                    continue    
            ##remaining values    
            setValues(turn,[2,99999,999,0,Ident[1]],RowNo)
        continue
            
    #LSA minor and walk#
    if 2 <= nodeType[i[0]][1] <= 3:        
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            RowNo+=1
            Ident = TSysIdent(turn,turn[3],turn[4])
            ##UTurns
            if UTurn(turn, RowNo) is True: continue
            ##closed link for car and walk/bike
            if Ident[0] is False:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##remaining
            if turn[0] <= 4: #major road
                setValues(turn,[7,900,1.5,2,Ident[1]],RowNo)
                continue
            else: setValues(turn,[7,600,1.0,2,Ident[1]],RowNo)
            continue
        continue
            
    #Roundabout#
    if nodeType[i[0]][1] == 4:        
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            RowNo+=1
            Ident = TSysIdent(turn,turn[3],turn[4])
            ##UTurns
            if UTurn(turn, RowNo) is True: continue
            ##closed link for car and walk/bike
            if Ident[0] is False:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##go into
            if turn[8] != 4:
                if turn[0] <= 4: setValues(turn,[5,950,1.0,12,Ident[1]],RowNo)
                else: setValues(turn,[5,475,0.5,12,Ident[1]],RowNo)
                continue
            ##go straight
            if turn[8] == 4 and turn[10] == 4:
                if turn[0] <= 4: setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                else: setValues(turn,[9,99999,0.5,0,Ident[1]],RowNo)
                continue
            ##go out
            else:
                setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                continue  
        continue
    
    #Right-priority-lane#
    if nodeType[i[0]][1] == 5:        
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            RowNo+=1
            Ident = TSysIdent(turn,turn[3],turn[4])
            ##UTurns
            if UTurn(turn, RowNo) is True: continue
            ## main roads only
            if maxType[i[0]][1] <= 3:
                setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                continue
            ##sharp angles
            if turn[2] <= 15 or turn[2] > 345:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##closed link for car and walk/bike
            if Ident[0] is False:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##right
            if turn[2] > 15 and turn[2] <= 135:
                setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                continue
            ##remaining mains
            if turn[0] == 4 and turn[1] == 4: ##main roads (level 4)
                setValues(turn,[4,950,1.0,7,Ident[1]],RowNo)
                continue
            ##remaining values
            setValues(turn,[4,475,0.5,7,Ident[1]],RowNo)
        continue

    #Main-prioriry-lane#
    if 6 <= nodeType[i[0]][1] <= 7:        
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            RowNo+=1
            Ident = TSysIdent(turn,turn[3],turn[4])
            ##UTurns
            if UTurn(turn, RowNo) is True: continue
            ##closed link for car and walk/bike
            if Ident[0] is False:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ## main roads only
            if maxType[i[0]][1] <= 3:
                setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                continue
            ##sharp angles
            if turn[2] <= 15 or turn[2] > 345:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##right
            if turn[2] > 15 and turn[2] <= 135:
                if turn[0] == minType[i[0]][1]: ##from main
                    if turn[0] <= 4: setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                    else: setValues(turn,[5,475,0.5,12,Ident[1]],RowNo)
                    continue
                else:
                    if turn[0] <= 4:
                        setValues(turn,[5,475,0.5,12,Ident[1]],RowNo) ##from minor prior
                        continue
                    else:
                        setValues(turn,[5,285,0.3,12,Ident[1]],RowNo) ##from minor
                        continue
            ##straight
            if turn[2] > 135 and turn[2] <= 225:
                if turn[0] == minType[i[0]][1]: ##from main
                    setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                    continue
                else:
                    setValues(turn,[6,388,0.5,18,Ident[1]],RowNo) ##from minor
                    continue
            ##left
            if turn[2] > 225 and turn[2] <= 345:
                if turn[0] == minType[i[0]][1]: ##from main
                    if turn[0] == minType[i[0]][1] and turn[1] == minType[i[0]][1]: setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo) #bend                
                    elif turn[0] <= 3 and turn[1] <= 4: setValues(turn,[4,380,0.4,7,Ident[1]],RowNo)
                    else: setValues(turn,[4,380,0.4,7,Ident[1]],RowNo)
                    continue
                else:
                    if turn[0] <= 4:
                        setValues(turn,[6,310,0.4,18,Ident[1]],RowNo) ##from minor prior
                        continue  
                    else:
                        setValues(turn,[6,155,0.2,18,Ident[1]],RowNo) ##from minor
                        continue 
            ##remaining values    
            setValues(turn,[2,99999,999,0,Ident[1]],RowNo)
        continue

    #Ramp#
    if nodeType[i[0]][1] == 8:        
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            RowNo+=1
            Ident = TSysIdent(turn,turn[3],turn[4])
            ##UTurns
            if UTurn(turn, RowNo) is True: continue
            ##sharp angles
            if turn[2] <= 100 or turn[2] > 260:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##closed link for car and walk/bike
            if Ident[0] is False:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##from fast to fast
            if turn[0] <= 4 and turn[1] <= 4:
                setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                continue
            ##from slow to fast
            if turn[0] > 3 and turn[1] <= 3:
                setValues(turn,[8,325,1.0,3,Ident[1]],RowNo)
                continue
            ##from slow to slow
            if turn[0] > 3 and turn[1] > 3:
                setValues(turn,[8,325,1.0,3,Ident[1]],RowNo) ##from minor
                continue
            ##from fast to slow
            if turn[0] <= 3 and turn[1] > 3:
                setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo) ##from minor
                continue
            ##remaining values    
            setValues(turn,[2,99999,999,0,Ident[1]],RowNo)
        continue
    
    #fast-fast#
    if nodeType[i[0]][1] == 9:        
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            RowNo+=1
            Ident = TSysIdent(turn,turn[3],turn[4])
            ##UTurns
            if UTurn(turn, RowNo) is True: continue
            ##sharp angles
            if turn[2] <= 45 or turn[2] > 315:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##closed link for car and walk/bike
            if Ident[0] is False:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##from fast to fast
            if turn[0] <= 4 and turn[1] <= 4:
                setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                continue
            ##remaining values    
            setValues(turn,[2,99999,999,0,Ident[1]],RowNo)
        continue
    
    #connections#
    if nodeType[i[0]][1] == 10:        
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            RowNo+=1
            Ident = TSysIdent(turn,turn[3],turn[4])
            ##UTurns
            if UTurn(turn, RowNo) is True: continue
            ##closed link for car and walk/bike
            if Ident[0] is False:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##remaining values
            setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
        continue    

    #Connector#
    if nodeType[i[0]][1] == 12:        
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            RowNo+=1
            Ident = TSysIdent(turn,turn[3],turn[4])
            ##UTurns
            if UTurn(turn, RowNo) is True: continue
            ## main roads only
            if maxType[i[0]][1] <= 3:
                setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                continue
            ##connector
            if turn[0] == 10 or turn[1] == 10:
                setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                continue
            ##sharp angles
            if turn[2] <= 15 or turn[2] > 345:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##closed link for car and walk/bike
            if Ident[0] is False:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##right
            if turn[2] > 15 and turn[2] <= 135:
                if turn[0] == minType[i[0]][1]: ##from main
                    setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                    continue
                else:
                    setValues(turn,[5,475,0.5,12,Ident[1]],RowNo) ##from minor
                    continue
            ##straight
            if turn[2] > 135 and turn[2] <= 225:
                if turn[0] == minType[i[0]][1]: ##from main
                    setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
                    continue
                else:
                    setValues(turn,[6,388,0.5,18,Ident[1]],RowNo) ##from minor
                    continue
            ##left
            if turn[2] > 225 and turn[2] <= 345:
                if turn[0] == minType[i[0]][1]: ##from main
                    if turn[0] == minType[i[0]][1]: setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo) #bend                
                    else: setValues(turn,[4,475,0.5,7,Ident[1]],RowNo)
                    continue
                else:
                    setValues(turn,[6,310,0.4,18,Ident[1]],RowNo) ##from minor
                    continue    
            ##remaining values    
            setValues(turn,[2,99999,999,0,Ident[1]],RowNo)
        continue

    #Ramp-only#
    if nodeType[i[0]][1] == 13:        
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            RowNo+=1
            Ident = TSysIdent(turn,turn[3],turn[4])
            ##UTurns
            if UTurn(turn, RowNo) is True: continue
            ##closed link for car and walk/bike
            if Ident[0] is False:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##remaining
            setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
        continue
    
    #Raily#
    if nodeType[i[0]][1] == 14:        
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            RowNo+=1
            Ident = TSysIdent(turn,turn[3],turn[4])
            ##UTurns
            if UTurn(turn, RowNo) is True: continue
            ##sharp angles
            if turn[2] <= 45 or turn[2] > 315:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##closed link for rail systems
            if Ident[0] is False:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##remaining
            setValues(turn,[turn[11],2000,0.0,0,Ident[1]],RowNo)
        continue 
    
    #stops#
    if nodeType[i[0]][1] == 15:        
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            RowNo+=1
            Ident = TSysIdent(turn,turn[3],turn[4])
            ##UTurns
            if UTurn(turn, RowNo) is True: continue
            ##closed link for car and walk/bike
            if Ident[0] is False:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##remaining values
            setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
        continue
    
    #Remaining#
    if nodeType[i[0]][1] == 11:        
        node = VISUM.Net.Nodes.ItemByKey(i[1][1])
        turns = turnInfo(node)
        for turn in turns:
            RowNo+=1
            Ident = TSysIdent(turn,turn[3],turn[4])
            ##UTurns
            if UTurn(turn, RowNo) is True: continue
            ##closed link for car and walk/bike
            if Ident[0] is False:
                setValues(turn,[0,0,0.0,0,""],RowNo)
                continue
            ##remaining
            setValues(turn,[9,99999,1.0,0,Ident[1]],RowNo)
        continue 

#excel output#
# wb.save(results)
wb.close()

#writing directly#
if direct == 1:
    book = xlrd.open_workbook(path_out)
    sheet = book.sheet_by_index(0) ##getting first sheet (0)
    turns = sheet.nrows-1 ##-1 due to headlines
    for i in range(turns):
        from_n = int(sheet.cell_value(rowx=1+i,colx=0))
        via_n = int(sheet.cell_value(rowx=1+i,colx=1))
        to_n = int(sheet.cell_value(rowx=1+i,colx=2))
        type_n = sheet.cell_value(rowx=1+i,colx=3)
        capa = sheet.cell_value(rowx=1+i,colx=4)
        Lanes = sheet.cell_value(rowx=1+i,colx=9)
        t0 = sheet.cell_value(rowx=1+i,colx=10)
        Vsys = sheet.cell_value(rowx=1+i,colx=15)
        
        turn = VISUM.Net.Turns.ItemByKey(from_n,via_n,to_n)
        
        turn.SetAttValue("TypeNo", type_n)
        turn.SetAttValue("TSysSet", Vsys)
        turn.SetAttValue("Kapazitaet_Nachm", capa)
        turn.SetAttValue("Kapazitaet_Tag", capa)
        turn.SetAttValue("Kapazitaet_Vorm", capa)
        turn.SetAttValue("Kapazitaet_Rest", capa)
        turn.SetAttValue("Kapazitaet_Nacht", capa)
        turn.SetAttValue("SPURIGKEIT_abb", Lanes)
        turn.SetAttValue("t0_Nachm", t0)
        turn.SetAttValue("t0_Tag", t0)
        turn.SetAttValue("t0_Vorm", t0)
        turn.SetAttValue("t0_Rest", t0)
        turn.SetAttValue("t0_Nacht", t0)


       
    
##end
Sekunden = int(time.time() - start_time)
print("--finished after ",Sekunden,"seconds--")
if Sharp == 1:
    f.write("--finished after "+str(Sekunden)+" seconds--")
    f.close()