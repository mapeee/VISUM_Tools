# -*- coding: utf-8 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        compare and change numbers in two visum networks
# Purpose:
#
# Author:      mape
#
# Created:     03/08/2020
# Copyright:   (c) mape 2020
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import win32com.client.dynamic
import numpy as np

import time
start_time = time.time()

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(path),'VISUM_Tools','Abgleich.txt')
f = path.read_text()
f = f.split('\n')

#--Parameter--#
in_net = f[1]
to_change = f[0]

Node = True
Link = True

#--outputs--#
f = open("Nummernabgleich.txt","w+")
f.write(time.ctime()+"\n")
f.write("Input Network: "+in_net+"\n")
f.write("Change Network: "+to_change+"\n")

f.write("\n\n\n\n")

#--open VISUM input--#
##this is the VISUM version not to change
VISUM_in = win32com.client.dynamic.Dispatch("Visum.Visum.20")
VISUM_in.loadversion(in_net)
VISUM_in.Filters.InitAll()

#--numbers in version--#
NodeNr_in = np.array(VISUM_in.Net.Nodes.GetMultiAttValues("No")).astype("int")[:,1]
LinkNr_in = np.array(VISUM_in.Net.Links.GetMultiAttValues("No")).astype("int")[:,1]
del VISUM_in


#--open VISUM change--#
##this is the VISUM version to change
VISUM_change = win32com.client.dynamic.Dispatch("Visum.Visum.20")
VISUM_change.loadversion(to_change)
VISUM_change.Filters.InitAll()

#--changes--#
#NODES#
if Node == True:
    f.write("Node changes \n")
    NodeNr_change = np.array(VISUM_change.Net.Nodes.GetMultiAttValues("No")).astype("int")[:,1]

    for i in NodeNr_change:
        if i in NodeNr_in:
            for random in np.random.randint(5000000, size=100):
                Node_new = i+random
                if np.any(NodeNr_in==Node_new) == False and np.any(NodeNr_change==Node_new) == False:
                    try: VISUM_change.Net.Nodes.ItemByKey(i).SetAttValue("No",Node_new)
                    except: continue ##in case of modified nodes before but not changed in NodeNr_change
                    f.write("node: old: "+str(i)+" new: "+str(Node_new)+"\n")
                    break
    f.write("\n\n\n")

#LINKS#
if Link == True:
    f.write("Link changes \n")
    LinkNr_change = np.array(VISUM_change.Net.Links.GetMultiAttValues("No")).astype("int")
    Link_from = np.array(VISUM_change.Net.Links.GetMultiAttValues("FromNodeNo")).astype("int")
    Link_to = np.array(VISUM_change.Net.Links.GetMultiAttValues("ToNodeNo")).astype("int")
    Link_change = np.concatenate([LinkNr_change, Link_from,Link_to], axis=1)

    Nr = 0
    for i in Link_change:
        if i[1] == Nr:continue ##for different directions
        Nr = i[1]
        if Nr in LinkNr_in:
            for random in np.random.randint(5000000, size=100):
                Link_new = Nr+random
                if np.any(LinkNr_in==Link_new) == False and np.any(LinkNr_change==Link_new) == False:
                    try: VISUM_change.Net.Links.ItemByKey(i[3],i[5]).SetNo(Link_new)
                    except: continue ##in case of modified links before but not changed in LinkNr_change
                    f.write("link: old: "+str(Nr)+" new: "+str(Link_new)+"\n")
                    break
    f.write("\n\n\n")
    
##end
Sekunden = int(time.time() - start_time)
print("--finished after ",Sekunden,"seconds--")

f.write("--finished after "+str(Sekunden)+" seconds--")
f.close()
