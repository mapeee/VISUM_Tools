# -*- coding: cp1252 -*-
#!/usr/bin/python

import win32com.client.dynamic
Visum = win32com.client.dynamic.Dispatch("Visum.Visum.24")
import ast
import os
import time

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
with open(path, mode='r') as f: path = f.readlines()
path = path[0] +"\\"+'VISUM_Tools'+"\\"+'Calculate.txt'
with open(path, mode='r') as path: f = path.read().splitlines()

#--Parameter--#
Network = True
Shutdown = False


Networks = ast.literal_eval(f[0])


Scenario = f[1]

#--Network--#
if Network == True:
    for Net in Networks:
        try:
            Visum.IO.loadversion(Net)
        except:
            print("> Network: "+Net+": error")
        try: 
            Visum.Procedures.Execute()
            Visum.IO.SaveVersion(Net)
        except:
            print("> Network: "+Net+": error")
            Visum.IO.SaveVersion(Net)

#--Scenario--#
if Network == False:
    sce = Visum.ScenarioManagement.OpenProject(Scenario)
    for s in sce.Scenarios.GetAll:
        print(s.AttValue("CODE"))
        
        if s.AttValue("Active") == 1:
            s.StartCalculation()
            time.sleep(300)
            while s.AttValue("CalculationState") == "CALCULATING":
                time.sleep(60)

if Shutdown == True: os.system('shutdown /s /t 1')