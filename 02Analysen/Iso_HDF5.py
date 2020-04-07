# -*- coding: cp1252 -*-
#!/usr/bin/python
#-------------------------------------------------------------------------------
# Name:        Script, um Isochronen für alle HstBer in HDF5-Datenbank zu speichern
# Purpose:
#
# Author:      mape
#
# Created:     19/08/2015
# Copyright:   (c) mape 2015
# Licence:     <your licence>
#-------------------------------------------------------------------------------


#--Module--#
import numpy
import h5py
import sqlite3

from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(r'C:'+path),'VISUM_Tools','Iso_HDF5.txt')
f = path.read_text()
f = f.split('\n')

#--Pfade--#
HDF = 'V:'+f[0]
Netz = 'V:'+f[1]
Datenbank = 'V:'+f[2]

#--Vorbereitung--#
f = h5py.File(HDF,'r+') ##HDF5-File
gr = f["test"] ##Gruppe
dats = gr["default"] ##Dataset
for i in gr.keys():
    print("folgende HDF5-Datasets gibt es: "+i)

#--SQLite--#
conn = sqlite3.connect(Datenbank)
c = conn.cursor()
c.execute("select name from 'sqlite_master'")
for i in c.fetchall():
    print("folgende SQL-Tabellen gibt es: "+i[0])


#--VISUM--#
import win32com.client.dynamic
##VISUM = win32com.client.dynamic.Dispatch("Visum.Visum.14")
##VISUM.loadversion(Netz) ##Netz,Szenrario,VISUM-Instanz
##VISUM.Filters.InitAll()


#--Ende--#
print("fertig")