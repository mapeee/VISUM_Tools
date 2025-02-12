# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        Script to calculate maximum vehicle demand for TNM
# Author:      mape
# Created:     12/02/2025
# Copyright:   (c) mape 2025
# Licence:     GNU GENERAL PUBLIC LICENSE
# #---------------------

import numpy as np
import pandas as pd

# Functions
def minutes_to_timestring(minutes):
    hours = minutes // 60  # Get hours
    mins = minutes % 60    # Get remaining minutes
    return f"{hours:02}:{mins:02}"  # Format as "HH:MM"

def cal_maxVH(_VJ, _attr):
    VJ_grouped = _VJ.groupby(_attr)
    for _attr_Name, VJ_data in VJ_grouped:
        dep_arr_matrix = (VJ_data['Dep'].values[:, None] <= mins) & (VJ_data['Arr'].values[:, None] >= mins)
        count_per_minute = dep_arr_matrix.sum(axis=0)  # Sum along the first axis to count the number of vehicles
        i_max = count_per_minute.argmax()  # Index of max vehicles
        results.append([_attr_Name, VJ_data["VSys"].iloc[0], int(count_per_minute[i_max]), str(minutes_to_timestring(mins[i_max]))])

    Table.AddMultiTableEntries(len(results))
    Table.TableEntries.SetMultipleAttributes([_attr, "VSys", "Spitzenbedarf", "Spitzenzeit"], results)

def create_UDT(_table_name):
    if not any(i.AttValue("Name") == "TNM" for i in Visum.Net.UserDefinedGroups.GetAll):
        Visum.Net.AddUserDefinedGroup("TNM", "Teilnetzmanagement")
    Tables = Visum.Net.TableDefinitions.GetMultiAttValues("Name")
    if len(list(filter(lambda x: _table_name in x, Tables))) > 0:
        _Table = Visum.Net.TableDefinitions.ItemByKey(_table_name)
        Visum.Net.RemoveTableDefinition(_Table)
    _Table = Visum.Net.AddTableDefinition(_table_name)
    _Table.SetAttValue("GROUP","TNM")
    if _table_name == "TNM Spitzenbedarf Teilnetz":
        _Table.SetAttValue("COMMENT", "Fahrzeuge zur Spitzenzeit je Teilnetz")
        _Table.TableEntries.AddUserDefinedAttribute("Teilnetz","Teilnetz","Teilnetz",5)
        _Table.TableEntries.AddUserDefinedAttribute("FZG","Fzg","Fzg",5)
        _Table.TableEntries.AddUserDefinedAttribute("FZG_Rang","Fzg_Rang","Fzg_Rang",1)
        _Table.TableEntries.AddUserDefinedAttribute("Spitzenbedarf_Rang","Spitzenbedarf_Rang","Spitzenbedarf_Rang",1)
    if _table_name == "TNM Spitzenbedarf Linie":
        _Table.SetAttValue("COMMENT","Fahrzeuge Spitzenstunde je Linie")
        _Table.TableEntries.AddUserDefinedAttribute("LINIE","Linie","Linie",5)
        _Table.TableEntries.AddUserDefinedAttribute("VSYS","VSys","VSys",5)
    if _table_name == "TNM Spitzenbedarf Fzg":
        _Table.SetAttValue("COMMENT","Fahrzeuge Spitzenstunde je Fahrzeugtyp")
        _Table.TableEntries.AddUserDefinedAttribute("FZG","Fzg","Fzg",5)
        _Table.TableEntries.AddUserDefinedAttribute("VSYS","VSys","VSys",5)
        
    _Table.TableEntries.AddUserDefinedAttribute("Spitzenbedarf","Spitzenbedarf","Spitzenbedarf",1)
    _Table.TableEntries.AddUserDefinedAttribute("Spitzenzeit","Spitzenzeit","Spitzenzeit",5)
    for i in _Table.TableEntries.Attributes.GetAll:
        if i.Category == 'User-defined attributes':
            i.UserDefinedGroup = "TNM"

    return _Table

def create_VJ_df():
    Dep = np.array(Visum.Net.VehicleJourneys.GetMultiAttValues("Dep"), dtype=int)[:, 1] // 60
    Arr = np.array(Visum.Net.VehicleJourneys.GetMultiAttValues("Arr"), dtype=int)[:, 1] // 60
    Lines = np.array(Visum.Net.VehicleJourneys.GetMultiAttValues(r"LINEROUTE\LINENAME"))[:, 1]
    Fzg = np.array(Visum.Net.VehicleJourneys.GetMultiAttValues(r"MIN:VEHJOURNEYSECTIONS\VEHCOMBIDENTIFIER"))[:, 1]
    Fzg_Rang = np.array(Visum.Net.VehicleJourneys.GetMultiAttValues(r"MIN:VEHJOURNEYSECTIONS\VEHCOMB\FZG_RANG"))[:, 1]
    TN = np.array(Visum.Net.VehicleJourneys.GetMultiAttValues(r"LINEROUTE\LINE\TEILNETZGRUPPE"))[:, 1]
    VSys = np.array(Visum.Net.VehicleJourneys.GetMultiAttValues("TSYSCODE"))[:, 1]
    _VJ = pd.DataFrame({'Dep': Dep, 'Arr': Arr, 'Linie': Lines, 'TN': TN, "VSys": VSys, 'FZG': Fzg, 'FZG_Rang': Fzg_Rang})
    return _VJ

# Prepare Output Table
mins = np.arange(0, 1600)
for table_name in ["TNM Spitzenbedarf Teilnetz", "TNM Spitzenbedarf Linie", "TNM Spitzenbedarf Fzg"]:
    Table = create_UDT(table_name)
    results = []
    VJ = create_VJ_df()
    if table_name == "TNM Spitzenbedarf Linie":
        cal_maxVH(VJ, "Linie")
    
    if table_name == "TNM Spitzenbedarf Fzg":
        cal_maxVH(VJ, "FZG")

    if table_name == "TNM Spitzenbedarf Teilnetz":
        VJ = VJ.loc[VJ['VSys'] == 'Bus']
        VJ_grouped = VJ.groupby('TN')
        for TN, VJ_TN in VJ_grouped:
            result_df = pd.DataFrame({'Minute': mins})
            for FZG_TN, sub_VJ_TN in VJ_TN.sort_values("FZG_Rang").groupby("FZG"):
                result_df[FZG_TN] = np.array([np.sum((sub_VJ_TN["Dep"].values <= x) & 
                                                     (sub_VJ_TN["Arr"].values >= x)) for x in mins], dtype=int) 
            unique_FZG = VJ_TN.sort_values("FZG_Rang", ascending = False).drop_duplicates("FZG")["FZG"].tolist()

            FZG_Rang_prev, FZG_TN_prev = 99, ''
            
            for FZG_TN in unique_FZG:
                if FZG_TN != "":
                    FZG_Rang = Visum.Net.VehicleCombinations.ItemByKey(int(FZG_TN.split(" ")[0])).AttValue("FZG_Rang")
                    FZG_Name = Visum.Net.VehicleCombinations.ItemByKey(int(FZG_TN.split(" ")[0])).AttValue("Name")
                    if FZG_Rang == FZG_Rang_prev or FZG_Rang_prev == 99:
                        result_df["new "+FZG_TN] = result_df[FZG_TN]
                        results.append([TN, FZG_Name, FZG_Rang, int(result_df[FZG_TN].max()), int(result_df[FZG_TN].max()), str(minutes_to_timestring(mins[result_df.loc[result_df[FZG_TN].idxmax(), "Minute"]]))])
                        FZG_Rang_prev = FZG_Rang
                        FZG_TN_prev = FZG_TN
                    else:
                        result_df["new " + FZG_TN] = result_df[FZG_TN] - (result_df["new " + FZG_TN_prev].max() - result_df["new " + FZG_TN_prev])
                        results.append([TN, FZG_Name, FZG_Rang, int(result_df[FZG_TN].max()), int(result_df["new " + FZG_TN].max()), str(minutes_to_timestring(mins[result_df.loc[result_df["new " + FZG_TN].idxmax(), "Minute"]]))])
                        FZG_Rang_prev = FZG_Rang
                        FZG_TN_prev = FZG_TN
                else:
                    results.append([TN, FZG_Name, 0, int(result_df[FZG_TN].max()), int(result_df[FZG_TN].max()), str(minutes_to_timestring(mins[result_df.loc[result_df[FZG_TN].idxmax(), "Minute"]]))])
                    FZG_Rang_prev = 99
                    FZG_TN_prev = ''
            del result_df

        Table.AddMultiTableEntries(len(results))
        Table.TableEntries.SetMultipleAttributes(["Teilnetz", "FZG", "FZG_Rang", "Spitzenbedarf", "Spitzenbedarf_Rang", "Spitzenzeit"], results)
        