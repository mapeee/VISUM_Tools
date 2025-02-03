# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        Script to import AFZ data to PTV Visum
# Author:      mape
# Created:     30/01/2025
# Copyright:   (c) mape 2025
# Licence:     GNU GENERAL PUBLIC LICENSE
# #-------------------------------------------------------------------------------

import json
import numpy as np
import os
import pandas as pd
from pathlib import Path
import sys
import win32com.client.dynamic

directory = Path(os.path.dirname(os.path.abspath(sys.argv[0])))
with open(os.path.join(directory ,"config.json"), "r") as json_file:
    data = json.load(json_file)

def AFZ_info(_AFZ_data):
    # translate time stringe into seconds a day 
    _lines = ', '.join(_AFZ_data['LinieNo'].unique())
    print("> Lines:", _lines)
    print()
    
    first_date = _AFZ_data['Datum'].min()
    last_date = _AFZ_data['Datum'].max()
    print("> Startdatum:", first_date.strftime('%d.%m.%Y'))
    print("> Endatum:", last_date.strftime('%d.%m.%Y'))
    print("-" * 30)
    print()
    
def filter_lines(_Visum, _AFZ):
    # Set Value 1 in AddVal3 if Line in AFZ
    def _checkLine(l, _uniqueLines):
        return 1 if l in _uniqueLines else 0
    
    _Lines = _Visum.Net.Lines.GetMultiAttValues("Name", False)
    _uniqueLines = _AFZ['LinieNo'].unique()
    replacetuple = tuple((i[0], _checkLine(i[1], _uniqueLines)) for i in _Lines)
    _Visum.Net.Lines.SetMultiAttValues("AddVal3", replacetuple, False)

    # Filter
    Lines = _Visum.Filters.LineGroupFilter()
    Line = Lines.LineFilter()
    Line.AddCondition("OP_NONE",False,"AddVal3","EqualVal",1)
    Lines.UseFilterForLines = True
    Lines.UseFilterForLineRoutes = True
    Lines.UseFilterForLineRouteItems = True
    Lines.UseFilterForTimeProfiles = True
    Lines.UseFilterForTimeProfileItems = True
    Lines.UseFilterForVehJourneys = True
    Lines.UseFilterForVehJourneySections = True
    Lines.UseFilterForVehJourneyItems = True
    
    InsertedLinks = _Visum.Filters.LinkFilter()
    InsertedLinks.AddCondition("OP_NONE",False,"ONACTIVELINEROUTE", "EqualVal", 1)

def time_to_seconds(time_str):
    hours, minutes = map(int, time_str.split(':'))
    return hours * 3600 + minutes * 60

# load Visum
Visum = win32com.client.dynamic.Dispatch("Visum.Visum.25")
Visum.IO.loadversion(data["Network"])
Visum.Filters.InitAll()

# open AFZ data
AFZ_data = pd.read_csv(data["AFZ_data"], delimiter=";")
AFZ_data['Abzeit_seconds'] = AFZ_data['Abzeit'].apply(time_to_seconds)
AFZ_data["LinieNo"] = AFZ_data["Linie"].astype(str).str[2:6]
AFZ_data['Datum'] = pd.to_datetime(AFZ_data['Datum'], format='%d.%m.%Y')

AFZ_data["HstIDXprev"] = AFZ_data["HstIDX"].shift(1) # to compare IDX later
AFZ_data["HstIDXprev"] = AFZ_data["HstIDXprev"].fillna(0) # for first row
AFZ_data['HstIDXreset'] = AFZ_data["HstIDXprev"] + 1 != AFZ_data["HstIDX"]
AFZ_data['HstIDXresetGroup'] = AFZ_data['HstIDXreset'].cumsum()
AFZ_data["EinAusSaldo"] = AFZ_data["Ein"] - AFZ_data["Aus"]
AFZ_data['Belastung'] = AFZ_data.groupby('HstIDXresetGroup')['EinAusSaldo'].cumsum().groupby(AFZ_data['HstIDXresetGroup']).apply(lambda x: np.maximum(x, 0)).reset_index(drop=True)

year = AFZ_data["Datum"].iloc[0].year
AFZ_info(AFZ_data)

# iterate over all lines
lines = AFZ_data.groupby("Linie")
for line_name, line_data in lines:
    print(f"> Line: {str(line_name)[2:6]}")
    Line_Journeys_V = Visum.Net.VehicleJourneyItems.GetFilteredSet(rf'[VEHJOURNEY\LINENAME]="{str(line_name)[2:6]}"')
    
    # iterate over all journey days in the given AFZ dataset
    journey_day = line_data.groupby("Datum")
    for journey_day_name, journey_day_data in journey_day:
        # get attributes of journeyItems from the model to merge with AFZ data
        Date_Journeys_V = pd.DataFrame(Line_Journeys_V.GetMultipleAttributes([r"VEHJOURNEY\FROMSTOPPOINT\ISA_NO",
                                                                              r"VEHJOURNEY\TOSTOPPOINT\ISA_NO",r"VEHJOURNEY\DEP",r"TIMEPROFILEITEM\LINEROUTEITEM\STOPPOINT\ISA_NO"]),
                                       columns = ["AbHst","AnHst","Abzeit_seconds","Hst"])
        Date_Journeys_V = Date_Journeys_V.astype(int)
        merged_df = pd.merge(Date_Journeys_V, journey_day_data, on = ['AbHst', 'AnHst', 'Abzeit_seconds', 'Hst'], how = 'left')
        merged_df = merged_df.fillna("")
        FahrtNr = [(i + 1, value) for i, value in enumerate(merged_df["FahrtNr"])]
        FzgNr = [(i + 1, value) for i, value in enumerate(merged_df["FzgNr"])]
        Ein = [(i + 1, value) for i, value in enumerate(merged_df["Ein"])]
        Aus = [(i + 1, value) for i, value in enumerate(merged_df["Aus"])]
        Bel = [(i + 1, value) for i, value in enumerate(merged_df["Belastung"])]
        
        Line_Journeys_V.SetMultiAttValues(rf'FAHRTNR_AFZ_{year}({journey_day_name})', FahrtNr)
        Line_Journeys_V.SetMultiAttValues(rf'FZGNR_AFZ_{year}({journey_day_name})', FzgNr)
        Line_Journeys_V.SetMultiAttValues(rf'EINSTEIGER_AFZ_{year}({journey_day_name})', Ein)
        Line_Journeys_V.SetMultiAttValues(rf'Aussteiger_AFZ_{year}({journey_day_name})', Aus)
        Line_Journeys_V.SetMultiAttValues(rf'BELASTUNG_AB_AFZ_{year}({journey_day_name})', Bel)
    

# finalizing
filter_lines(Visum, AFZ_data)
