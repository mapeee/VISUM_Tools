# -*- coding: cp1252 -*-
#!/usr/bin/python

#-------------------------------------------------------------------------------
# Name:        Script to import AFZ data to PTV Visum
# Author:      mape
# Created:     30/01/2025
# Copyright:   (c) mape 2025
# Licence:     GNU GENERAL PUBLIC LICENSE
# #-------------------------------------------------------------------------------

# Mo-Fr; Wochenende
# add journey found to csv

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
    print("> Linien:", _lines)
    print()
    
    first_date = _AFZ_data['Date'].min()
    last_date = _AFZ_data['Date'].max()
    print("> Startdatum:", first_date.strftime('%d.%m.%Y'))
    print("> Endatum:", last_date.strftime('%d.%m.%Y'))
    print("-" * 30)
    print()
    
def extrapolate_AFZ(_AFZ_extra):
    _AFZ_extra["HstIDXprev"] = _AFZ_extra["HstIDX"].shift(1) # to compare IDX later
    _AFZ_extra["HstIDXprev"] = _AFZ_extra["HstIDXprev"].fillna(0) # for first row
    _AFZ_extra['HstIDXreset'] = _AFZ_extra["HstIDXprev"] + 1 != _AFZ_extra["HstIDX"]
    _AFZ_extra['HstIDXresetGroup'] = _AFZ_extra['HstIDXreset'].cumsum()
    _AFZ_extra['Belastung'] = _AFZ_extra.groupby('HstIDXresetGroup')['EinAusSaldo'].cumsum().groupby(_AFZ_extra['HstIDXresetGroup']).apply(lambda x: np.maximum(x, 0)).reset_index(drop=True)
    return _AFZ_extra
    
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
    
def modify_AFZ(_AFZmod):
    _AFZmod['Abzeit_seconds'] = _AFZmod['Abzeit'].apply(time_to_seconds)
    _AFZmod["LinieNo"] = _AFZmod["Linie"].astype(str).str[2:6]
    _AFZmod['Date'] = pd.to_datetime(_AFZmod['Datum'], format='%d.%m.%Y')
    _AFZmod["EinAusSaldo"] = _AFZmod["Ein"] - _AFZmod["Aus"]
    AFZ_info(_AFZmod)
    return _AFZmod

def time_to_seconds(time_str):
    hours, minutes = map(int, time_str.split(':'))
    return hours * 3600 + minutes * 60


# open and modify AFZ data
AFZ_data = pd.read_csv(data["AFZ_data"], delimiter=";")
AFZ_data = modify_AFZ(AFZ_data)

# extrapolate data
AFZ_data = extrapolate_AFZ(AFZ_data)
year = AFZ_data["Date"].iloc[0].year

# load Visum
Visum = win32com.client.dynamic.Dispatch("Visum.Visum.25")
Visum.IO.loadversion(data["Network"])
Visum.Graphic.StopDrawing = True
Visum.Filters.InitAll()
filter_lines(Visum, AFZ_data)

# iterate over all journey days in the given AFZ dataset
journey_day = AFZ_data.groupby("Datum")
for journey_day_name, journey_day_data in journey_day:
    # get attributes of journeyItems from the model to merge with AFZ data
    Date_Journeys_V = pd.DataFrame(Visum.Net.VehicleJourneyItems.GetMultipleAttributes([r"VEHJOURNEY\FROMSTOPPOINT\ISA_NO", r"VEHJOURNEY\TOSTOPPOINT\ISA_NO", r"VEHJOURNEY\DEP", 
                                                                                        r"VEHJOURNEY\LINENAME", r"TIMEPROFILEITEM\LINEROUTEITEM\STOPPOINT\ISA_NO"], True),
                                   columns = ["AbHst", "AnHst", "Abzeit_seconds", "LinieNo", "Hst"])
    Date_Journeys_V = Date_Journeys_V.astype({'AbHst':int, 'AnHst':int, 'Abzeit_seconds':int, 'LinieNo':str, 'Hst':int})
    merged_df = pd.merge(Date_Journeys_V, journey_day_data, on = ['AbHst', 'AnHst', 'Abzeit_seconds', 'LinieNo', 'Hst'], how = 'left')
    merged_df = merged_df.fillna("")
    
    Values = merged_df[['FahrtNr', 'FzgNr', "Ein", "Aus", "Belastung"]].values.tolist()
    Attributes = [rf'FAHRTNR_AFZ_{year}({journey_day_name})', rf'FZGNR_AFZ_{year}({journey_day_name})',
                  rf'EINSTEIGER_AFZ_{year}({journey_day_name})', rf'Aussteiger_AFZ_{year}({journey_day_name})', rf'BELASTUNG_AB_AFZ_{year}({journey_day_name})']
    Visum.Net.VehicleJourneyItems.SetMultipleAttributes(Attributes, Values, True)

# finalizing
Visum.Graphic.StopDrawing = False