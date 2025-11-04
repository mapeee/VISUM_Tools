# -*- coding: utf-8 -*-
"""
Created on Tue Nov  4 11:40:39 2025

@author: peter
"""

import pandas as pd
from VisumPy.helpers import SetMulti

Teilnetze = ["PI 1-3"]


LRETable = pd.DataFrame(Visum.Net.LineRouteItems.GetMultipleAttributes(["STOPPOINTNO", r"NEXTROUTEPOINT\STOPPOINTNO", "ISROUTEPOINT", "POSTLENGTH", "EFW"],
                                                                       False), columns = ["FROMNO", "TONO", "RP", "METER", "EFW"])

for TN in Teilnetze:
    EFWTable = pd.DataFrame(Visum.Net.TableDefinitions.ItemByKey(f"{TN} EFW").TableEntries
                                     .GetMultipleAttributes(["VONNO", "NACHNO", "METER"], True), columns = ["FROMNO", "TONO", "METER"])
    EFWTable['METER'] = EFWTable['METER'] / 1000
    EFWTable = pd.merge(LRETable, EFWTable, how='left', on=["FROMNO", "TONO"])
    EFWTable.loc[EFWTable['METER_y'].notna(), 'EFW'] = 1
    EFWTable['METER_y'] = EFWTable['METER_y'].fillna(EFWTable['METER_x'])
    
    SetMulti(Visum.Net.LineRouteItems, "POSTLENGTH", EFWTable['METER_y'].tolist(), False)
    SetMulti(Visum.Net.LineRouteItems, "EFW", EFWTable['EFW'].tolist(), False)