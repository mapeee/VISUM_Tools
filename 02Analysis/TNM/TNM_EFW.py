#!/usr/bin/env python3
"""
Kopiere die Entfernungen aus dem EFW in die LinienRouten-Verläufe des aktiven TN

Erstellt: 25.12.2025
@author: mape
"""
import pandas as pd
from VisumPy.helpers import SetMulti


def checks(Visum):
    if not Visum.Net.LineRouteItems.AttrExists("EFW_METER"):
        Visum.Log(12288, "LinienRoutenElemente: BDA 'EFW_METER' fehlt")
    if Visum.Net.Lines.CountActive == 0:
        Visum.Log(12288, "Keine aktiven Linien vorhanden")

checks(Visum)
TN = Visum.Net.AttValue("TN")
# EFW Table
EFW = Visum.Net.TableDefinitions.ItemByKey(f"{TN}_EFW")
EFW = pd.DataFrame(EFW.TableEntries.GetMultipleAttributes(["LINIE", "VONHPNAME", "NACHHPNAME", "METER"], True),
                           columns = ["LINIE", "VONHAPNAME", "NACHHPNAME", "METER"])
EFW = EFW.drop_duplicates(subset=["LINIE", "VONHAPNAME", "NACHHPNAME"])
# LineRouteItems
LineRouteItems = pd.DataFrame(Visum.Net.LineRouteItems.GetMultipleAttributes(["LINENAME", r"STOPPOINT\NAME", r"NEXTROUTEPOINT\STOPPOINT\NAME"], True),
                           columns = ["LINIE", "VONHAPNAME", "NACHHPNAME"])
# Merge
EFW_result = LineRouteItems.merge(EFW, on=["LINIE", "VONHAPNAME", "NACHHPNAME"],how="left")
EFW_result["METER"] = EFW_result["METER"].fillna(0).astype(int)
# Copy into PTV Visum
SetMulti(Visum.Net.LineRouteItems, "EFW_METER", EFW_result["METER"].tolist(), True)
Visum.Log(20480, f"EFW auf LinienRouten-Verläufe übertragen: {TN}_EFW")
