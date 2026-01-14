#!/usr/bin/env python3
"""
Übertrage die Entfernungen aus dem EFW in die Linienrouten-Verläufe des aktiven TN

Erstellt: 25.12.2025
@author: mape
Version: 0.9
"""
import pandas as pd
from VisumPy.helpers import SetMulti
from utils.TNM_Checks import check_bda


def main(Visum):
    '''Übertrage Entfernungen (Meter) aus EFW auf die Linien des betroffenen TN
    - Beginne mit Checks (BDA und BDT vorhanden, LinienRouten aktiv?)
    - Übertrage Entfernungen auf aktive LinienRouten-Verläufe
    '''
    if not _checks(Visum):
        return False
    TN = Visum.Net.AttValue("TN")
    # EFW Table
    EFW = Visum.Net.TableDefinitions.ItemByKey(f"{TN} EFW")
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
    return True

def _checks(Visum):
    '''Führe vor der Übertragung der Entfernungen verschiedene Prüfroutinen durch:
        - Netzattribut BDA 'TN' vorhanden?
        - LinienRouten-Verläufe BDA 'EFW_METER' vorhanden?
        - BDT mit Entfernungswerten vorhanden?
        - Aktive LinienRouten-Verläufe vorhanden?
    '''
    if not check_bda(Visum, Visum.Net, "Network", "TN"):
        return False
    TN = Visum.Net.AttValue("TN")
    BDT = Visum.Net.TableDefinitions.GetMultiAttValues("NAME")
    if not any(f"{TN} EFW" in item for item in BDT):
        Visum.Log(12288, f"BDT existiert nicht: {TN} EFW")
        return False
    if Visum.Net.LineRouteItems.CountActive == 0:
        Visum.Log(12288, "Keine aktiven LinienRouten-Verläufe vorhanden")
        return False
    if not check_bda(Visum, Visum.Net.LineRouteItems, "LinienRouten-Verläufe", "EFW_METER"):
        return False
    return True


if __name__ == "__main__":           
    main(Visum)
