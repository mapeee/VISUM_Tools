#!/usr/bin/env python3
"""
Übertrage die Entfernungen aus dem EFW in die Linienrouten-Verläufe des aktiven TN

Erstellt: 25.12.2025
@author: mape
Version: 0.92
"""
import numpy as np
import pandas as pd
from VisumPy.helpers import SetMulti
from utils.TNM_Checks import check_bda


def main(Visum):
    '''Übertrage Entfernungen (Meter) aus EFW auf die Linien des betroffenen TN
    - Beginne mit Checks (BDA und BDT vorhanden, LinienRouten aktiv?)
    - Übertrage Entfernungen auf aktive LinienRouten-Verläufe gemäß identischen HP-Namen und Linienname
    - Übertrage anschließend Entfernungen auch dann, wenn nur HP-Namen identisch
    '''
    if not bool(Visum.Net.AttValue("EFW")): # Nutze nur EFW, wenn in Parametern vorgesehen.
        return True
    if not _checks(Visum):
        return False
    TN = Visum.Net.AttValue("TN")
    # EFW Table
    EFW = Visum.Net.TableDefinitions.ItemByKey(f"{TN} EFW")
    EFW = pd.DataFrame(EFW.TableEntries.GetMultipleAttributes(["LINIE", "VONHPNAME", "NACHHPNAME", "METER"], True),
                               columns = ["LINIE", "VONHAPNAME", "NACHHPNAME", "METER"])
    EFW = EFW.drop_duplicates(subset=["LINIE", "VONHAPNAME", "NACHHPNAME"]) # Drop duplicates, damit anschließend keine Doppellungen beim Merge
    # LineRouteItems
    LineRouteItems = pd.DataFrame(Visum.Net.LineRouteItems.GetMultipleAttributes(["LINENAME", r"STOPPOINT\NAME",
                                                                                  r"NEXTROUTEPOINT\STOPPOINT\NAME", "ISROUTEPOINT"], True),
                               columns = ["LINIE", "VONHAPNAME", "NACHHPNAME", "ISROUTEPOINT"])
    LineRouteItems.loc[LineRouteItems["ISROUTEPOINT"] == 0, "VONHAPNAME"] = np.nan #  Übertrage EFW nur auf Routenpunkte
    # Merge
    EFW_result = LineRouteItems.merge(EFW, on=["LINIE", "VONHAPNAME", "NACHHPNAME"],how="left")
    EFW = EFW.drop_duplicates(["VONHAPNAME", "NACHHPNAME"],keep="first") # Drop duplicates, damit anschließend keine Doppellungen beim Merge
    EFW_result = EFW_result.merge(EFW, on=["VONHAPNAME", "NACHHPNAME"],how="left") # zweiter Merge nur über identische HP Namen
    EFW_result[["METER_x", "METER_y"]] = EFW_result[["METER_x", "METER_y"]].fillna(0).astype(int)
    EFW_result.loc[EFW_result['METER_x'] == 0, 'METER_x'] = EFW_result['METER_y'] # Nehme nur Wert aus zweiten Merge (HP-Namen) wenn erster keinen Wert lieferte
    # Copy into PTV Visum
    SetMulti(Visum.Net.LineRouteItems, "EFW_METER", EFW_result["METER_x"].tolist(), True)
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
