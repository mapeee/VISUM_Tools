# -*- coding: utf-8 -*-
"""
Created on Fri Feb  2 13:33:29 2024
"""
import numpy as np
from VisumPy.AddIn import AddIn
_ = AddIn.gettext

def calc(Visum,Case):
    
    Table = Visum.Net.TableDefinitions.ItemByKey("Ohnefall-Mitfall-Vergleich")
    if Case == "Ohnefall":
        for i in Table.TableEntries.GetAll: i.SetAttValue("Ohnefall",0)
    for i in Table.TableEntries.GetAll: i.SetAttValue("Mitfall",0)

    #PuT
    if Case == "Ohnefall":
        sum_PuT = Visum.Net.Matrices.ItemByKey(10).AttValue("SUM")+Visum.Net.Matrices.ItemByKey(12).AttValue("SUM")
        dig_PuT = Visum.Net.Matrices.ItemByKey(10).AttValue("DIAGONALSUM")+Visum.Net.Matrices.ItemByKey(12).AttValue("DIAGONALSUM")
    else:
        sum_PuT = Visum.Net.Matrices.ItemByKey(11).AttValue("SUM")+Visum.Net.Matrices.ItemByKey(12).AttValue("SUM")
        dig_PuT = Visum.Net.Matrices.ItemByKey(11).AttValue("DIAGONALSUM")+Visum.Net.Matrices.ItemByKey(12).AttValue("DIAGONALSUM")
    
    if Case == "Mitfall" and Visum.Net.Matrices.ItemByKey(11).AttValue("SUM") == 0 and Visum.Net.Matrices.ItemByKey(2).AttValue("SUM") == 0:
        Visum.Log(12288,_("PT trips planning case missing!"))
        return False
    assign_PuT = sum_PuT-dig_PuT
    Table.TableEntries.ItemByKey(1).SetAttValue(Case,sum_PuT)
    Table.TableEntries.ItemByKey(2).SetAttValue(Case,dig_PuT)
    Table.TableEntries.ItemByKey(3).SetAttValue(Case,np.array(Visum.Net.TSystems.GetMultiAttValues('PASSKMTRAV(AP)')).astype("float")[:,1].sum())
    Table.TableEntries.ItemByKey(4).SetAttValue(Case,np.array(Visum.Net.TSystems.GetMultiAttValues('PASSHOURTRAV(AP)')).astype("float")[:,1].sum()/60/60)
    Table.TableEntries.ItemByKey(5).SetAttValue(Case,Table.TableEntries.ItemByKey(3).AttValue(Case)/assign_PuT)
    Table.TableEntries.ItemByKey(6).SetAttValue(Case,Table.TableEntries.ItemByKey(4).AttValue(Case)*60/assign_PuT)
    Table.TableEntries.ItemByKey(7).SetAttValue(Case,int(assign_PuT-np.nan_to_num(np.array(Visum.Net.Zones.GetMultiAttValues(r"SUM:ORIGCONNECTORS\VOLPERSPUT(AP)")).astype("float")[:,1],False,0).sum()))
    
    #PrT
    if Case == "Ohnefall":
        sum_PrT = Visum.Net.Matrices.ItemByKey(1).AttValue("SUM")
        dig_PrT = Visum.Net.Matrices.ItemByKey(1).AttValue("DIAGONALSUM")
    else:
        sum_PrT = Visum.Net.Matrices.ItemByKey(2).AttValue("SUM")
        dig_PrT = Visum.Net.Matrices.ItemByKey(2).AttValue("DIAGONALSUM")
        
    assign_PrT = sum_PrT-dig_PrT
    if assign_PrT == 0:
        Visum.Log(12288,_("Car trips %s missing!") %(Case))
        return False
    
    d_T = np.array(Visum.Net.Links.GetMultiAttValues('VEHKMTRAVPRT_DSEG(LKW_GV,AP)')).astype("float")[:,1].sum()
    t_T = np.array(Visum.Net.Links.GetMultiAttValues('VEHHOURTRAVTCUR_DSEG(LKW_GV,AP)')).astype("float")[:,1].sum()
    Table.TableEntries.ItemByKey(8).SetAttValue(Case,sum_PrT*1.3)
    Table.TableEntries.ItemByKey(9).SetAttValue(Case,dig_PrT*1.3)
    Table.TableEntries.ItemByKey(10).SetAttValue(Case,(np.array(Visum.Net.Links.GetMultiAttValues('VEHKMTRAVPRT(AP)')).astype("float")[:,1].sum()-d_T)*1.3)
    Table.TableEntries.ItemByKey(11).SetAttValue(Case,((np.array(Visum.Net.Links.GetMultiAttValues('VEHHOURTRAVTCUR(AP)')).astype("float")[:,1].sum()-t_T)*1.3)/60/60)
    Table.TableEntries.ItemByKey(12).SetAttValue(Case,Table.TableEntries.ItemByKey(10).AttValue(Case)/(assign_PrT*1.3))
    Table.TableEntries.ItemByKey(13).SetAttValue(Case,Table.TableEntries.ItemByKey(11).AttValue(Case)*60/(assign_PrT*1.3))
    Table.TableEntries.ItemByKey(14).SetAttValue(Case,int(assign_PrT-np.nan_to_num(np.array(Visum.Net.Zones.GetMultiAttValues(r"SUM:ORIGCONNECTORS\VOLVEH_TSYS(C,AP)")).astype("float")[:,1],False,0).sum()))
    
    Table.TableEntries.ItemByKey(15).SetAttValue(Case,sum_PrT)
    Table.TableEntries.ItemByKey(16).SetAttValue(Case,np.array(Visum.Net.Links.GetMultiAttValues('VEHKMTRAVPRT(AP)')).astype("float")[:,1].sum()-d_T)
    Table.TableEntries.ItemByKey(17).SetAttValue(Case,(np.array(Visum.Net.Links.GetMultiAttValues('VEHHOURTRAVTCUR(AP)')).astype("float")[:,1].sum()-t_T)/60/60)
    Table.TableEntries.ItemByKey(18).SetAttValue(Case,Visum.Net.Matrices.ItemByKey(3).AttValue("SUM"))
    if Visum.Net.Matrices.ItemByKey(3).AttValue("SUM") == 0:
        return True
    Table.TableEntries.ItemByKey(19).SetAttValue(Case,d_T)
    Table.TableEntries.ItemByKey(20).SetAttValue(Case,t_T/60/60)
    Table.TableEntries.ItemByKey(21).SetAttValue(Case,Table.TableEntries.ItemByKey(19).AttValue(Case)/Table.TableEntries.ItemByKey(18).AttValue(Case))
    Table.TableEntries.ItemByKey(22).SetAttValue(Case,Table.TableEntries.ItemByKey(20).AttValue(Case)*60/Table.TableEntries.ItemByKey(18).AttValue(Case))
    return True