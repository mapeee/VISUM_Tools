#-------------------------------------------------------------------------------
# Name:        xxxx
# Author:      mape
# Created:     01/04/2025
# Copyright:   (c) mape 2025
# Licence:     GNU GENERAL PUBLIC LICENSE
# #---------------------

import numpy as np
from VisumPy.helpers import GetSkimMatrixRaw
from VisumPy.helpers import SetMulti

def ClosnessCentrality(_Visum, _JRT_filtered, _SFQ_filtered, _maskStopAreas):
    Visum.Log(20480, "Calculate Closness Centrality (SNAMUTS)")
    _SFQ_filtered = np.where(_SFQ_filtered == 0, np.nan, _SFQ_filtered) # Avoid divide-by-0
    CC = 4 * np.sqrt(_JRT_filtered / _SFQ_filtered)
    CC = np.nan_to_num(CC, nan=100)
    CC = _SetDiagonal(CC)
    CC_averages = np.mean(CC, axis=1)
    CC_iter = iter(CC_averages)
    CC_result = [next(CC_iter) if b else 0 for b in _maskStopAreas]
    SetMulti(_Visum.Net.StopAreas, "CC_SNAMUTS", CC_result, False)

def DegreeCentrality(_Visum, _NTR_filtered, _maskStopAreas):
    Visum.Log(20480, "Calculate Degree Centrality (SNAMUTS)")
    _NTR_filtered = _SetDiagonal(_NTR_filtered)
    DC_averages = np.mean(_NTR_filtered, axis=1)
    DC_iter = iter(DC_averages)
    DC_result = [next(DC_iter) if b else 0 for b in _maskStopAreas]
    SetMulti(_Visum.Net.StopAreas, "DC_SNAMUTS", DC_result, False)

def FilterMatrices(_JRT, _NTR, _SFQ, _maskStopAreas):
    _JRT_filtered = _JRT[_maskStopAreas][:, _maskStopAreas]
    _NTR_filtered = _NTR[_maskStopAreas][:, _maskStopAreas]
    _SFQ_filtered = _SFQ[_maskStopAreas][:, _maskStopAreas]
    
    return _JRT_filtered, _NTR_filtered, _SFQ_filtered

def GetMatrices(_Visum, _Hour_interval):
    _JRT = GetSkimMatrixRaw(_Visum, "JRT", "OEV_E_O")
    _JRT[_JRT > 200] = 200
    _NTR = GetSkimMatrixRaw(_Visum, "NTR", "OEV_E_O")
    _NTR[_NTR > 7] = 7
    _SFQ = GetSkimMatrixRaw(_Visum, "SFQ", "OEV_E_O")
    _SFQ = _SFQ / _Hour_interval
    
    return _JRT, _NTR, _SFQ
    
def MaskStopAreas(_Visum):
    _StopAreas = Visum.Net.StopAreas.GetMultipleAttributes(["NO", r"NODE\MIN:CONTAININGZONES\LAGE"])
    _StopAreas = np.array(_StopAreas, dtype = object)
    _StopAreas[_StopAreas == None] = 10
    _maskStopAreas = [val[1] < 4.0 for val in _StopAreas]
    
    return _maskStopAreas

def _SetDiagonal(_Matrice):
    row_means = np.mean(_Matrice, axis=1)
    np.fill_diagonal(_Matrice, row_means)
    
    return _Matrice

## Calculations ##
Visum.Log(20480, "Starting SNAMUTS Calculations...")

maskStopAreas = MaskStopAreas(Visum)
JRT, NTR, SFQ = GetMatrices(Visum, Hour_interval)
JRT_filtered, NTR_filtered, SFQ_filtered = FilterMatrices(JRT, NTR, SFQ, maskStopAreas)

ClosnessCentrality(Visum, JRT_filtered, SFQ_filtered, maskStopAreas)
DegreeCentrality(Visum, NTR_filtered, maskStopAreas)

Visum.Log(20480, "Finished")