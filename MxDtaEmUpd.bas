Attribute VB_Name = "MxDtaEmUpd"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Dta"
Const CMod$ = CLib & "MxDtaEmUpd."
Enum eUpdRpt
    eRptOnly
    eUpdAndRpt
    eUpdOnly
End Enum

Enum eAnyHdr
    eHdrYes
    eHdrNo
End Enum

Function eUpdRptStr$(Upd As eUpdRpt)
Dim O$
Select Case True
Case Upd = eRptOnly: O = "*RptOnly"
Case Upd = eUpdAndRpt: O = "*UpdAndRpt"
Case Upd = eUpdOnly: O = "*UpdOnly"
Case Else: O = "eUpdRptEr(" & Upd & ")"
End Select
eUpdRptStr = O
End Function

Function IsRpt(Upd As eUpdRpt) As Boolean
Select Case Upd
Case eRptOnly, eUpdAndRpt: IsRpt = True
End Select
End Function

Function IsUpd(Upd As eUpdRpt) As Boolean
Select Case True
Case Upd = eUpdAndRpt, Upd = eUpdOnly: IsUpd = True
End Select
End Function
