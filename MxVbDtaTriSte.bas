Attribute VB_Name = "MxVbDtaTriSte"
Option Explicit
Option Compare Text
Const CNs$ = "Vb.Dta"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbDtaTriSte."
Enum eTri: eTriOpn: eTriYes: eTriNo: End Enum
Enum eSe01: eSeAll: eSe0: eSe1: End Enum
Function BoolzTri(A As eTri) As Boolean
Select Case True
Case A = eTriYes: BoolzTri = True
Case A = eTriNo:  BoolzTri = False
Case Else: Stop
End Select
End Function

Function Se01(Se0 As Boolean, Se1 As Boolean) As eSe01
Select Case True
Case Se0 And Not Se1: Se01 = eSe0
Case Not Se0 And Se1: Se01 = eSe1
End Select
End Function

Function HitSe01(B As Boolean, Se01 As eSe01) As Boolean
Select Case True
Case Se01 = eSeAll: HitSe01 = True
Case Se01 = eSe1: HitSe01 = B = True
Case Se01 = eSe0: HitSe01 = B = False
End Select
End Function
