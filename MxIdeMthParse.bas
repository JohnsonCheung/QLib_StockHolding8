Attribute VB_Name = "MxIdeMthParse"
Option Explicit
Option Compare Text
Const CNs$ = "Mth.Ln"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthParse."

Function TakMthKd$(Ln)
TakMthKd = PfxzAySpc(Ln, MthKdAy)
End Function

Function TakMthTy$(Ln)
TakMthTy = PfxzAySpc(Ln, MthTyAy)
End Function

Function DltMthTy$(Ln)
DltMthTy = RmvPfxSySpc(Ln, MthTyAy)
End Function

Function ShfTyChr$(OLin$)
ShfTyChr = ShfChr(OLin, TyChrLis)
End Function

Function TakTyChr$(S)
TakTyChr = TakChr(S, TyChrLis)
End Function

Function ShfRetTyzAftPm$(OAftPm$)
Dim A$: A = ShfTermAftAs(OAftPm)
If LasChr(A) = ":" Then
    ShfRetTyzAftPm = RmvLasChr(A)
    OAftPm = ":" & OAftPm
Else
    ShfRetTyzAftPm = A
End If
End Function

Function IsPubMdy(Mdy) As Boolean
Select Case Mdy
Case "Public", "": IsPubMdy = True
End Select
End Function
