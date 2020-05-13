Attribute VB_Name = "MxXlsLoPutDrs"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsLoPutDrs."

Private Sub PutDrsToLo__Tst()
Dim Lo As ListObject, D As Drs
GoSub Z
Exit Sub

Z:
Set Lo = NwLozDrs(SampDrs1, NwA1)
PutDrsToLo SampDrs2, Lo
Stop
'PutDrsToLo SampDrs3, Lo
Stop
ClsCWbNoSav WbzLo(Lo)
Return
End Sub

Sub PutDrsToLo(D As Drs, Lo As ListObject)
ClrLo Lo
PutDyToLo SelDrsAlwEzFny(D, FnyzLo(Lo)).Dy, Lo
End Sub

Sub PutDyToLo(Dy(), Lo As ListObject)
RgzDy Dy, RgRC(Lo.Range, 2, 1)
End Sub
