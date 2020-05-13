Attribute VB_Name = "MxXlsLoNw"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsLoNw."

Function NwLozSq(Sq(), At As Range, Optional Lon$) As ListObject
Set NwLozSq = NwLo(RgzSq(Sq(), At), Lon)
End Function

Function NwLo(Rg As Range, Optional Lon$) As ListObject
Dim S As Worksheet: Set S = WszRg(Rg)
Dim O As ListObject: Set O = S.ListObjects.Add(xlSrcRange, Rg, , xlYes)
BdrAround Rg
Rg.EntireColumn.AutoFit
SetLon O, Lon
Set NwLo = O
End Function

Function NwLozDrs(D As Drs, At As Range, Optional Lon$) As ListObject
Set NwLozDrs = NwLo(RgzDrs(D, At), Lon)
End Function

Function CrtEmpLo(At As Range, FF$, Optional Lon$) As ListObject
Set CrtEmpLo = NwLo(RgzAyH(FnyzFF(FF), At), Lon)
End Function

Function PutSq(Sq(), At As Range) As Range
PutSq = RgzSq(Sq, At)
End Function

Function RgzSq(Sq(), At As Range) As Range
If Si(Sq) = 0 Then
    Set RgzSq = A1zRg(At)
    Exit Function
End If
Dim O As Range
Set O = ResiRg(At, Sq)
O.MergeCells = False
O.Value = Sq
Set RgzSq = O
End Function

Sub NwLozDbt(D As Database, T, At As Range, Optional AddgWay As eWsdAddWay)
Const CSub$ = CMod & "NwLozDbt"
Select Case AddgWay
Case eWsdAddWay.eSqWay: NwLozSq SqzT(D, T), At
Case eWsdAddWay.eWcWay: NwLozFbt At, D.Name, T
Case Else: Thw CSub, "Invalid AddgWay"
End Select
End Sub

Sub NwLozDbt1(D As Database, T, At As Range, Optional AddgWay As eWsdAddWay)
NwLo NwLozSq(SqzT(D, T), At), Lon(T)
End Sub

Sub NwLozDbtWs(D As Database, T, Ws As Worksheet)
NwLozDbt D, T, A1zWs(Ws)
End Sub
