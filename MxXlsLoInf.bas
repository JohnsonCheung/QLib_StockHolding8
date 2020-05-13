Attribute VB_Name = "MxXlsLoInf"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsLoInf."
Public Const LoInfFF$ = "Wsn Lon R C NR NC"

Function LoInfDy(Wb As Workbook) As Variant()
Dim O()
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    PushI O, LoInfDyzWs(Ws)
Next
LoInfDy = O
End Function

Function LoInfDyzWs(Ws As Worksheet) As Variant()
Dim Lo As ListObject: For Each Lo In Ws.ListObjects
    PushI LoInfDyzWs, LoInfDr(Lo)
Next
End Function

Private Sub LoInfDr__Tst()
Dim Lo As ListObject: Set Lo = SampLo
D LoInfDr(Lo)
ClsCWbNoSav WbzLo(Lo)
End Sub

Function LoInfDr(L As ListObject) As Variant()
Dim Wsn: Wsn = WsnzLo(L)
Dim Lon$:: Lon = L.Name
Dim NR&: NR = NRowOfLo(L)
Dim NC&: NC = L.ListColumns.Count
LoInfDr = Array(WsnzLo(L), L.Name, L.Range.Row, L.Range.Column, NR, NC)
End Function

Private Sub LoInfDy__Tst()
Dim Wb As Workbook: Set Wb = NwWb
AddWszSq Wb, SampSq
AddWszSq Wb, SampSq1
BrwSq LoInfDy(Wb)
End Sub

Function LoInfDrs(Wb As Workbook) As Drs
LoInfDrs = DrszFF(LoInfFF, LoInfDy(Wb))
End Function
