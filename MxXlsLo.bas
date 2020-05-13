Attribute VB_Name = "MxXlsLo"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxXlsLo."

Function LozWs(A As Worksheet, Lon$) As ListObject
Set LozWs = FstObjByNm(A.ListObjects, Lon)
End Function

Function LozWb(A As Workbook, Lon$) As ListObject
Dim S As Worksheet: For Each S In A.Sheets
    Set LozWb = LozWs(S, Lon)
    If Not IsNothing(LozWb) Then Exit Function
Next
End Function

Function FstLo(A As Worksheet) As ListObject 'Return LoOpt
Set FstLo = FstItm(A.ListObjects)
End Function
