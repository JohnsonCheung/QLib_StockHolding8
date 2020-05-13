Attribute VB_Name = "MxXlsPt"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxXlsPt."

Function FstPt(Ws As Worksheet) As PivotTable
Set FstPt = Ws.PivotTables(1)
End Function
