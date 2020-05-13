Attribute VB_Name = "MxXlsRfh"
Option Explicit
Option Compare Text
Const CNs$ = "Xls.Op"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsRfh."
Sub RfhPc(A As PivotCache)
A.MissingItemsLimit = xlMissingItemsNone
A.Refresh
End Sub

Sub RfhWs(A As Worksheet)
Dim Q As QueryTable: For Each Q In A.QueryTables: Q.BackgroundQuery = False: Q.Refresh: Next
Dim P As PivotTable: For Each P In A.PivotTables: P.Update: Next
Dim L As ListObject: For Each L In A.ListObjects: L.Refresh: Next
End Sub

Sub RfhWbPc(W As Workbook)
Dim P As PivotCache: For Each P In W.PivotCaches
    P.MissingItemsLimit = xlMissingItemsNone
    P.Refresh
Next
End Sub
Sub RfhWbWc(W As Workbook, Fb)
Dim C As WorkbookConnection: For Each C In W.Connections
    RfhWc C, Fb
Next
End Sub
Sub RfhWbWs(W As Workbook)
Dim Ws As Worksheet: For Each Ws In W.Sheets
    RfhWs Ws
Next
End Sub
Function RfhWb(Wb As Workbook, Fb) As Workbook
Wb.Application.DisplayAlerts = False
RplLozFb Wb, Fb
RfhWbWc Wb, Fb
RfhWbPc Wb
RfhWbWs Wb
StdFmtWbLo Wb
ClsWczWb Wb
DltWc Wb
Set RfhWb = Wb
Wb.Application.DisplayAlerts = True
End Function

Sub RfhWc(A As WorkbookConnection, Fb)
If IsNothing(A.OLEDBConnection) Then Exit Sub
SetWczFb A, Fb
A.OLEDBConnection.BackgroundQuery = False
A.OLEDBConnection.Refresh
End Sub

Function RfhFx(Fx$, Tp$, Fb) As Workbook
CpyFfn Tp, Fx
Dim O As Workbook: Set O = NwXlsMinv.Workbooks.Open(Fx)
RfhWb O, Fb
O.Save
Set RfhFx = O
End Function
