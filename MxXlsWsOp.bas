Attribute VB_Name = "MxXlsWsOp"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxXlsWsOp."

Sub RmvAllColAft(Ws As Worksheet, AftCol)
With WsAllColAft(Ws, AftCol)
    With .Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    .Delete Shift:=xlToLeft
End With
End Sub

Function WsAllColAft(Ws As Worksheet, AftCol) As Range
Dim C1 As Range: Set C1 = RgRC(Ws.Cells(1, AftCol), 1, 2)
Dim C2 As Range: Set C2 = Ws.Cells(1, MaxCno)
Set WsAllColAft = Ws.Range(C1, C2).EntireColumn
End Function
