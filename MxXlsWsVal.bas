Attribute VB_Name = "MxXlsWsVal"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsWsVal."
Const CNs$ = "Xls.Names"

Private Sub SetWsv__Tst()
Dim W As Worksheet: Set W = NwWs
Dim Wsvn$, V$
GoSub YY
GoTo X
Stop
YY:
    Wsvn = "AAA"
    V = String(256, "A")
    Ept = V
    GoTo Tst
Tst:
    SetWsv W, Wsvn, V
    Act = Wsv(W, Wsvn)
    C
    Return
X:
    ClsCWsNoSav W
End Sub
Function Wsv$(Ws As Worksheet, Wsvn$)
If HasWsv(Ws, Wsvn) Then Wsv = RmvLasChr(Mid(Ws.Names(Wsvn).Value, 3))
End Function

Sub SetWsv(Ws As Worksheet, Wsvn$, V$)
If NoWsv(Ws, Wsvn) Then
    Ws.Names.Add Wsvn, Empty
End If
Ws.Names(Wsvn).RefersToR1C1 = V
End Sub

Function HasWsv(Ws As Worksheet, Wsvn) As Boolean
On Error GoTo X
HasWsv = Aft(Ws.Names(Wsvn).Name, "!") = Wsvn
Exit Function
X:
End Function

Function NoWsv(Ws As Worksheet, Wsvn) As Boolean
NoWsv = Not HasWsv(Ws, Wsvn)
End Function
