Attribute VB_Name = "MxXlsRgPrp"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsRgPrp."

Function IsSngRow(R As Range) As Boolean
IsSngRow = NRoZZRg(R) = 1
End Function

Function IsSngCol(R As Range) As Boolean
IsSngCol = NColzRg(R) = 1
End Function

Function SqzRg(R As Range) As Variant()
If NColzRg(R) = 1 Then
    If NRoZZRg(R) = 1 Then
        Dim O()
        ReDim O(1 To 1, 1 To 1)
        O(1, 1) = R.Value
        SqzRg = O
        Exit Function
    End If
End If
SqzRg = R.Value
End Function

Function IsA1(R As Range) As Boolean
If R.Row <> 1 Then Exit Function
If R.Column <> 1 Then Exit Function
IsA1 = True
End Function

Function WsAdrzRg$(R As Range) ' WsAdr of @R.  WsAdr always with Wsn
WsAdrzRg = "'" & WszRg(R).Name & "'!" & R.Address
End Function

Function RCzRg(R As Range) As RC
With RCzRg
.R = R.Row
.C = R.Column
End With
End Function

Function RRCCzRg(R As Range) As RRCC
With RRCCzRg
.R1 = R.Row
.R2 = .R1 + NRoZZRg(R) - 1
.C1 = R.Column
.C2 = .C1 + NColzRg(R) - 1
End With
End Function

Function DrzRg(Rg As Range, Optional R = 1) As Variant()
DrzRg = DrzSq(SqzRg(RgR(Rg, R)))
End Function

Function A1Adr$(R As Range)
A1Adr = A1zRg(R).Address(External:=True)
End Function

Function NRoZZRg&(R As Range)
NRoZZRg = R.Rows.Count
End Function

Function NColzRg&(R As Range)
NColzRg = NColzRg(R)
End Function

Function ErCellValMsg$(Ws As Worksheet, Adr$, ExpectedVal$)
If Ws.Range(Adr).Value <> ExpectedVal Then ErCellValMsg = "Cell[" & Adr & "] should be [" & ExpectedVal & "] but now[" & Ws.Range(Adr).Value & "]"
End Function

Function IsCell(R As Range) As Boolean
If NRoZZRg(R) > 1 Then Exit Function
If NColzRg(R) > 1 Then Exit Function
IsCell = True
End Function
