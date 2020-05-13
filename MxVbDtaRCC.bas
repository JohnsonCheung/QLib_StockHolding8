Attribute VB_Name = "MxVbDtaRCC"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbDtaRCC."

Type RCC
    R As Long
    C1 As Integer
    C2 As Integer
End Type


Function NxtRCC(A As RCC) As RCC
If A.C1 = A.C2 Then Exit Function 'No Nxt
NxtRCC = RCC(A.R, A.C1 + 1, A.C2)
End Function
Function RCC(R&, C1%, C2%) As RCC
With RCC
    .R = R
    .C1 = C1
    .C2 = C2
End With
End Function
