Attribute VB_Name = "MxVbStrDft"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Vb"
Const CMod$ = CLib & "MxVbStrDft."
Function Dft(V, DftV)
If IsEmp(V) Then
   Dft = DftV
Else
   Dft = V
End If
End Function

Function DftStr$(Str, Dft)
DftStr = IIf(Str = "", Dft, Str)
End Function

Function Limit(V, A, B)
Select Case V
Case V > B: Limit = B
Case V < A: Limit = A
Case Else: Limit = V
End Select
End Function
