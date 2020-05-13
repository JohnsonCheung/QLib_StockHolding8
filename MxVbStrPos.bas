Attribute VB_Name = "MxVbStrPos"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbStrPos."


Function DifPos&(A$, B$) ' Position that A & B has first dif char
Dim LA&, LB&
LA = Len(A)
LB = Len(B)
Dim O&: For O = 1 To Min(LA, LB)
    If Mid(A, O, 1) <> Mid(B, O, 1) Then DifPos = O: Exit Function
Next
If LA > LB Then
    DifPos = LB + 1
Else
    DifPos = LA + 1
End If
End Function
