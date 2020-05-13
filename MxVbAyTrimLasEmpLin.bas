Attribute VB_Name = "MxVbAyTrimLasEmpLin"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Str.Lines"
Const CMod$ = CLib & "MxVbAyTrimLasEmpLin."
Function HasLasEmpLn(Ly$()) As Boolean
Dim N&: N = Si(Ly)
If N = 0 Then Exit Function
Dim O As Boolean: O = Ly(N - 1) = ""
HasLasEmpLn = O
End Function

Sub TrimLasEmpLnzFt(Ft$)
Dim Ly$(): Ly = LyzFt(Ft)
If HasLasEmpLn(Ly) Then
    WrtAy RmvLasEle(Ly), Ft
End If
End Sub
