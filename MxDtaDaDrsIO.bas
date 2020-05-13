Attribute VB_Name = "MxDtaDaDrsIO"
Option Explicit
Option Compare Text
Const CNs$ = "Drs"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaDrsIO."

Sub WrtDrs(A As Drs, Ft$)
WrtAy TsyzDrs(A), Ft
End Sub

Function DrszFt(DrsFt$) As Drs
DrszFt = DrszTsy(LyzFt(DrsFt))
End Function

Function DrszTsy(DrsTsy$()) As Drs
Dim A$(): A = SplitCrLf(DrsTsy)
If Si(A) = 0 Then Thw CSub, "No lines in @Cxt"
Dim O As Drs
O.Fny = SplitTab(A(0))
If Si(A) = 1 Then Exit Function
Dim T() As VbVarType: T = VbTyAy(SplitTab(A(1)))
Dim J&: For J = 2 To UB(A)
    PushI O.Dy, DrzTsl(A(J), T)
Next
DrszTsy = O
End Function

Private Function DrzTsl(Tsl, T() As VbVarType) As Variant()
DrzTsl = DrzSy(SplitTab(Tsl), T())
End Function

Private Function DrzSy(Sy$(), T() As VbVarType) As Variant()
Dim J%, S: For Each S In Sy
    PushI DrzSy, VzS(S, T(J))
    J = J + 1
Next
End Function

Function TsyzDrs(A As Drs) As String()
PushI TsyzDrs, JnTab(A.Fny)
Dim Dr: For Each Dr In Itr(A.Dy)
    PushI TsyzDrs, JnTab(Dr)
Next
End Function
