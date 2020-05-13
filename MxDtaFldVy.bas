Attribute VB_Name = "MxDtaFldVy"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "MxDtaFldVy."
Type FldVy
    F As String
    Vy() As Variant
End Type
Function FldVy(F, Vy()) As FldVy
With FldVy
    .F = F
    .Vy = Vy
End With
End Function

Private Sub DisFldVyAyzFx__Tst()
Dim A() As FldVy: A = DisFldVyAyzFx(MB52LasIFx, MB52Wsn, Termy("[Base Unit of measure] Material Plant"))
Stop
End Sub

Function DisFldVyAyzFx(Fx$, W$, Fny$()) As FldVy()
Dim F: For Each F In Fny
    PushFldVy DisFldVyAyzFx, FldVy(F, DisFvyzFx(Fx, W, CStr(F)))
Next
End Function

Function FnyzFldVyAy(A() As FldVy) As String()
Dim J%: For J = 0 To FldVyUB(A)
    PushI FnyzFldVyAy, A(J).F
Next
End Function

Function FFzFldVyAy$(A() As FldVy)
FFzFldVyAy = Termln(FnyzFldVyAy(A))
End Function

Sub PushFldVy(O() As FldVy, M As FldVy): Dim N&: N = FldVySi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function FldVySi&(A() As FldVy): On Error Resume Next: FldVySi = UBound(A) + 1: End Function
Function FldVyUB&(A() As FldVy): FldVyUB = FldVySi(A) - 1: End Function
Function FndVy(A() As FldVy, F$) As Variant()
Dim J%: For J = 0 To FldVyUB(A)
    With A(J)
        If .F = F Then
            FndVy = .Vy
            Exit Function
        End If
    End With
Next
Thw CSub, "@Fld not found in given @FldVyAy", "@Fld FF-in-@FldVyAy", F, FnyzFldVyAy(A)
End Function
