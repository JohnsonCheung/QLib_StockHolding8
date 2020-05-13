Attribute VB_Name = "MxDtaColTy"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "MxDtaColTy."
Type ColTy: F As String: Ty As ADODB.DataTypeEnum: End Type

Private Sub ColTyAy__Tst()
Dim A() As ColTy: A = ColTyAyzFxw(MB52IFx(LasOHYmd), "Sheet1")
End Sub

Function FnyzColTyAy(A() As ColTy) As String()
Dim J%: For J = 0 To ColTyUB(A)
    PushI FnyzColTyAy, A(J).F
Next
End Function

Function ColTyAyzFxw(Fx$, W$) As ColTy()
Dim C As Adox.Column: For Each C In CattzFxw(Fx, W).T.Columns
    PushColTy ColTyAyzFxw, ColTy(C.Name, C.Type)
Next
End Function
Sub PushColTy(O() As ColTy, M As ColTy): Dim N&: N = ColTySi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function ColTySi&(A() As ColTy): On Error Resume Next: ColTySi = UBound(A) + 1: End Function
Function ColTyUB&(A() As ColTy): ColTyUB = ColTySi(A) - 1: End Function
Function ColTy(F$, Ty As ADODB.DataTypeEnum) As ColTy
With ColTy
    .F = F
    .Ty = Ty
End With
End Function
Function FndColTy(A() As ColTy, F) As ADODB.DataTypeEnum
Dim J%: For J = 0 To ColTyUB(A)
    If A(J).F = F Then FndColTy = A(J).Ty: Exit Function
Next
Thw CSub, "@Fld not found in ColTy()", "@Fld Fny-in-@ColTy", F, Termln(FnyzColTyAy(A))
End Function
