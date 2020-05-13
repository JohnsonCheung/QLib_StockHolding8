Attribute VB_Name = "MxXlsWsChkWsCol"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsWsChkWsCol."
Private Type ColEr: F As String: ActTy As ADODB.DataTypeEnum: EptTy As EmXlsTy: End Type
Private Type EptTy: F As String: Ty As EmXlsTy: End Type

Private Sub ChkWsCol__Tst()
Dim Fx$, W$, FldNmCsv$, FldTyCsv$
Fx = MB52LasIFx
GoSub Z1
Exit Sub
Z1:
    W = MB52Wsn
    FldNmCsv = MB52FldNmCsv
    FldTyCsv = MB52FldTyCsv
    GoTo Tst
Tst:
    ChkWsCol Fx, W, FldNmCsv, FldTyCsv
    Return
End Sub
Sub ChkWsCol(Fx$, W$, FldNmCsv$, FldTyCsv$)
Dim E() As EptTy: E = EptTyAyzCsv(FldNmCsv, FldTyCsv)
Dim A() As ColTy: A = ColTyAyzFxw(Fx, W)
Dim Er1$(): Er1 = FldMisMsg(FnyzEpt(E), FnyzAct(A))
Dim Er2$(): Er2 = ColErAyMsg(ColErAy(E, A))
ChkFxwEr Fx, W, AddSy(Er1, Er2)
End Sub

Private Function FnyzAct(Act() As ColTy) As String()
FnyzAct = FnyzColTyAy(Act)
End Function
Private Function FnyzEpt(E() As EptTy) As String()
Dim J%: For J = 0 To EptTyUB(E)
    PushI FnyzEpt, E(J).F
Next
End Function

Private Function EptTyAyzCsv(FldNmCsv$, FldTyCsv$) As EptTy()
Dim Fny$(): Fny = AmTrim(Split(FldNmCsv, ","))
Dim Ty() As EmXlsTy: Ty = XlsTyAyzCsv(FldTyCsv)
If Si(Fny) <> Si(Ty) Then Thw "EptTyAyzCsv", "Sz(Fny)=<>Si(XlsTyAy)", "FldNmCsv FldTyCsv", FldNmCsv, FldTyCsv
Dim J%: For J = 0 To UB(Fny)
    PushEptTy EptTyAyzCsv, EptTy(Fny(J), Ty(J))
Next
End Function

Private Function FldMisMsg(EptFny$(), ActFny$()) As String()
Dim MisFny$(): MisFny = MinusSy(EptFny, ActFny)
If Si(MisFny) = 0 Then Exit Function
Dim O$()
PushI O, ""
PushI O, "There are missing Column"
PushI O, "========================"
PushI O, Si(MisFny) & " Missing Column:"
Dim J%: For J = 0 To UB(MisFny)
    PushI O, vbTab & J + 1 & " [" & MisFny(J) & "]"
Next
PushI O, Si(ActFny) & " Worksheet columns:"
For J = 0 To UB(ActFny)
    PushI O, vbTab & J + 1 & " [" & ActFny(J) & "]"
Next
FldMisMsg = O
End Function

Private Function ColErAyMsg(A() As ColEr) As String()
Dim NEr%: NEr = ColErSi(A)
If NEr = 0 Then Exit Function
Dim O$()
PushI O, ""
PushUL O, "There are [" & NEr & "] column has unexpected data type:"
Dim J%: For J = 0 To NEr - 1
    PushI O, ColErMsgl(J, A(J))
Next
End Function
Private Function ColErMsgl$(Ix%, A As ColEr)
Dim F$: F = A.F
Dim Act$: Act = ShtDaoTy(A.ActTy)
Dim Ept$: Ept = ShtXlsTy(A.EptTy)
ColErMsgl = FmtQQ("? Col[?] should be [?] but now [?]", Ix + 1, F, Act, Ept)
End Function

Private Function ColErAy(E() As EptTy, A() As ColTy) As ColEr()
Dim Exist$(): Exist = IntersectAy(FnyzEpt(E), FnyzAct(A))
Dim F: For Each F In Itr(Exist)
    Dim ETy As EmXlsTy: ETy = FndEptTy(E, F)
    Dim ATy As ADODB.DataTypeEnum: ATy = FndColTy(A, F)
    If Not IsEqXlsTy(ETy, ATy) Then
        PushColEr ColErAy, ColEr(F, ATy, ETy)
    End If
Next
End Function
Private Function ColEr(F, ActTy As ADODB.DataTypeEnum, EptTy As EmXlsTy) As ColEr
With ColEr
    .F = F
    .ActTy = ActTy
    .EptTy = EptTy
End With
End Function
Private Sub PushColEr(O() As ColEr, M As ColEr): Dim N&: N = ColErSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Function FndEptTy(E() As EptTy, F) As EmXlsTy
Dim J%: For J = 0 To EptTyUB(E)
    If E(J).F = F Then FndEptTy = E(J).Ty: Exit Function
Next
Thw CSub, "@Fld not found in EptTy()", "@Fld Fny-in-@EptTy", F, Termln(FnyzEpt(E))
End Function


Private Function EptTy(F, Ty As EmXlsTy) As EptTy
With EptTy
    .F = F
    .Ty = Ty
End With
End Function
Private Function EptTyUB&(A() As EptTy): EptTyUB = EptTySi(A) - 1: End Function
Private Function EptTySi&(A() As EptTy): On Error Resume Next: EptTySi = UBound(A) + 1: End Function
Private Sub PushEptTy(O() As EptTy, M As EptTy): Dim N&: N = EptTySi(O): ReDim Preserve O(N): O(N) = M: End Sub

Function ColErSi&(A() As ColEr)

End Function
