Attribute VB_Name = "MxDtaDaDs"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaDs."
Function AddDsDt(Ds As Ds, Dt As Dt) As Ds ' add ds dt becomes ds
If HasDt(Ds, Dt.DtNm) Then Thw CSub, "@Ds already has @Dt", "Ds Dt", Ds.DsNm, Dt.DtNm
AddDsDt = Ds
PushDt AddDsDt.DtAy, Dt
End Function



Function DtzDs(A As Ds, DtNm$) As Dt
Const CSub$ = CMod & "DtzDsNm"
Dim Ay() As Dt: Ay = A.DtAy
Dim J&: For J = 0 To DtUB(Ay)
    If Ay(J).DtNm = DtNm Then
        DtzDs = Ay(J)
        Exit Function
    End If
Next
Thw CSub, "No such DtNm in Ds", "Such-DtNm DtNy-In-Ds", DtNm, TnyzDs(A)
End Function
Function HasDt(A As Ds, DtNm$) As Boolean
Dim Ay() As Dt: Ay = A.DtAy
Dim J&: For J = 0 To DtUB(A.DtAy)
    If Ay(J).DtNm = DtNm Then HasDt = True: Exit Function
Next
End Function
Function DtSi&(A() As Dt): On Error Resume Next: DtSi = UBound(A) + 1: End Function
Function DtUB&(A() As Dt): DtUB = DtSi(A) - 1: End Function
Sub PushDt(O() As Dt, M As Dt): Dim N&: N = DtSi(O): ReDim Preserve O(N): O(N) = M: End Sub

Function TnyzDs(A As Ds) As String()
Dim Ay() As Dt: Ay = A.DtAy
Dim J&: For J = 0 To DtUB(Ay)
    PushI TnyzDs, Ay(J).DtNm
Next
End Function

Function VzDicIf$(A As Dictionary, K)
If A.Exists(K) Then VzDicIf = A(K)
End Function

Function VzDicK(A As Dictionary, K, Optional Dicn$ = "Dic", Optional Kn$ = "Key", Optional Fun$)
If A.Exists(K) Then VzDicK = A(K): Exit Function
Dim M$: M = FmtQQ("[?] does not [?]", Dicn, Kn)
Dim NN$: NN = FmtQQ("[?] [?]", Dicn, Kn)
Thw Fun, M, NN, A, K
End Function
