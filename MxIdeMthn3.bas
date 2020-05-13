Attribute VB_Name = "MxIdeMthn3"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthn3."
Type Mthn3: Nm As String: ShtTy As String: ShtMdy As String: End Type ' Deri(Ctor Ay)
Type Mthn4: Mdn As String: Mthn3 As Mthn3: End Type
Function Mthn3Si&(A() As Mthn3): On Error Resume Next: Mthn3Si = UBound(A) + 1: End Function
Function Mthn3UB&(A() As Mthn3): Mthn3UB = Mthn3Si(A) - 1: End Function
Sub PushMthn3(O() As Mthn3, M As Mthn3): Dim N&: N = Mthn3Si(O): ReDim Preserve O(N): O(N) = M: End Sub


Function Mthn4(Mdn, Mthn3 As Mthn3) As Mthn4
With Mthn4: .Mdn = Mdn: End With
End Function

Function Mthn3(Nm, ShtMdy, ShtTy) As Mthn3
With Mthn3
    .Nm = Nm
    .ShtMdy = ShtMdy
    .ShtTy = ShtTy
End With
End Function

Private Sub Mthn3yP__Tst()
Dim A() As Mthn3: A = Mthn3yP
Stop
End Sub

Function Mthn3yP() As Mthn3()
Mthn3yP = Mthn3yzS(SrczP(CPj))
End Function

Function Mthn3yzS(Src$()) As Mthn3()
Dim L: For Each L In MthlnyzS(Src)
    PushMthn3 Mthn3yzS, Mthn3zL(L)
Next
End Function

Function Mthn3zL(Ln) As Mthn3
Mthn3zL = ShfMthn3(CStr(Ln))
End Function

Function Mthn3yzM(M As CodeModule) As Mthn3()
Dim L: For Each L In Itr(Mthlny(RmvFalseSrc(SrczM(M))))
    PushMthn3 Mthn3yzM, Mthn3zL(L)
Next
End Function

Function ShfMthn3(OLin$) As Mthn3
Dim M$: M = ShfShtMdy(OLin)
Dim T$: T = ShfShtMthTy(OLin):: If T = "" Then Exit Function
ShfMthn3 = Mthn3(ShfNm(OLin), M, T)
End Function

Function DltMthn3$(Ln)
Const CSub$ = CMod & "DltMthn3"
Dim L$: L = Ln
RmvMdy L
If ShfMthTy(L) = "" Then Exit Function
If ShfNm(L) = "" Then Thw CSub, "Not as SrcLin", "Ln", Ln
DltMthn3 = L
End Function

Function FmtMthn3$(A As Mthn3)
With A
FmtMthn3 = JnDotAp(.Nm, .ShtMdy, .ShtTy)
End With
End Function
Sub DmPubMthn3(A As Mthn3)
D FmtMthn3(A)
End Sub
