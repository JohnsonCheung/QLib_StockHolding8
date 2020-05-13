Attribute VB_Name = "MxIdeMthLnoCnt"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthlnoCnt."
#If Doc Then
'Lcnt:Fun|Sub has one StartLineNo|Count.  Prp may have 2.
#End If
Type Lcnt: Lno As Long: Cnt As Long: End Type 'Deriving(Ctor Ay Opt)
Type Lcnt2: A As Lcnt: B As Lcnt: End Type
Function Lcnt(Lno&, Cnt&) As Lcnt: With Lcnt: .Lno = Lno: .Cnt = Cnt: End With: End Function
Sub PushLcnt(O() As Lcnt, M As Lcnt): Dim N&: N = LcntSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Sub PushLcntAy(O() As Lcnt, A() As Lcnt): Dim J&: For J = 0 To LcntUB(O): PushLcnt O, A(J): Next: End Sub
Function LcntSi&(A() As Lcnt): On Error Resume Next: LcntSi = UBound(A) + 1: End Function
Function LcntUB&(A() As Lcnt): LcntUB = LcntSi(A): End Function
Function LcntStr$(A As Lcnt)
LcntStr = FmtQQ("Lcnt(? ?)", A.Lno, A.Cnt)
End Function
Function Lcnt2(A As Lcnt, B As Lcnt) As Lcnt2: With Lcnt2: .A = A: .B = B: End With: End Function
Function LcntByBei(Bix&, Eix&) As Lcnt
With LcntByBei
    .Lno = Bix + 1
    .Cnt = Eix - Bix + 1
End With
End Function
Function LcntzBei(A As Bei) As Lcnt
If IsEmpBei(A) Then Exit Function
LcntzBei = Lcnt(A.Bix + 1, A.Eix - A.Bix + 1)
End Function

Function IsEmpLcnt(A As Lcnt) As Boolean
Select Case True
Case A.Cnt <= 0, A.Lno <= 0: IsEmpLcnt = True
End Select
End Function
