Attribute VB_Name = "MxVbDtaPos"
Option Explicit
Option Compare Text
Const CNs$ = "Pos"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbDtaPos."
Type Loc: Pos As Long: Length As Integer: End Type 'Deriving(Ay Ctor Opt)
Type LocOpt: Som As Boolean: Loc As Loc: End Type
Type LnLoc: Lno As Long: Loc As Loc: End Type 'Deriving(Ay Ctor Opt)
Type LnLocOpt: Som As Boolean: LnLoc As LnLoc: End Type

Function TabPos&(S, Optional Nth = 1)
TabPos = SubStrPos(S, vbTab, Nth)
End Function

Function TabPairPos(S, Optional Nth) As C12
TabPairPos = SepPairPos(S, vbTab, Nth)
End Function

Function SepPairPos(S, Sep$, Optional Nth, Optional CmpMth As VbCompareMethod = VbCompareMethod.vbTextCompare) As C12
Dim P1%: P1 = SubStrPos(S, Sep, Nth, CmpMth)
If P1 = 0 Then Exit Function
Dim P2%: P2 = InStr(P1 + Len(Sep), S, Sep, CmpMth)
If P2 = 0 Then
    SepPairPos = C12(P1, Len(S) + 1)
Else
    SepPairPos = C12(P1, P2)
End If
End Function

Function SubStrPos&(S, SubStr, Optional Nth = 1, Optional CmpMth As VbCompareMethod = VbCompareMethod.vbTextCompare)
Dim Fm&: Fm = 1
Dim L%: L = Len(SubStr)
Dim O&
Dim N%: For N = 1 To Nth
    O = InStr(Fm, S, SubStr, CmpMth)
    If O = 0 Then Exit Function
    Fm = O + L
Next
SubStrPos = O
End Function
Function Loc(Pos, Length) As Loc
With Loc
    .Pos = Pos
    .Length = Length
End With
End Function
Function AddLoc(A As Loc, B As Loc) As Loc(): PushLoc AddLoc, A: PushLoc AddLoc, B: End Function
Sub PushLocAy(O() As Loc, A() As Loc): Dim J&: For J = 0 To LocUB(A): PushLoc O, A(J): Next: End Sub
Sub PushLoc(O() As Loc, M As Loc): Dim N&: N = LocUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LocSi&(A() As Loc): On Error Resume Next: LocSi = UBound(A) + 1: End Function
Function LocUB&(A() As Loc): LocUB = LocSi(A) - 1: End Function
Function LocOpt(Som, A As Loc) As LocOpt: With LocOpt: .Som = Som: .Loc = A: End With: End Function
Function SomLoc(A As Loc) As LocOpt: SomLoc.Som = True: SomLoc.Loc = A: End Function
Function LnLoc(Lno, Loc As Loc) As LnLoc
With LnLoc
    .Lno = Lno
    .Loc = Loc
End With
End Function
Function AddLnLoc(A As LnLoc, B As LnLoc) As LnLoc(): PushLnLoc AddLnLoc, A: PushLnLoc AddLnLoc, B: End Function
Sub PushLnLocAy(O() As LnLoc, A() As LnLoc): Dim J&: For J = 0 To LnLocUB(A): PushLnLoc O, A(J): Next: End Sub
Sub PushLnLoc(O() As LnLoc, M As LnLoc): Dim N&: N = LnLocUB(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LnLocSi&(A() As LnLoc): On Error Resume Next: LnLocSi = UBound(A) + 1: End Function
Function LnLocUB&(A() As LnLoc): LnLocUB = LnLocSi(A) - 1: End Function
Function LnLocOpt(Som, A As LnLoc) As LnLocOpt: With LnLocOpt: .Som = Som: .LnLoc = A: End With: End Function
Function SomLnLoc(A As LnLoc) As LnLocOpt: SomLnLoc.Som = True: SomLnLoc.LnLoc = A: End Function
