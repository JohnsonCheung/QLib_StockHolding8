Attribute VB_Name = "MxXlsWsChkRecShdUdt"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxXlsWsChkRecShdUdt."
Enum RecShdOp
    ShdAllInLis
    ShdNBlnk
    ShdAllInTbl
    ShdAllInTbc
End Enum
Type RecShdAllInTbl
    F As String
End Type
Type RecShdAllInTbc
    F As String
End Type
Type RecShdAllInLis
    F As String
End Type
Type RecShdNBlnk
    F As String
End Type
Type RecShd
    Op As RecShdOp
    F As String
    AllInTbl As RecShdAllInTbl
    AllInLis As RecShdAllInLis
    AllInTbc As RecShdAllInTbc
    NBlnk As RecShdNBlnk
End Type
Public Const RecShdOpss = "ShdAllIn ShdNotBlnk"
Public Const RecErTySS$ = "NoEr SomBlnk MisVal SomNotIn"
Enum ColErTy
    NoEr
    Blnk
    Mis
    NInLis
    NInTbl
    NInTbc
End Enum
Type RecMisErDta
A As String
End Type
Type RecNInLisErDta
A As String
End Type
Type RecNInTblErDta
A As String
End Type
Type RecNInTbcErDta
A As String
End Type
Type RecErDta
    Ty As ColErTy
    Mis As RecMisErDta
    NInLis As RecNInLisErDta
    NInTbl As RecNInTbcErDta
    NInTbc As RecNInTblErDta
End Type

Function FnyzRecShdAy(A() As RecShd) As String()
Dim J%: For J = 0 To RecShdUB(A)
    PushI FnyzRecShdAy, A(J).F
Next
End Function

Function RecShdAy(RecShdSy$()) As RecShd()
Dim R: For Each R In Itr(RecShdSy)
    PushRecShd RecShdAy, RecShdzLin(R)
Next
End Function

Private Function RecShdzLin(Ln) As RecShd
Dim L$: L = Ln
Dim F$: F = ShfTerm(L)
Dim Op As RecShdOp: Op = RecShdOp(ShfTerm(L))
Select Case True
Case Op = ShdAllInLis:   RecShdzLin = RecShdAllIn(F, L)
Case Op = ShdNBlnk: RecShdzLin = RecShdNBlnk(F): If L = "" Then Thw CSub, ""
Case Else: Thw CSub, ""
End Select
End Function

Function RecShdAllIn(F, ParamArray ValAp()) As RecShd
With RecShdAllIn
    .F = F
    .Op = ShdAllInLis
End With
End Function

Function RecShdNBlnk(F) As RecShd
With RecShdNBlnk
    .F = F
    .Op = ShdNBlnk
End With
End Function

Private Function RecShdOp(RecShdOpStr$) As RecShdOp
Dim RecShdOpAy() As RecShdOp
Dim Ix%: Ix = IxzAy(RecShdOpAy, RecShdOpStr)
RecShdOp = Ix
End Function

Function RecShdUB&(A() As RecShd): RecShdUB = RecShdSi(A) - 1: End Function
Function RecShdSi&(A() As RecShd): On Error Resume Next: RecShdSi = UBound(A) + 1: End Function
Sub PushRecShd(O() As RecShd, M As RecShd): Dim N&: N = RecShdSi(O): ReDim Preserve O(N): O(N) = M: End Sub
