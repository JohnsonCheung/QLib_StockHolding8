Attribute VB_Name = "MxDaoDbLnk"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDbLnk."
Type LnkFxwUd: T As String: Fx As String: Wsn As String:    End Type ' Deriving(Ctor Ay)
Type LnkFbtUd: T As String: Fb As String: SrcTbn As String: End Type ' Deriving(Ctor Ay)
Type LnkTblUd: Ws() As LnkFxwUd: Tb() As LnkFbtUd: End Type

Sub LnkTblzUd(D As Database, U As LnkTblUd)
W1LnkFxwzUdy D, U.Ws
W1LnkFbtzUdy D, U.Tb
End Sub
Private Sub W1LnkFxwzUdy(D As Database, U() As LnkFxwUd): Dim J%: For J = 0 To LnkFxwUdUB(U): W1LnkFxwzUd D, U(J): Next: End Sub
Private Sub W1LnkFbtzUdy(D As Database, U() As LnkFbtUd): Dim J%: For J = 0 To LnkFbtUdUB(U): W1LnkFbtzUd D, U(J): Next: End Sub
Private Sub W1LnkFxwzUd(D As Database, U As LnkFxwUd): With U: LnkFxw D, .Fx, .Wsn:    End With: End Sub
Private Sub W1LnkFbtzUd(D As Database, U As LnkFbtUd): With U: LnkFbt D, .Fb, .SrcTbn: End With: End Sub

Sub LnkTbl(D As Database, T, S$, Cn$) ' Crt Tb-T as Lnk Tbl with @S::SrcTbn & @Cn::CnStr
Const CSub$ = CMod & "LnkTbl"
On Error GoTo X
DrpT D, T
D.TableDefs.Append W2TdzCnStr(T, S, Cn)
Exit Sub
X:
    Dim Er$: Er = Err.Description
    Thw CSub, "Error in linking table", "Er Db T SrcTbl Cn", Er, D.Name, T, S, Cn
End Sub
Private Function W2TdzCnStr(T, Src$, Cn$) As DAO.TableDef
Set W2TdzCnStr = New DAO.TableDef
With W2TdzCnStr
    .Connect = Cn
    .Name = T
    .SourceTableName = Src
End With
End Function

Sub CLnkFxw(Fx$, Wsn$, T$)
LnkFxw CDb, Fx, Wsn, T
End Sub

Sub LnkFxw(D As Database, T, Fx, Optional Wsn = "Sheet1"): LnkTbl D, T, Wsn & "$", DaoCnStrzFx(Fx):      End Sub
Sub LnkFbt(D As Database, T, Fb, Optional Fbt$):           LnkTbl D, T, DftStr(Fbt, T), DaoCnStrzFb(Fb): End Sub

Function LnkFxwUd(T, Fx, Wsn) As LnkFxwUd
With LnkFxwUd
    .T = T
    .Fx = Fx
    .Wsn = Wsn
End With
End Function
Function AddLnkFxwUd(A As LnkFxwUd, B As LnkFxwUd) As LnkFxwUd(): PushLnkFxwUd AddLnkFxwUd, A: PushLnkFxwUd AddLnkFxwUd, B: End Function
Sub PushLnkFxwUdAy(O() As LnkFxwUd, A() As LnkFxwUd): Dim J&: For J = 0 To LnkFxwUdUB(A): PushLnkFxwUd O, A(J): Next: End Sub
Sub PushLnkFxwUd(O() As LnkFxwUd, M As LnkFxwUd): Dim N&: N = LnkFxwUdSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LnkFxwUdSi&(A() As LnkFxwUd): On Error Resume Next: LnkFxwUdSi = UBound(A) + 1: End Function
Function LnkFxwUdUB&(A() As LnkFxwUd): LnkFxwUdUB = LnkFxwUdSi(A) - 1: End Function
Function LnkFbtUd(T, Fb, SrcTbn) As LnkFbtUd
With LnkFbtUd
    .T = T
    .Fb = Fb
    .SrcTbn = SrcTbn
End With
End Function
Function AddLnkFbtUd(A As LnkFbtUd, B As LnkFbtUd) As LnkFbtUd(): PushLnkFbtUd AddLnkFbtUd, A: PushLnkFbtUd AddLnkFbtUd, B: End Function
Sub PushLnkFbtUdAy(O() As LnkFbtUd, A() As LnkFbtUd): Dim J&: For J = 0 To LnkFbtUdUB(A): PushLnkFbtUd O, A(J): Next: End Sub
Sub PushLnkFbtUd(O() As LnkFbtUd, M As LnkFbtUd): Dim N&: N = LnkFbtUdSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LnkFbtUdSi&(A() As LnkFbtUd): On Error Resume Next: LnkFbtUdSi = UBound(A) + 1: End Function
Function LnkFbtUdUB&(A() As LnkFbtUd): LnkFbtUdUB = LnkFbtUdSi(A) - 1: End Function
