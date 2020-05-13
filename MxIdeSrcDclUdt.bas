Attribute VB_Name = "MxIdeSrcDclUdt"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Udt"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcDclUdt."
Private Type AA: AA() As String: End Type
'**Udt
Private Sub UdtyzMdn__Tst()
Dim Act() As Udt, Mdn$
GoSub T2
Exit Sub
T1: Mdn = "MxDaoDbSchmUd": GoTo Tst
T2: Mdn = CMdn
Tst:
    Act = UdtyzMdn(Mdn)
    Stop
    Return
End Sub

Function UdtzN(Dcl$(), Udtn$) As Udt: UdtzN = VVUdt(UdtSrc(Dcl, Udtn)): End Function


Function UdtyzM(M As CodeModule) As Udt(): UdtyzM = Udty(Dcl(M)): End Function
Function UdtyM() As Udt(): UdtyM = UdtyzM(CMd): End Function
Function UdtyzMdn(Mdn) As Udt(): UdtyzMdn = UdtyzM(Md(Mdn)): End Function
Function UdtyP() As Udt(): UdtyP = Udty(DclP): End Function
Function Udty(Dcl$()) As Udt() 'It will skip and Udtn with Opt as sfx.
Dim S: For Each S In Itr(UdtSrcy(Dcl))
    PushUdt Udty, VVUdt(CvSy(S))
Next
End Function
Function VVUdt(UdtSrc$()) As Udt
If Si(UdtSrc) = 0 Then Exit Function
Dim S$(): S = Stmt(UdtSrc)
Dim O As Udt
W1FstStmt O, S(0)
W1MidStmt O, S
W1LasStmt O, LasEle(S)
VVUdt = O
End Function
Private Sub W1FstStmt(O As Udt, FstStmt$) 'Set IsPrv & Udtn
Dim L$: L = FstStmt
O.IsPrv = ShfMdy(L) = "Private"
If Not ShfTermX(L, "Type") Then Thw CSub, "Give UdtStmt is invalid: No Udtn", "FstUdtStmt", FstStmt
O.Udtn = TakNm(L)
End Sub
Private Sub W1LasStmt(O As Udt, LasUdtStmt$) ' Set IsGenUdtAy IsGentUdtCtor IsGenUdtOpt
Dim R$: R = AftSngQ(LasUdtStmt)
O.Rmk = RmvNmBkt(R, "Deriving")
Dim Bet$: Bet = BetNmBkt(R, "Deriving")
Dim N: For Each N In ItrzSS(Bet)
    Select Case N
    Case "Ay": O.IsGenUdtAy = True
    Case "Ctor": O.IsGenUdtCtor = True
    Case "Opt": O.IsGenUdtOpt = True
    Case Else: Thw CSub, "The * in LasUdtStmt-Deriving(*) has invalid value.  Valid value are (Ay Ctor Opt)", "LasUdtStmt", LasUdtStmt
    End Select
Next
End Sub
Private Function W1MidStmt(O As Udt, UdtStmt$()) ' Set O.Mbr() by the Middle part of UdtStmt (No Fst no las stmt)
Dim M() As UdtMbr, S$
Dim J%: For J = 1 To UB(UdtStmt) - 1
    S = BrkVrmk(UdtStmt(J)).S1
    If S <> "" Then
        PushUdtMbr M, W1UdtMbr(S)
    End If
Next
O.Mbr = M
End Function
Private Function W1UdtMbr(UdtMbln) As UdtMbr
Dim L$: L = Brk1(UdtMbln, vbSngQ).S1
Dim O As UdtMbr
O.Mbn = ShfNm(L)
O.IsAy = ShfBkt(L)
If Not ShfAs(L) Then Thw CSub, "Invalid UdtMbln, [ As ] is unexpected", "UdtMbln", UdtMbln
O.Tyn = L
W1UdtMbr = O
End Function
