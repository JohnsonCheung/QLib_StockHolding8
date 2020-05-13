Attribute VB_Name = "MxDaoDbLnkImp"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoCrttblLnkImp."
'-- Dta
Private Type ImpUd: T As String: Map() As FldMap: Bexp As String: End Type 'Deriving(Ctor Ay)
Private Type Dta
    ImpUd() As ImpUd
End Type
Private Type Psr: Dta As Dta: Er() As String: End Type
Private Type Bld: Lnk As LnkTblUd: ImpSqy() As String: End Type

Private Sub LnkiszS__Tst()
Dim S As Lnkis: S = LnkiszS(SampLnkImpSrc)
Stop
End Sub

Function LnkiszS(LnkImpSrc$()) As Lnkis
Dim S As Spec: S = SpeczL(LnkImpSrc)
Stop
With LnkiszS
    .Inp = W1Inp(S)
    .FbTbl = W1FbTbl(S)
    .FxTbl = W1FxTbl(S)
    .Stru = W1Stru(S)
    .TblWh = W1TblWh(S)
    .MustHasRecTbl = W1MustHasRecTbl(S)
End With
End Function
Function ILnyzSpec(S As Spec, Specit$) As ILn()
Dim J%: For J = 0 To SpeciUB(S.Itms)
    If S.Itms(J).Specit = Specit Then
        ILnyzSpec = S.Itms(J).ILny
        Exit Function
    End If
Next
End Function
Private Function W1Inp(S As Spec) As LnkisInp()
Dim ILn() As ILn: ILn = ILnyzSpec(S, "Inp")
Dim J%: For J = 0 To ILnUB(ILn)
    PushLnkisInp W1Inp, W1Inpi(ILn(J))
Next
End Function
Private Function W1Inpi(L As ILn) As LnkisInp

End Function
Private Function W1FbTbl(S As Spec) As LnkisFb()
Dim ILn() As ILn: ILn = ILnyzSpec(S, "FbTbl")
Dim J%: For J = 0 To ILnUB(ILn)
    PushLnkisFb W1FbTbl, W1FbTbli(ILn(J))
Next
End Function
Private Function W1FbTbli(L As ILn) As LnkisFb
Dim Inpn$, Tny$(), A$
AsgTRst L.Ln, Inpn, A
Tny = SyzSS(A)
W1FbTbli = LnkisFb(L.Ix, Inpn, Tny)
End Function
Private Function W1FxTbl(S As Spec) As LnkisFx()
Dim ILn() As ILn: ILn = ILnyzSpec(S, "FxTbl")
Dim J%: For J = 0 To ILnUB(ILn)
    PushLnkisFx W1FxTbl, W1FxTbli(ILn(J))
Next
End Function
Private Function W1FxTbli(L As ILn) As LnkisFx
Dim Inpn$, Inpnw$, Stru$, A$
AsgTRst L.Ln, Inpn, A
AsgTRst A, Inpnw, Stru
W1FxTbli = LnkisFx(L.Ix, Inpn, Inpnw, Stru)
End Function
Private Function W1Stru(S As Spec) As LnkisStru()
Dim I() As Speci: I = SpeciyzT(S, "Stru")
Dim J%: For J = 0 To SpeciUB(I)
    PushLnkisStru W1Stru, W1Strui(I(J))
Next
End Function
Private Function W1Strui(I As Speci) As LnkisStru
Dim Stru$: Stru = I.Specin
Dim Fld() As LnkisFld: Fld = W1StruiFld(I.ILny)
W1Strui = LnkisStru(I.Ix, Stru, Fld)
End Function
Private Function W1StruiFld(L() As ILn) As LnkisFld()
Dim Ix%, Intn$, Ty$, Extn$
Dim J%: For J = 0 To ILnUB(L)
    Ix = L(J).Ix
    AsgTTRst L(J).Ln, Intn, Ty, RmvSqBkt(Trim(Extn))
    PushLnkisFld W1StruiFld, LnkisFld(Ix, Intn, Ty, Extn)
Next
End Function

Private Function W1TblWh(S As Spec) As LnkisWh()
Dim ILny() As ILn: ILny = ILnyzSpec(S, "Table.Where")
Dim J%: For J = 0 To ILnUB(ILny)
    PushLnkisWh W1TblWh, W1TblWhi(ILny(J))
Next
End Function
Private Function W1TblWhi(L As ILn) As LnkisWh
Dim Tbn$, Bexp$
AsgS12 BrkSpc(L.Ln), Tbn, Bexp
W1TblWhi = LnkisWh(L.Ix, Tbn, Bexp)
End Function

Private Function W1MustHasRecTbl(S As Spec) As ILn()
W1MustHasRecTbl = ILnyzSpec(S, "MustHasRecTbl")
End Function
Private Sub LnkImp__Tst()
Dim LnkImpSrc$(), D As Database
GoSub T0
Exit Sub
T0:
    LnkImpSrc = SampLnkImpSrc
    Set D = TmpDb
    GoTo Tst
Tst:
    LnkImp D, LnkImpSrc
    Return
End Sub

Sub LnkImp(D As Database, LnkImpSrc$())
Dim S As Lnkis: S = LnkiszS(LnkImpSrc)
Stop
Dim P As Psr: P = UUPsr(S)
ChkEr P.Er, CSub
Dim B As Bld: B = UUBld(P.Dta)
LnkTblzUd D, B.Lnk
RunSqy D, B.ImpSqy
End Sub

Private Function UUPsr(S As Lnkis) As Psr

End Function
Private Function UUBld(D As Dta) As Bld
With UUBld
    .Lnk = VbldLnk(D)
    .ImpSqy = VbldImpSqy(D.ImpUd)
End With
End Function
Private Function VbldLnk(D As Dta) As LnkTblUd

End Function
Private Function VbldImpSqy(U() As ImpUd) As String()
Dim J%: For J = 0 To ImpUdUB(U)
    VbldImpSql U(J)
Next
End Function
Private Function VbldImpSql(U As ImpUd) As String
With U
Dim X$:       X = QpSelX_FldMapy(U.Map)
Dim Into$: Into = "#I" & .T
Dim Fm$:     Fm = ">" & .T
VbldImpSql = SqlSel_X_Into_Fm(X, Into, Fm, .Bexp)
End With
End Function

Private Function VbldLnkFbt(D As Dta) As LnkFbtUd()
Dim U%: U = 1
Dim J%: For J = 0 To U
    Dim T$
    Dim Fb$
    Dim SrcTbn$
    PushLnkFbtUd VbldLnkFbt, LnkFbtUd(T, Fb, SrcTbn)
Next
End Function

Function VbldLnkFxw(D As Dta) As LnkFxwUd()
Dim U%: U = 1
Dim J%: For J = 0 To U
    Dim T$
    Dim Fx$
    Dim Wsn$
    PushLnkFxwUd VbldLnkFxw, LnkFxwUd(T, Fx, Wsn)
Next
End Function

Private Function ImpUd(T, Map() As FldMap, Bexp) As ImpUd
With ImpUd
    .T = T
    .Map = Map
    .Bexp = Bexp
End With
End Function
Function AddImpUd(A As ImpUd, B As ImpUd) As ImpUd(): PushImpUd AddImpUd, A: PushImpUd AddImpUd, B: End Function
Sub PushImpUdAy(O() As ImpUd, A() As ImpUd): Dim J&: For J = 0 To ImpUdUB(A): PushImpUd O, A(J): Next: End Sub
Sub PushImpUd(O() As ImpUd, M As ImpUd): Dim N&: N = ImpUdSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function ImpUdSi&(A() As ImpUd): On Error Resume Next: ImpUdSi = UBound(A) + 1: End Function
Function ImpUdUB&(A() As ImpUd): ImpUdUB = ImpUdSi(A) - 1: End Function
