Attribute VB_Name = "MxIdeSrcGenDeriUdtDvc"
Option Explicit
Option Compare Text
      Const W1TpAddFun$ = "Function Add?(A As ?, B As ?) As ?(): Push? Add?, A: Push? Add?, B: End Function"
     Const W1TpPushFun$ = "Sub Push?(O() As ?, M As ?): Dim N&: N = ?Si(O): ReDim Preserve O(N): O(N) = M: End Sub"
   Const W1TpPushAyFun$ = "Sub Push?Ay(O() As ?, A() As ?): Dim J&: For J = 0 To ?UB(A): Push? O, A(J): Next: End Sub"
       Const W1TpSiFun$ = "Function ?Si&(A() As ?): On Error Resume Next: ?Si = UBound(A) + 1: End Function"
       Const W1TpUBFun$ = "Function ?UB&(A() As ?): ?UB = ?Si(A) - 1: End Function"
    Const W1TpUdtlForOpt$ = "Type ?Opt: Som As Boolean: ? As ?: End Type"
    Const W1TpOptCtorFun$ = "Function ?Opt(Som, A As ?) As ?Opt: With ?Opt: .Som = Som: .? = A: End With: End Function"
     Const W1TpOptSomFun$ = "Function Som?(A As ?) As ?Opt: Som?.Som = True: Som?.? = A: End Function"
    
        Const W2TpCtorln$ = "Function ?(?) As ?"
 
       Const W1TpMthnnAy$ = "?Si ?UB Add? Push? Push?Ay"
     Const W1TpMthnnCtor$ = "?"
      Const W1TpMthnnOpt$ = "?Opt Som?"

'--- UdtDvc #Dvc:Dv-c:Derived-Cd#
Type UdtDvc
    Mthl As String          ' Is AddLinesAp of CtorMthl, AyMthl OptUdtl
    OptUdtl As String
    Mthny() As String      '
End Type
'---
Const CNs$ = "Src.Deri"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcGenDeriUdtDvc."

'**UdtDvc
Sub UdtDvcPzN__Tst()
Dim Udtn$, Act As UdtDvc, Ept As UdtDvc
GoSub Z
Exit Sub
Z:
    Udtn = "S12"
    GoTo Tst
Tst:
    Act = UdtDvcPzN(Udtn)
    Stop
    Return
End Sub
Sub UdtDvc__Tst()
Dim O$(), U() As Udt: U = UdtyP
Dim J%: For J = 0 To UdtUB(U)
    GoSub Push
Next
BrwLinesy O
Exit Sub
Push:
    With UdtDvc(U(J))
        PushNB O, AddLinesAp(JnSpc(.Mthny), .OptUdtl, .Mthl)
    End With
    Return
End Sub
Function UdtDvcPzN(Udtn$) As UdtDvc: UdtDvcPzN = UdtDvc(UdtPzN(Udtn)): End Function '#Fst-UdtDvc-In-CPj#
Function UdtDvc(U As Udt) As UdtDvc
If Not IsGenzUdt(U) Then Exit Function
Stop
Dim O As UdtDvc
Dim N$: N = U.Udtn
Dim A$, B$, C$
If U.IsGenUdtCtor Then A = MthForCtor(U)
  If U.IsGenUdtAy Then B = W1MthForAy(U.IsPrv, N)
 If U.IsGenUdtOpt Then C = W1MthForOpt(U.IsPrv, N)
 If U.IsGenUdtOpt Then O.OptUdtl = W1UdtlForOpt(U.IsPrv, N)
 O.Mthny = W1Mthny(U)
O.Mthl = AddLinesAp(A, B, C)
UdtDvc = O
End Function
Private Function W1Mthny(U As Udt) As String()
Dim N$: N = U.Udtn
Dim O$()
    With U
    If .IsGenUdtCtor Then PushI O, N
      If .IsGenUdtAy Then PushIAy O, SyzSS(RplQ(W1TpMthnnAy, N))
     If .IsGenUdtOpt Then PushIAy O, SyzSS(RplQ(W1TpMthnnOpt, N))
    End With
W1Mthny = O
End Function
Private Function W1MthForAy$(IsPrv As Boolean, Udtn$)
Dim O$()
PushI O, RplQ(W1TpAddFun, Udtn)
PushI O, RplQ(W1TpPushAyFun, Udtn)
PushI O, RplQ(W1TpPushFun, Udtn)
PushI O, RplQ(W1TpSiFun, Udtn)
PushI O, RplQ(W1TpUBFun, Udtn)
W1MthForAy = JnCrLf(O)
End Function
Private Function W1UdtlForOpt$(IsPrv As Boolean, Udtn$)
W1UdtlForOpt = RplQ(W1TpUdtlForOpt, Udtn)
End Function
Private Function W1MthForOpt$(IsPrv As Boolean, Udtn$)
Dim A$: A = RplQ(W1TpOptCtorFun, Udtn): A = AddPrv(A, IsPrv)
Dim B$: B = RplQ(W1TpOptSomFun, Udtn):  B = AddPrv(B, IsPrv)
W1MthForOpt = AddLines(A, B)
End Function
'---=============================
Private Function MthForCtor$(U As Udt)
Dim O$()
    Dim M() As UdtMbr: M = U.Mbr
    Dim N$: N = U.Udtn
    PushI O, W2Ctorln(U)
    PushI O, "With " & N
    Dim J%: For J = 0 To UdtMbrUB(M)
        PushI O, "    " & W2Mbln(M(J))
    Next
    PushI O, "End With"
    PushI O, "End Function"
MthForCtor = JnCrLf(O)
End Function

Private Function W2Ctorln$(U As Udt)
Dim N$: N = U.Udtn
Dim Pm$: Pm = W2Pm(U.Mbr)
Dim O$: O = FmtQQ(W2TpCtorln, N, Pm, N)
If U.IsPrv Then O = "Private " & O
W2Ctorln = O
End Function

Private Function W2Pm$(M() As UdtMbr)
Dim O$()
Dim J%: For J = 0 To UdtMbrUB(M)
    PushI O, W2Arg(M(J))
Next
W2Pm = JnCommaSpc(O)
End Function

Private Function W2Arg$(M As UdtMbr)
Dim N$: N = M.Mbn
Dim T$: T = M.Tyn
Dim O$
Dim IsPrim As Boolean: IsPrim = IsPrimTy(M.Tyn)
Select Case True
Case IsPrim And M.IsAy: O = FmtQQ("??()", N, TyChrzN(T))
Case IsPrim:            O = N
Case M.IsAy:            O = FmtQQ("?() As ?", N, T)
Case Else:              O = FmtQQ("? As ?", N, T)
End Select
W2Arg = O
End Function

Private Function W2Mbln$(U As UdtMbr) ' The Udt constructor member line
W2Mbln = RplQ(W2Tp(U), U.Mbn)
End Function

Private Function W2Tp$(U As UdtMbr) ' The template for Udt Constructor member line
Const SetTp$ = "Set .? = ?"
Const AsgTp$ = ".? = ?"
Select Case True
Case Not U.IsAy And IsObjTyn(U.Tyn): W2Tp = SetTp
Case Else: W2Tp = AsgTp
End Select
End Function
'---=============================

Function SampUdt1() As Udt
With SampUdt
    .IsPrv = True
    .Udtn = "Xyz"
    PushUdtMbr .Mbr, UdtMbr(True, "AA", "Integer")
    PushUdtMbr .Mbr, UdtMbr(True, "BB", "Integer")
End With
End Function
