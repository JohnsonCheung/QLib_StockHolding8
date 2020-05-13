Attribute VB_Name = "MxIdeSrcDclUdtFmt"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeSrcDclUdtFmt."
Const UdtFmtStr$ = "Se0Prv Se1Prv Se0Gen Se1Gen NoTyn NoRmk NoMbr NoGen"
Enum eUdtFmtOpt: eUfoNoOp: eUfoSe0Prv: eUfoSe1Prv: eUfoSe0Gen: eUfoSe1Gen: eUfoNoTyn: eUfoNoRmk: eUfoNoMbr: eUfoNoGen: End Enum
Private Type FWh: Prv As eSe01: Gen As eSe01: End Type
Private Type FFmt: NoTyn As Boolean: NoMbr As Boolean: NoRmk As Boolean: NoGen As Boolean: End Type '#Smr:Cml:Simplier#
Private Type Fmtr: Wh As FWh: Fmt As FFmt: End Type 'Rmk

Function FmtUdt$(A As Udt, Optional UdtFmtOpt$)
Dim S As Fmtr: S = XFmtr(UdtFmtOpt)
FmtUdt = XFmtUdt(A, S.Fmt, XNoPrv(S))
End Function

Private Sub FmtUdty__Tst()
BrwAy FmtUdty(UdtyP, "NoTyn NoGen Se0Prv NoMbr")
End Sub

Function FmtUdty(A() As Udt, Optional UdtFmtOpt$) As String()
Dim S As Fmtr: S = XFmtr(UdtFmtOpt)
Dim B() As Udt: B = W1WhUdt(A, S.Wh)
Dim O$(): O = W1FmtUdty(B, S)
Dim N%: N = 4
If XNoPrv(S) Then N = N - 1
If S.Fmt.NoGen Then N = N - 1
FmtUdty = AliLyzNTerm(O, N)
End Function

Private Function W1FmtUdty(U() As Udt, S As Fmtr) As String()
Dim NoPrv As Boolean: NoPrv = XNoPrv(S)
Dim J%: For J = 0 To UdtUB(U)
    PushI W1FmtUdty, XFmtUdt(U(J), S.Fmt, NoPrv)
Next
End Function

Private Function W1WhUdt(U() As Udt, W As FWh) As Udt() 'Simplify @Udt by @S
Dim J%: For J = 0 To UdtUB(U)
    If W1Hit(U(J), W) Then PushUdt W1WhUdt, U(J)
Next
End Function

Private Function W1Hit(U As Udt, W As FWh) As Boolean
W1Hit = W1HitPrv(U, W.Prv) And W1HItGen(U, W.Gen)
End Function

Private Function W1HitPrv(U As Udt, Prv As eSe01) As Boolean ' Is hit in term of is-private
W1HitPrv = HitSe01(U.IsPrv, Prv)
End Function

Private Function W1HItGen(U As Udt, Gen As eSe01) As Boolean ' Is hit in term of is-generate
W1HItGen = HitSe01(IsGenzUdt(U), Gen)
End Function

'---=================================================================
Private Function XFmtr(UdtFmtOpt$) As Fmtr ' Udt formatter
Dim A() As eUdtFmtOpt: A = X2Opty(UdtFmtOpt)
XFmtr.Wh.Gen = Se01(HasEle(A, eUfoSe0Gen), HasEle(A, eUfoSe1Gen))
XFmtr.Wh.Gen = Se01(HasEle(A, eUfoSe0Prv), HasEle(A, eUfoSe1Prv))
With XFmtr.Fmt
    .NoMbr = HasEle(A, eUfoNoMbr)
    .NoRmk = HasEle(A, eUfoNoRmk)
    .NoTyn = HasEle(A, eUfoNoTyn)
    .NoGen = HasEle(A, eUfoNoGen)
End With
End Function

Private Function X2Opty(UdtFmtOpt$) As eUdtFmtOpt() ' Udt-format-option-enum array
Dim N: For Each N In Itr(SyzSS(UdtFmtOpt))
    Select Case N
    Case "Se0Prv": PushI X2Opty, eUfoSe0Prv
    Case "Se1Prv": PushI X2Opty, eUfoSe1Prv
    Case "Se0Gen": PushI X2Opty, eUfoSe0Gen
    Case "Se1Gen": PushI X2Opty, eUfoSe1Gen
    Case "Se1Gen": PushI X2Opty, eUfoSe1Gen
    Case "NoTyn": PushI X2Opty, eUfoNoTyn
    Case "NoRmk": PushI X2Opty, eUfoNoRmk
    Case "NoMbr": PushI X2Opty, eUfoNoMbr
    Case "NoGen": PushI X2Opty, eUfoNoGen
    Case Else: Thw CSub, "UdtFmtOpt has invalid item", "Invalid-item ValidItem", N, UdtFmtStr
    End Select
Next
End Function
'----------------------------------------------------------------------
Function XFmtUdt$(U As Udt, F As FFmt, NoPrv As Boolean) ' a line of formating a Udt
With F
XFmtUdt = X1Hdr(U, NoPrv, .NoGen) & X1Mbr(U.Mbr, .NoMbr) & X1Rmk(U.Rmk, .NoRmk)
End With
End Function

Private Function X1Rmk$(Rmk$, NoRmk As Boolean) ' [Remark]-column string value
If NoRmk Then Exit Function
X1Rmk = " ' " & Rmk
End Function

Private Function X1Hdr$(U As Udt, NoPrv As Boolean, NoGen As Boolean) ' header part of formatting a Udt
Dim Gen$: Gen = X1Gen(U, NoGen)
Dim Prv$: Prv = IIf(NoPrv, "", IIf(U.IsPrv, " Prv", " ."))
X1Hdr = FmtQQ("Udt ? ? ?", U.Udtn, Prv, Gen)
End Function

Private Function X1Gen$(U As Udt, NoGen As Boolean)  ' [Any-generate]-column
If NoGen Then Exit Function
Dim O$()
With U
    If .IsGenUdtCtor Then PushI O, "Ctor"
    If .IsGenUdtAy Then PushI O, "Ay"
    If .IsGenUdtOpt Then PushI O, "Opt"
End With
If Si(O) = 0 Then X1Gen = " .": Exit Function
X1Gen = " " & JnDot(O)
End Function

Private Function X1Mbr$(A() As UdtMbr, NoMbr As Boolean) ' member-column
If NoMbr Then Exit Function
Dim O$()
Dim J%: For J = 0 To UdtMbrUB(A)
    PushI O, X1Mbri(A(J))
Next
X1Mbr = " " & JnSpc(O)
End Function

Private Function X1Mbri$(A As UdtMbr) ' one member item string value
With A
X1Mbri = .Mbn & X1Sfx(.IsAy, .Tyn)
End With
End Function

Private Function X1Sfx$(IsAy As Boolean, Tyn$) ' member item suffix
Dim O$
    If Tyn <> "Variant" And Tyn <> "" Then
        Dim T$: T = TyChrzN(Tyn)
        If T = "" Then
            O = ":" & ShtTyn(Tyn)
        Else
            O = T
        End If
    End If
X1Sfx = O & IIf(IsAy, "()", "")
End Function
'-----------------------------------------
Private Function XNoPrv(S As Fmtr) As Boolean
XNoPrv = S.Wh.Prv <> eSeAll
End Function
