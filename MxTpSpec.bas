Attribute VB_Name = "MxTpSpec"
Option Compare Text
Option Explicit
#If Doc Then
'Ivl:Cml #Invalid# A-kind-of-error
'Exc:Cml #Excess# A-kind-of-error
'Mis:Cml #Miss# A-kind-of-error
'IndSpecSep::Chr #Indt-Spec-Separator# a Vbar before IndSpec
'IndSpec::Tml #Indted-Spec# it a Rst-String of T4Rst of fst line of a Spec.  In format of Spect+ where + means has Pfx* and/or Sfx-
'Speci::Udt #Spec-Item# It is 1-Hdr-N-Chd Ly inside a spec
'Spect::Tyn #Spec-Type# It is
'Specit::Tyn #Spec-Item-Type#
'Specity::Ny #Spec-Item-Type-Array#
'Specin:: #Spec-Item-name#
'Lixy:: Line index array.  The index pointing given TpLy
'Tp::   #template#
'Ly::   #line-array#
'SpecFmt:: #Spec-Format#
'  #0 Fmt. 2 or more HChd.
'  .       Fst-HChd::
'  .          Hdr     is { *Spec [Specit] [Specin] [IndSpec] }
'  .          Chd     is Rmk
'  .       Rst-HChd:: Speci
'  .          Hdr     is { Specit Specn ..Rmk }
'  .          Chd     is speciILny
'  #1 No D3.  All D3 will be removed before process.
'  .          D3 is D3ln or D3str
'  .          D3ln will remove the line complete
'  .          D3str will remove and the line is kept
'  .          D3Str is string of {{--- XXXX}}
'  .          Str aft D3 will always be removed
'  #2 Rewrite.  If any error, SpecFt will be rewritten by
'  .            adding D3str at end of a line
'  .            adding D3ln infront only
'  .            in next cycle D3 is removed before process.
'  #3 Fstln.  Aft D3 is removed,
'  .          { *Spec :Spect Specn | IndSpec }
'  '          Fst term: must be *Spec, otherwise throw error
'  .          Snd term: :specTy, must with pfx :, otherwise throw error.
'  .                    The name aft : will be the Spect
'  .          Third Term: Specn
'  .          Rest: { | IndSpec }, if missing error
'  #4 IndSpec: it is SS of specTy+
'  .           speciTy+ is speciTy with optional pfx * and/or sfx -
'  .           speci is spec-item
'  .           pfx * means the speciTy is a must
'  .           sfx - means the speciTy is at most one
'  #5 HChd    [Spec] is Fst-HdrChd + Speci-HdrChd
'  .           HChd are group of lines
'  .               Hdr: with first line is not indent (fstChr ne space)
'  .               Chd: following lines that having fsrChr eq space
'  #6 FstHdrChd hcHdr is see #3 Fstln. where hc is HdrChd.
'  .            hcChd is all remark
'  #7 RstHdrChd rest of HdrChd are speci
'  .            hcHdr is [Specit] [Specin] [ShtRmk]
'  .            hcChd is ILny.  The ln is Trim and with D3Rmv
'SpecFmt:: #Spec-Format-Specification#  See !EdtSampLnkImp 1
#End If
Private Type XHChd: HdrIx As Integer: Hdr As String: Chd() As String: End Type ' Deriving(Ctor Ay)
'-- Eu
Private Type XEuIvl
    Vdt() As String  ' :Specity which is valid
    Ivl() As Itmxy   ' :Speci which is invalid
    End Type
Private Type XEuExc
    Sng() As String  ' :Specity which be 0 or 1 speci
    Exc() As Itmxy   ' :Speci which is invalid
    End Type
Private Type XEu '#Spec-Error-Udt#  Use to tell what is wrong about the :Spec.  It can convert to :SpecEu by !UUSpecEu together :Ly formated by !FmtSpecEu
    IsLnMis As Boolean
    IsSpeciMis As Boolean
    IsSigMis As Boolean
    IsSpectMis As Boolean
    IsSpecnMis As Boolean
    IsIndSpecMis As Boolean
    EuExc As XEuExc  ' Specit is Exc er.  In the spec, there is Specit not found in the IndSpec->Must.  It is the Lixy pointing to such Speci-Hdr
    EuIvl As XEuIvl  ' Specit is Ivl er.  In the spec, there is Specit not found in the IndSpec
    EuMis() As String     ' Specit is Mis er. Miss tyn error.
    IsEr As Boolean         ' If there is any in above error.
End Type


'-- Brw Spec
Sub BrwSpec(S As Spec): Brw FmtSpec(S): End Sub
Sub VcSpec(S As Spec):  Vc FmtSpec(S):  End Sub

'-- Fmt Spec
Function FmtSpec(S As Spec) As String() ' Fmt @S:Spec Hdr + Speciy
FmtSpec = W1Hdr(S)
Dim I() As Speci: I = S.Itms
Dim J&: For J = 0 To SpeciUB(I)
    PushIAy FmtSpec, W1Speci(I(J))
Next
End Function
Private Sub W1___FmtSpec(): End Sub
Private Function W1Hdr(S As Spec) As String() 'Fmt the @S:Spec-Hdr = Hdr-Ln + Rmk-Ly
PushI W1Hdr, "*Spec " & S.Spect & " " & S.Specn & " " & S.IndSpec
PushIAy W1Hdr, AmAddPfx(S.Rmk, "  ")
End Function
Private Function W1Speci(I As Speci) As String() 'Fmt one @I:Speci
PushI W1Speci, W1ItmHdr(I)
PushIAy W1Speci, W1ItmLLn(I.ILny)
End Function
Private Function W1ItmHdr$(I As Speci)  ' Fmt spec-item-header-@I as a Speci-Hdr-Ln
W1ItmHdr = I.Specit & " " & I.Specin & " " & I.Rst
End Function
Private Function W1ItmLLn(L() As ILn) As String() ' Fmt speci @L::ILny with 2 space as pfx
W1ItmLLn = AmAddPfx(LyzILny(L), "  ")
End Function

'==NwSpec
Function SpeczFt(SpecFt$) As Spec ' Load Spec from @SpecFt.  Thw Er if Er
Dim Ly$(): Ly = LyzFt(SpecFt)
Dim S As Spec: S = UUSpec(Ly)
Dim E As XEu:   E = UUEu(S)
If E.IsEr Then
    BkuFfn SpecFt
    WrtAy FmtSpecEu(Ly, UUSpecEu(E)), SpecFt, OvrWrt:=True
    Raise "Edit the SpecFt in notepad as open.  SpecFt=[" & SpecFt & "]"
End If
SpeczFt = S
End Function

Private Sub SpeczL__Tst()
Dim Act As Spec, IndLy$()
GoSub T1
Exit Sub
T1:
    IndLy = SampSchm(1)
    GoTo Tst
Tst:
    Act = SpeczL(IndLy)
    BrwAy FmtSpec(Act)
    Return
End Sub
Function SpeczL(IndLy$()) As Spec ' Thw if any Er
Dim S As Spec: S = UUSpec(IndLy)
Dim E As XEu:  E = UUEu(S)
If E.IsEr Then ChkEr FmtSpecEu(IndLy, UUSpecEu(E))
SpeczL = S
End Function

'== UU:: a lcl node (not shared).  Only called by one caller.  To Start a new Node. UU memonic Y for node, which has a point with lines, so it a node]
Private Function UUSpecEu(E As XEu) As SpecEu
UUSpecEu.Top = W4Top(E)
UUSpecEu.LnEnd = W4LnEnd(E)
End Function
Private Sub W4___UUSpecEu(): End Sub
Private Function W4Top(E As XEu) As String()
Dim A$(), B$(), C$(), D$()
A = W4Top6(E)
B = W4TopSpeciMis(E.EuMis)
C = W4TopSpeciIvl(E)
D = W4TopSpeciExc(E)
W4Top = AddSyAp(A, B, C, D)
End Function
Private Function W4TopSpeciMis(MisSpecit$()) As String()
Const Exc$ = "--- #Exc:: IndSpec indicate that there are this Speci should have only 1 such Specit.  Now it is found that there are more than 1.  So they are Exc."
Const TynEr$ = "--- #Specitn-Invalid:: IndSpec indicate that list of valid Specit, but this Specit in not in the list.  So they are TynEr."

End Function
Private Function W4Top6(E As XEu) As String()
Dim O$()
With E
    If .IsIndSpecMis Then PushI O, "IndSpec is missing."
    If .IsLnMis Then PushI O, "No line in TpLy at all"
    If .IsSigMis Then PushI O, "*Spec is missing"
    If .IsSpeciMis Then PushI O, "Speci is missing"
    If .IsSpecnMis Then PushI O, "Specn is missng"
    If .IsSpectMis Then PushI O, "Spect is missing"
End With
W4Top6 = O
End Function
Private Function W4TopSpeciIvl(E As XEu) As String()

End Function
Private Function W4TopSpeciExc(E As XEu) As String()

End Function
Private Function W4LnEnd(E As XEu) As ILn()
W4LnEnd = W4LnEndi(E.EuExc.Exc, WW2SpecitExc)
'PushILnAy W4D3ILnyzE, W4D3ILny(E.SpecitIvl, WW2SpecitIvl)
End Function
Private Function W4LnEndi(I() As Itmxy, M$) As ILn() ' It is part of :Eu
Dim Ix: For Each Ix In Itr(DisIxyzItmxyAy(I))
    PushILn W4LnEndi, ILn(Ix, "--- " & M)
Next
End Function

'--
Private Sub UUSpec__Tst()
Dim Act As Spec, IndLy$()
GoSub T1
Exit Sub
T1:
    IndLy = SampSchm(1)
    GoTo Tst
Tst:
    Act = UUSpec(IndLy)
    BrwAy FmtSpec(Act)
    Return
End Sub

Private Function UUSpec(IndLy$()) As Spec ' Load @IndLy as :Spec.  If any error find @IndLy, thw error
'@IndLy Hdrl is not no hdr space line
'       Chdl is following with space line
'       any D3Msg will be removed
If Si(IndLy) = 0 Then Thw CSub, "IndLy is empty"
Dim A$(): A = W2RmvD3(IndLy$())
If IsIndtln(A(0)) Then Thw CSub, "First chr of first line of IndLy must not be blank", "IndLy aft rmv D3", A
Dim B() As XHChd: B = W2HChdy(A)
UUSpec = W2SpecHdr(B(0))  ' Fst HChd-element is Spec-Hdr
UUSpec.Itms = W2SpecItm(B) ' Rst HChd-element are Spec-Itm
End Function
Private Sub W2___UUSpec(): End Sub
Private Function W2HChdy(IndLy$()) As XHChd()
Dim M As XHChd, Fst As Boolean: Fst = True
Dim L, Ix%: For Each L In Itr(IndLy)
    Dim IsHdr As Boolean: IsHdr = IsHdrln(L)
    Select Case True
    Case Fst And IsHdr: Fst = False: M = W2HCHdr(Ix, L)
    Case Fst:           Imposs CSub, "First and Hdr is impossible, due to it has been checed Fst Chr Fst Ln must not be blank"
    Case IsHdr: XXPushHChd W2HChdy, M
                M = W2HCHdr(Ix, L)
    Case Else:  PushI M.Chd, LTrim(L)
    End Select
    Ix = Ix + 1
Next
XXPushHChd W2HChdy, M
End Function
Private Function W2HCHdr(HdrIx, Ln) As XHChd
With W2HCHdr
    .HdrIx = HdrIx
    .Hdr = Ln
End With
End Function
Private Function W2SpecItm(I() As XHChd) As Speci()
Dim J%: For J = 1 To XXHChdUB(I)
    PushSpeci W2SpecItm, W2Speci(I(J))
Next
End Function
Private Function W2Speci(I As XHChd) As Speci
W2Speci = W2SpecizHChd(I)
W2Speci.ILny = W2SpeciILny(I)
End Function
Private Function W2SpecizHChd(I As XHChd) As Speci
With W2SpecizHChd
    .Ix = I.HdrIx
    AsgTTRst I.Hdr, .Specit, .Specin, .Rst
    Dim L: For Each L In Itr(I.Chd)
        PushILn W2SpecizHChd.ILny, ILn(.Ix, L)
    Next
End With
End Function
Private Function W2SpeciILny(I As XHChd) As ILn() ' Ret ILny by chd-ly & hdr-ix, where  chd-ly is @I.Chd and hdr-ix is @I.Hdrix.  The Chd-ly starts as hdr-ix + 1
Dim J%: For J = 0 To UB(I.Chd)
    PushILn W2SpeciILny, ILn(I.HdrIx + J + 1, I.Chd(J))
Next
End Function
Private Function W2RmvD3(TpLy$()) As String() ' Rmv all D3ln and D3SubStr
Dim L: For Each L In Itr(TpLy)
    If HasSubStr(L, "---") Then
        With Brk1(L, "---", NoTrim:=True)
            If Trim(.S1) <> "" Then
                PushI W2RmvD3, .S1
            End If
        End With
    Else
        PushI W2RmvD3, L
    End If
Next
End Function
Private Function W2SpecHdr(I As XHChd) As Spec ' ret a new Spec with Hdr is set.  Hdr is all except .Itms, ie [Spect Specn] & Rmk
With W2SpecHdr
    Dim Sig$
    Asg3TRst I.Hdr, Sig, .Spect, .Specn, .IndSpec
    If Sig <> "*Spec" Then Thw CSub, "First term of first lline should be *Spec", "@HChd.Hdr", I.Hdr
    .Rmk = I.Chd
End With
End Function

Private Function W2HdrRmk(IndLy$()) As String() ' ret Hdr rmk, which is fm snd up to next speci-hdr
Dim J%: For J = 1 To UB(IndLy)
    If FstChr(IndLy(J)) <> " " Then Exit Function
    PushI W2HdrRmk, LTrim(IndLy(J))
Next
End Function
Private Function W2SpeciHdr(Ix%, ItmHdrLn$) As Speci
With W2SpeciHdr
    AsgTTRst ItmHdrLn, .Specit, .Specin, .Rst
    .Ix = Ix
End With
End Function

'--
Private Function UUEu(S As Spec) As XEu ' return er if IndTp has any error
'@IndSpec is a SS with speciTyx as term.
'         the Specitx is Specit with optional * pfx or - sfx.
'         * pfx means must
'         - sfx means single
'         eg AA *BB- *CC DD-
'            means AA is 0-N
'            means BB is 1
'            means CC is 1-N
'            means DD is 0-1
'            VdtNN  will be AA BB CC DD
'            MustNN will be BB CC       (With * pfx)
'            SngNN  will be BB DD       (With - sfx) @@
'@IndTp is lines with HdrLn and ChdLn.  see @IndLy
Dim Vdt$(), Must$(), Sng$() ' ele of these array are: Specit
    Dim N$(): N = SyzSS(S.IndSpec)
    Vdt = W3IndSpec_Vdt(N)    ' W1 is returning Specit
    Must = W3IndSpec_Must(N)
    Sng = W3IndSpec_Sng(N)
    
Dim I() As Speci: I = S.Itms
Dim O As XEu
With O
    .EuIvl = W3EuIvl(S.Itms, Must)
    .EuMis = W3EuMis(I, Vdt) 'W1 is return error-of-Ixy%() or Misspeciny
    .EuExc = W3EuExc(I, Sng)

    .IsIndSpecMis = S.IndSpec = ""
    .IsLnMis = S.IsLnMis
    .IsSigMis = S.IsSigMis
    .IsSpeciMis = SpeciSi(S.Itms) = 0
    .IsSpecnMis = S.Specn = ""
    .IsSpectMis = S.Spect = ""
    .IsEr = .IsIndSpecMis Or .IsLnMis Or .IsSigMis Or .IsSpeciMis Or .IsSpecnMis Or .IsSpectMis Or _
        WW3IsEr(.EuIvl.Ivl) Or _
        WW3IsEr(.EuExc.Exc) Or _
        Si(.EuMis) > 0
End With
End Function
Private Sub W3___UUEu(): End Sub
Private Function W3IndSpec_Vdt(N$()) As String()
Dim I: For Each I In N
    PushI W3IndSpec_Vdt, W3RmvPfxSfx(I)
Next
End Function
Private Function W3IndSpec_Must(N$()) As String() '#
Dim I: For Each I In N
    If HasPfx(I, "*") Then PushI W3IndSpec_Must, W3RmvPfxSfx(I)
Next
End Function
Private Function W3IndSpec_Sng(N$()) As String()
Dim I: For Each I In N
    If HasSfx(I, "-") Then PushI W3IndSpec_Sng, W3RmvPfxSfx(I)
Next
End Function
Private Function W3EuMis(S() As Speci, Must$()) As String() ' Missing speciTy
W3EuMis = MinusSy(Must, Specity(S))
End Function
Private Function W3EuExc(S() As Speci, Sng$()) As XEuExc
W3EuExc.Sng = Sng
Dim Sngi: For Each Sngi In Itr(Sng) 'Sngi:Cml #single-item# Specit which is should have only single :specin
    PushItmxyAy W3EuExc.Exc, W3EuExci(Sngi, S)
Next
End Function
Private Function W3EuExci(Sngi, S() As Speci) As Itmxy() '#excess.spec.item-item#

End Function
Private Function W3EuIvl(S() As Speci, Vdt$()) As XEuIvl ' #excess.spec.item-Item#
W3EuIvl.Vdt = Vdt
Dim Specit: For Each Specit In Itr(Vdt)
    PushItmxy W3EuIvl.Ivl, W3EuIvli(Specit, S)
Next
End Function
Private Function W3EuIvli(Specit, S() As Speci) As Itmxy '#invalid.spec.item-item#

End Function
Private Function W3RmvPfxSfx$(N) ' Rmv Pfx * and Sfx -
W3RmvPfxSfx = RmvSfx(RmvPfx(N, "*"), "-")
End Function

Private Sub WW1___MsgTop(): End Sub
Private Function WW1Mis(Mis$(), Must$()) As String()
Dim O$()
PushI O, "Following Specit must exist, but they are missed:"
PushI O, ". Missed  : " & JnSpc(Mis)
PushI O, ". All Must: " & JnSpc(Must)
WW1Mis = O
End Function
Private Function WW1Exc(E As XEuExc) As String()
Dim O$()
With E
PushI O, "Following Specit must be single, but they are found more than one:"
Dim J%: For J = 0 To ItmxyUB(.Exc)
    With .Exc(J)
        PushI O, ". Exc Specit / Lno : " & .Itm & " / " & JnSpc(AmInc(.Ixy, 1))
    End With
Next
PushIAy O, " . All sng specit : " & JnSpc(.Sng)
End With
WW1Exc = O
End Function
Private Function WW1Ivl(IvlSpecity$(), VdtSpecity$()) As String()
Dim O$()
PushI O, "Following Specit are invalid:"
PushI O, ". Invalid : " & JnSpc(IvlSpecity)
PushI O, ". Valid   : " & JnSpc(VdtSpecity)
WW1Ivl = O
End Function
Private Function WW12Exc(Exc() As Itmxy) As String()

End Function

Private Sub WW2___MsgEndLn(): End Sub
Private Function WW2SpecitIvl$(): WW2SpecitIvl = "#Specit-Invalid#:End Function": End Function
Private Function WW2SpecitExc$(): WW2SpecitExc = "#Specit-Exc#:End Function": End Function

Private Sub WW3___IsEr(): End Sub
Private Function WW3IsErSpeciExc(E As XEu) As Boolean: WW3IsErSpeciExc = WW3IsEr(E.EuExc.Exc): End Function
Private Function WW3IsErSpeciIvl(E As XEu) As Boolean: WW3IsErSpeciIvl = WW3IsEr(E.EuIvl.Ivl): End Function
Private Function WW3IsEr(I() As Itmxy) As Boolean: WW3IsEr = ItmxySi(I) > 0: End Function

'==XX::FunPfx LclUdt
Private Function XXHChd(HdrIx, Hdr, Chd$()) As XHChd
With XXHChd
    .HdrIx = HdrIx
    .Hdr = Hdr
    .Chd = Chd
End With
End Function
Function XXAddHChd(A As XHChd, B As XHChd) As XHChd(): XXPushHChd XXAddHChd, A: XXPushHChd XXAddHChd, B: End Function
Sub XXPushHChdAy(O() As XHChd, A() As XHChd): Dim J&: For J = 0 To XXHChdUB(A): XXPushHChd O, A(J): Next: End Sub
Sub XXPushHChd(O() As XHChd, M As XHChd): Dim N&: N = XXHChdSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function XXHChdSi&(A() As XHChd): On Error Resume Next: XXHChdSi = UBound(A) + 1: End Function
Function XXHChdUB&(A() As XHChd): XXHChdUB = XXHChdSi(A) - 1: End Function

