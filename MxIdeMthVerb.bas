Attribute VB_Name = "MxIdeMthVerb"
Option Explicit
Option Compare Text
Const CNs$ = "Mth.Verb"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthVerb."
':Rx: :RegExp #Regular-Expression#
Function HasVerb(Nm) As Boolean
HasVerb = Verb(Nm) <> ""
End Function


Function MthVerbNyV() As String()
Dim Mthn: For Each Mthn In Itr(MthnyV)
    PushI MthVerbNyV, Mthvn(Mthn)
Next
End Function

Private Sub MthVerbAetV__Tst()
VcAet SrtAet(MthVerbAetV)
End Sub

Function MthVerbAetV() As Dictionary
Set MthVerbAetV = Aet(MthVerbNyV)
End Function

Function MthQVerbNsetV() As Dictionary
Set MthQVerbNsetV = Aet(MthQVerbNyV)
End Function

Sub VcMthQVerbNsetV()
VcAet SrtAet(MthQVerbNsetV)
End Sub

Sub VcMthQVerbNyV()
VcAet Aet(MthQVerbNyV).Srt
End Sub

Function MthQVerbNyV() As String() '6204
MthQVerbNyV = MthQVerbNyzV(CVbe)
End Function

Function MthQVerbNyzV(A As Vbe) As String()
MthQVerbNyzV = QVerbNy(MthnyzV(A))
End Function

Function QVerbNy(Ny$()) As String()
Dim N: For Each N In Itr(Ny)
    PushI QVerbNy, QVerbNm(N)
Next
End Function

Function QVerbNm$(Nm)
Dim V$: V = Verb(Nm)
If V = "" Then
    QVerbNm = "#" & Nm
Else
    QVerbNm = Replace(Nm, V, QuoBkt(V), Count:=1)
End If
End Function

Function Mthvn$(Mthn)
Mthvn = Verb(Mthn) & "." & Mthn
End Function

Function VerbRx() As RegExp
Static X As RegExp
If IsNothing(X) Then Set X = Rx(VerbPatn(VerbSS))
Set VerbRx = X
End Function

Sub BrwVerb()
Brw SyzSS(VerbSS)
End Sub

Sub VcNVTDNmAetV()
VcAet NVTDNmAetV.Srt
End Sub
Property Get NVTDNmAetV() As Dictionary
Set NVTDNmAetV = Aet(NVTDNyV)
End Property
Property Get NVTDNyV() As String()
NVTDNyV = NVTDNyzV(CVbe)
End Property
Function NVTDNyzV(A As Vbe) As String()
NVTDNyzV = NVTDNy(MthnyzV(A))
End Function
Function NVTDNy(Ny$()) As String()
Dim Nm$, I
For Each I In Itr(Ny)
    Nm = I
    PushI NVTDNy, NVTDNm(Nm)
Next
End Function
Function NVTDNm$(Nm) 'Nm.Verb.Ty.Dot-Nm
NVTDNm = NVTy(Nm) & "." & Nm
End Function
Function FstVerbSubNyV() As String()

End Function
Function NVTy$(Nm) 'Nm.Verb-Ty
Const CSub$ = CMod & "NVTy"
Select Case True
Case IsNoVerbNm(Nm): NVTy = "NoVerb"
Case IsFstVerbNm(Nm): NVTy = "FstVerb"
Case IsMidVerbNm(Nm): NVTy = "MidVerb"
Case Else: Thw CSub, "Program error: a Nm must be any of [NoVerb | FstVerb | MidVerb]", "Nm", Nm
End Select
End Function
Function IsNoVerbNm(Nm) As Boolean
IsNoVerbNm = Verb(Nm) = ""
End Function
Function IsMidVerbNm(Nm) As Boolean
Dim V$: V = Verb(Nm): If V = "" Then Exit Function
IsMidVerbNm = Not HasPfx(Nm, Verb(Nm))
End Function

Function IsFstVerbNm(Nm) As Boolean
IsFstVerbNm = HasPfx(Nm, Verb(Nm))
End Function

Function IsVerb(S) As Boolean
IsVerb = VerbAet.Exists(S)
End Function

Property Get VerbAet() As Dictionary
Static X As Dictionary
If IsNothing(X) Then Set X = AetzSS(VerbSS)
Set VerbAet = X
End Property

Function Verb$(Nm) 'ret the-verb of a name.
Dim Cml$, I, LetterCml$
For Each I In CmlAy(Nm)
    Cml = I
    LetterCml = RmvDigSfx(Cml)
    If VerbAet.Exists(LetterCml) Then Verb = Cml: Exit Function
Next
End Function
Function VerbSS$()
Const C$ = "Zip Wrt Wrp Wait Vis Vc ULn UnRmk UnEsc Trim Tile Thw Tak Sye Swap Sum Stop" & _
" Srt Split Solve Shw Shf Set Sel Sav Run Rpl Rmv Rmk Rfh Rev Resi Ren ReSz ReSeq ReOrd RTrim Quo" & _
" Quit Push Pmpt Pop Opn Nxt Nrm New Mov Mk Minus Min Mid Mge Max Map Lnk Lis Lik Las Kill Jn Jmp Is" & _
" IntersectAy Ins Ini Inf Indt Inc Imp Hit Has Halt Gen Fst Fmt Flat Fill Extend Expand Exp Xls" & _
" Evl Esc Ens EndTrim Edt Dyw Dye Dw De Drp Down Do Dmp Dlt Cv Cut Crt Cpy Compress Cls Clr Clone Cln" & _
" Chk3 Chk2 Chk1 Chk Chg Cfm Brw Brk Box Bld Bet Below Bef Bdr Bku Aw Ae AutoFit AutoExec Ass Asg" & _
" And Ali Aft Add Above"
VerbSS = NrmSS(C)
End Function

Function NrmSS$(SS$) ' Normalize Verb
NrmSS = JnSpc(QSrt(AwDis(SyzSS(SS))))
End Function

Function VerbPatn$(SSoVerb$)
Dim O$(), Verb$, I
For Each I In Aet(SyzSS(SSoVerb)).Itms
    Verb = I
    PushI O, VerbPatn_(Verb)
Next
VerbPatn = QuoBkt(JnVBar(O))
End Function

Function VerbPatn_$(Verb$)
ChktVerb Verb, CSub
VerbPatn_ = Verb & "[^a-z|0-9]*"
End Function

Sub ChktVerb(S, Fun$)
If Not IsNm(S) Then Thw Fun, "Verb must be a name", "Str", S
If Not IsAscUCas(Asc(FstChr(S))) Then Thw Fun, "Verb must started with UCase", "Str", S
End Sub

Function QuoVerb$(Nm)

End Function
Function WoVerbMthnyP() As String()
WoVerbMthnyP = WoVerbMthnyzP(CPj)
End Function

Function WoVerbMthnyzP(P As VBProject) As String()
Dim N: For Each N In Itr(MthnyzP(P))
    If Not HasVerb(N) Then PushI WoVerbMthnyzP, N
Next
End Function
