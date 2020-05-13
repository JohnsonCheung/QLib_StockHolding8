Attribute VB_Name = "MxIdeMthn"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthn."
':Mthn: :Nm ! Rule1-FstVerbBeingDo: the mthn will not return any value
'       ! Rule2-FstVerbBeingDo: tThe Cmls aft Do is a verb
':Dta_MthQVNm: :Nm ! It is a String dervied from Nm.  Q for quoted.  V for verb.  It has 3 Patn: NoVerb-[#xxx], MidVerb-[xxx(vvv)xxx], FstVerb-[(vvv)xxx]."
':Nm: :S ! less that 64 chr.
':FunNm: :Rul ! If there is a Subj in pm, put the Subj as fst CmlTerm and return that Subj;
'       ! give a Noun to the subj noun is MulCml.
'       ! Each Mthn must belong to one of these rule:
'       !   Noun | Noun.Verb.Extra | Verb.Variant | Noun.z.Variant
'       ! Pm-Rule
'       !   Subj    : Choose a subj in pm if there is more than one arg"
'       !   MuliNoun: It is Ok to group mul-arg as one subj
'       !   MulNounUseOneCml: Mul-noun as one subj use one Cml
':Noun: :Nm  ! it is 1 or more Cml to form a Noun."
':Cml:  :Nm  ! Tag:Type. P1.NumIsLCase:.  P2.LowDashIsLCase:.  P3.FstChrCanAnyNmChr:.
':Sfxss: :SS !  NmRul means variable or function name.
':VdtVerss: :SS ! P1.Opt: Each module may one DoczVdtSSoVerb.  P2.OneOccurance: "
':NounVerbExtra :SS !Tag: FunNmRule.  Prp1.TakAndRetNoun: Fst Cml is Noun and Return Noun.  Prp2.OneCmlNoun: Noun should be 1 Cml.  " & _
'                ! Prp3.VdtVerb: Snd Cml should be approved/valid noun.  Prp4.OptExtra: Extra is optional."

Function MthDnzM$(M As CodeModule, Ln)
Dim D$: D = MthDnzL(Ln): If D = "" Then Exit Function
MthDnzM = MdDn(M) & "." & D
End Function
Function MthnnM$()
MthnnM = MthnnzMd(CMd)
End Function

Function MthnnzMd$(M As CodeModule)
MthnnzMd = Mthnn(Src(M))
End Function

Function Mthnn$(Src$())
Mthnn = JnSpc(SrtAy(Mthny(Src)))
End Function

Function PubMthnnzM$(M As CodeModule)
PubMthnnzM = JnSpc(QSrt(PubMthnyzM(M)))
End Function

Function PubMthnnP$()
PubMthnnP = PubMthnnzP(CPj)
End Function

Function PubMthnnzP$(P As VBProject)
PubMthnnzP = JnSpc(PubMthnyzP(P))
End Function

Function PubMthnyzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy PubMthnyzP, PubMthnyzM(C.CodeModule)
Next
End Function

Function PubMthnnM$()
PubMthnnM = PubMthnnzM(CMd)
End Function

Function PubMthnyzM(M As CodeModule) As String()
PubMthnyzM = PubMthnyzS(Src(M))
End Function

Function PubMthnyzS(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB PubMthnyzS, PubMthn(L)
Next
End Function

Function PubMthMknyP() As String()
PubMthMknyP = PubMthMknyzP(CPj)
End Function

Function PubMthMknyzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy PubMthMknyzP, PubMthMknyzM(C.CodeModule)
Next
End Function

Function PubMthMknyzM(M As CodeModule) As String()
PubMthMknyzM = PubMthMknyzMS(Mdn(M), Src(M))
End Function

Function PubMthMknyzMS(Mdn$, Src$()) As String()
Dim K$
Dim L: For Each L In Itr(Src)
    K = MthKnzL(L)
    PushNB PubMthMknyzMS, PubMthMknzML(Mdn, L)
Next
End Function

Function PubMthMknzML$(Mdn$, Ln)
Dim K$: K = MthKnzL(Ln): If K = "" Then Exit Function
PubMthMknzML = Mdn & "." & K
End Function

Function PubMthn$(Ln)
If IsPubMdy(Mdy(Ln)) Then PubMthn = Mthn(Ln)
End Function

Function CMthn$() ' current method name
Dim M As CodeModule: Set M = CMd
If IsNothing(M) Then Exit Function
CMthn = CMthnzM(CMd)
End Function

Function Mthn$(Ln)
Dim L$: L = RmvMdy(Ln)
If ShfMthTy(L) = "" Then Exit Function
Mthn = TakNm(L)
End Function

Private Sub MthDnzL__Tst()
Debug.Print MthDnzL("Function MthnzMthDn$(MthDn$)")
End Sub

Function PubMthKnzL$(L) ' Method-key-name from a line
Dim A As Mthn3: A = Mthn3zL(L)
If A.ShtMdy <> "" Then Exit Function
PubMthKnzL = MthKnz3(A)
End Function

Function MthKnzL$(L) ' Method-key-name from a line
MthKnzL = MthKnz3(Mthn3zL(L))
End Function

Function MthKnz3$(A As Mthn3)
If A.Nm = "" Then Exit Function
MthKnz3 = MthKn(A.Nm, PrpTy(A.ShtTy))
End Function

Function PrpTy$(ShtMthTy$) ' Ret ShtMthTy if they are Get|Set|Let else Blank
Select Case ShtMthTy
Case "Get", "Let", "Set": PrpTy = ShtMthTy
End Select
End Function

Function MthKn$(Nm$, PrpTy$)
MthKn = JnDotApNB(Nm, PrpTy)
End Function

Function MthMkny() As String() 'Method-module-key-name-array
MthMkny = MthMknyzP(CPj)
End Function

Function Mth4ny() As String() 'Method-module-key-name-array
Mth4ny = Mth4nyP
End Function

Function Mth4nyP() As String() 'Method-module-key-name-array
Mth4nyP = Mth4nyzP(CPj)
End Function

Function Mth4nyzP(P As VBProject) As String()  'Method-module-key-name-array
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy Mth4nyzP, Mth4nyzM(C.CodeModule)
Next
End Function

Function Mth4nyzM(M As CodeModule) As String()
Mth4nyzM = Mth4nyzMS(Mdn(M), Src(M))
End Function

Function Mth4nyzMS(Mdn$, Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB Mth4nyzMS, Mth4nzML(Mdn, L)
Next
End Function

Function Mth4nzML$(Mdn$, L)
Dim N$: N = Mth3nzL(L): If N = "" Then Exit Function
Mth4nzML = Mdn & "." & N
End Function

Function Mth3nzL$(L)
Dim A As Mthn3: A = Mthn3zL(L)
With A
If .Nm = "" Then Exit Function
If .Nm = "Str" Then Stop
Mth3nzL = JnDotApNB(.Nm, .ShtTy, .ShtMdy)
End With
End Function

Function MthMknyzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy MthMknyzP, MthMknyzM(C.CodeModule)
Next
End Function

Function MthMknyM() As String()
MthMknyM = MthMknyzM(CMd)
End Function

Function MthMknyzM(M As CodeModule) As String()
MthMknyzM = AmAddPfx(MthKny(Src(M)), Mdn(M) & ".")
End Function

Function MthKnyzM(M As CodeModule) As String()
MthKnyzM = MthKny(Src(M))
End Function

Function MthKny(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB MthKny, MthKnzL(L)
Next
End Function

Function MthMkn$(Mdn$, Mthn$, PrpTy$)
':MthMkn: :DotNm ! #Mth-(Md-Key-Nm)#
MthMkn = JnDotApNB(Mdn, Mthn, PrpTy)
End Function

Function MthDn$(Nm$, ShtTy$, ShtMdy$)
MthDn = JnDotApNB(Nm, ShtTy, ShtMdy)
End Function

Function MthDnz3$(A As Mthn3)
MthDnz3 = MthDn(A.Nm, A.ShtMdy, A.ShtTy)
End Function

Function MthDnzL$(Ln)
MthDnzL = MthDnz3(Mthn3zL(Ln))
End Function

Function MthnzL(Ln)
MthnzL = Mthn(Ln)
End Function

Function Prpn$(Ln)
Dim L$: L = RmvMdy(Ln)
If ShfKd(L) <> "Property" Then Exit Function
Prpn = TakNm(L)
End Function

Private Sub Mthn__Tst()
GoTo Z
Dim A$
A = "Function Mthn(A)": Ept = "Mthn.Fun.": GoSub Tst
Exit Sub
Tst:
    Act = Mthn(A)
    C
    Return
Z:
    Dim O$(), L
    For Each L In SrczV(CVbe)
        PushNB O, Mthn(CStr(L))
    Next
    Brw O
End Sub


Function HasPubMth(Src$(), PubMthn) As Boolean
Dim L: For Each L In Itr(Src)
    If MxIdeMthn.PubMthn(L) = PubMthn Then HasPubMth = True
Next
End Function

Private Sub WiVerbMthnAetP__Tst()
BrwAet WiVerbMthnAetP.Srt
End Sub

Sub PushNDupDy(ODy(), Dr)
If HasDr(ODy, Dr) Then Exit Sub
PushI ODy, Dr
End Sub

Function WiVerbMthnyP() As String()
Dim J&
Dim Mthn: For Each Mthn In Itr(MthnyV)
    If J Mod 100 = 0 Then Debug.Print J
    If HasVerb(Mthn) Then PushI WiVerbMthnyP, Mthn
    J = J + 1
Next
End Function

Function WiVerbMthnAetP() As Dictionary
Set WiVerbMthnAetP = Aet(WiVerbMthnyP)
End Function

Function WoVerbMthnAetP() As Dictionary
Set WoVerbMthnAetP = Aet(WoVerbMthnyP)
End Function

Function MthnyzSrcIxy(Src$(), Mthixy&()) As String()
Dim Ix: For Each Ix In Itr(Mthixy)
    PushI MthnyzSrcIxy, Mthn(Src(Ix))
Next
End Function

Function MthnyV() As String()
MthnyV = MthnyzV(CVbe)
End Function

Function MthnAetP() As Dictionary
Set MthnAetP = Aet(MthnyP)
End Function

Function MthnyP() As String()
MthnyP = MthnyzP(CPj)
End Function

Function TstMthnyP() As String()
TstMthnyP = TstMthnyzP(CPj)
End Function
Function TstMthnyzP(P As VBProject) As String()
TstMthnyzP = TstMthny(MthnyzP(P))
End Function
Function TstMthny(Mthny$()) As String()
TstMthny = AwSfx(Mthny, "__Tst")
End Function

Function MthnyzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy MthnyzP, MthnyzM(C.CodeModule)
Next
End Function

Function MthnyzFb(Fb) As String()
MthnyzFb = MthnyzV(VbezPjf(Fb))
End Function

Sub MthnyzFb__Tst()
GoSub X_BrwAll
Exit Sub
X_BrwAll:
    Dim O$(), Fb
    For Each Fb In MhdFbaAy
        PushAy O, MthnyzFb(Fb)
    Next
    Brw O
    Return
X_BrwOne:
    Dim A$(): A = MhdFbaAy
    Brw MthnyzFb(A(0))
    Return
End Sub

Function Mthny(Src$()) As String()
Dim L: For Each L In Itr(RmvFalseSrc(Src))
    PushNB Mthny, Mthn(L)
Next
End Function

Function MthnyzM(M As CodeModule) As String()
MthnyzM = Mthny(Src(M))
End Function

Private Sub Mthny__Tst()
GoSub Z
Exit Sub
Z:
   B Mthny(SrczP(CPj))
   Return
End Sub

Function MthnyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy MthnyzV, MthnyzP(P)
Next
End Function

Function MthnAetV() As Dictionary
Set MthnAetV = Aet(MthnyV)
End Function

Function MthnyM() As String()
MthnyM = MthnyzM(CMd)
End Function

Function HasMth(Src$(), Mthn) As Boolean
HasMth = Mthix(Src, Mthn) >= 0
End Function

Function HasMthzM(M As CodeModule, Mthn) As Boolean
HasMthzM = HasMth(Src(M), Mthn)
End Function

Function MthnCmlAetV() As Dictionary
Set MthnCmlAetV = CmlAetzNy(MthnyV)
End Function

Sub BrwMthnP()
BrwDrs SrtDrs(MthnDrsP)
End Sub

Function PrpNyzCmp(A As VBComponent) As String()
PrpNyzCmp = Itn(A.Properties)
End Function

Function MthRetTy$(Ln)
Dim A$: A = AftBkt(Ln)
If ShfTermX(A, "As") Then MthRetTy = T1(A)
End Function

Function MthQnzMthn$(Mthn)
Const CSub$ = CMod & "MthQnzMthn"
Dim D As Drs: D = DwEq(MthDrsP, "Mthn", Mthn)
Select Case Si(D.Dy)
Case 0: InfLn CSub, "No such Mthn[" & Mthn & "]"
Case 1:
    Dim IxMdn%: IxMdn = IxzAy(D.Fny, "Mdn")
    MthQnzMthn = D.Dy(0)(IxMdn) & "." & Mthn
Case Else
    InfLn CSub, "No then one Md has Mthn[" & Mthn & "]"
    IxMdn = IxzAy(D.Fny, "Mdn")
    Dim Dr: For Each Dr In D.Dy
        Debug.Print Dr(IxMdn) & "." & Mthn
    Next
End Select
End Function

Function CMthnzM$(M As CodeModule)
Dim K As vbext_ProcKind
CMthnzM = M.ProcOfLine(CLnozM(M), K)
End Function

Function HasMthnzMNT(M As CodeModule, Nm, Optional ShtMthTy$) As Boolean
HasMthnzMNT = HasMthnzSNT(Src(M), Nm, ShtMthTy)
End Function

Function HasMthnzSNT(Src$(), Nm, Optional ShtMthTy$) As Boolean
Dim L: For Each L In Itr(Src)
    With Mthn3zL(L)
        If .Nm = Nm Then
            If HitOptEq(.ShtTy, ShtMthTy) Then
                HasMthnzSNT = True
                Exit Function
            End If
            Debug.Print FmtQQ("HasMthn: Ln has Mthn[?] but not hit given ShtMthTy[?].  Act ShtMthTy=[?]", Nm, ShtMthTy, .ShtTy)
        End If
    End With
Next
End Function
