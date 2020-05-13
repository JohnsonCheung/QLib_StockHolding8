Attribute VB_Name = "MxIdeMthDrs"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CNs$ = "Do.Mth"
Const CMod$ = CLib & "MxIdeMthDrs."
 Public Const MthFF$ = MdnFF$ & " L Mdy Ty Mthn Mthln"
Public Const PubMthFF$ = "Pjn MdTy Mdn L E Ty Mthn Mthln"
Public Const MthcFF$ = "Pjn Ns MdTy Mdn NLin L Mdy Ty Mthn Mthln Mthl"


'--MthDrs
Private Sub MthDrsP__Tst(): BrwDrsR MthDrsP: End Sub
Private Sub MthDrsM__Tst(): BrwDrsR MthDrsM: End Sub

Function MthDrsP() As Drs: MthDrsP = MthDrszP(CPj): End Function
Function MthDrsM() As Drs: MthDrsM = MthDrszM(CMd): End Function
'--
Function MthDrszS(Src$(), MdnDr()) As Drs
MthDrszS = DrszFF(MthFF, ZZMthDy(Src, MdnDr))
End Function

Function MthDrszM(M As CodeModule) As Drs
MthDrszM = MthDrszS(Src(M), MdnDrzM(M))
End Function

Function MthDrszP(P As VBProject) As Drs
Dim ODy()
Dim C As VBComponent: For Each C In P.VBComponents
    PushI ODy, ZZMthDy(Src(C.CodeModule), MdnDrzM(C.CodeModule))
Next
MthDrszP = DrszFF(MthFF, ODy)
End Function

'--MthcDrs

Private Sub MthcDrsP__Tst(): BrwDrsN MthcDrsP: End Sub

Function MthcDrsP() As Drs: MthcDrsP = MthcDrszP(CPj): End Function
Function MthcDrsM() As Drs: MthcDrsM = MthcDrszM(CMd): End Function

Function MthcDrszP(P As VBProject) As Drs:  MthcDrszP = ZZZMthcDrs(MthDrszP(P)):  End Function
Function MthcDrszM(M As CodeModule) As Drs: MthcDrszM = ZZZMthcDrs(MthDrszM(M)): End Function
Function MthcDrszS(Src$(), MdnDr()) As Drs: MthcDrszS = ZZZMthcDrs(MthDrszS(Src, MdnDr)): End Function

Function MthcDrszFxa(Fxa$, Optional Xls As Excel.Application) As Drs
Dim A As Excel.Application: Set A = DftXls(Xls)
MthcDrszFxa = MthcDrszP(PjzFxa(Fxa))
If IsNothing(Xls) Then QuitXls Xls
End Function

'-- MthcDrs
Private Sub MthcDrszP__Tst()
BrwDrs MthcDrszP(CPj)
End Sub

Function MthcWsP() As Worksheet
Set MthcWsP = WszDrs(MthcDrsP)
End Function

Function MthcDrszPjf(Pjf) As Drs
Dim V As Vbe, App, P As VBProject, PjDte As Date
OpnPjf Pjf ' Either Excel.Application or Access.Application
Set V = VbezPjf(Pjf)
Set P = PjzPjf(V, Pjf)
Select Case True
Case IsFb(Pjf):  PjDte = PjDtezAcs(CvAcs(App))
Case IsFxa(Pjf): PjDte = DtezFfn(Pjf)
Case Else: Stop
End Select
MthcDrszPjf = AddCol(MthcDrszP(P), "PjDte", PjDte)
If IsFb(Pjf) Then
    CvAcs(App).CloseCurrentDatabase
End If
End Function

Function MthcDrszPjfy(Pjfy$()) As Drs
Dim F: For Each F In Pjfy
    AppDrs MthcDrszPjfy, MthcDrszPjf(F)
Next
End Function

Function MthcDrszV(V As Vbe) As Drs
Dim P As VBProject: For Each P In V.VBProjects
    Dim A As Drs: A = MthcDrszP(P)
    Dim O As Drs: O = AddDrs(O, A)
Next
MthcDrszV = O
End Function

Private Function ZZMthDy(Src$(), MdnDr()) As Variant()
Dim L, Ix&: For Each L In Itr(Src)
    If IsMthln(L) Then
        PushI ZZMthDy, W1MthDr(Ix, L, MdnDr)
    End If
    Ix = Ix + 1
Next
End Function

Private Function W1MthDr(Ix&, Mthln, MdnDr()) As Variant()
Dim A As Mthn3:      A = Mthn3zL(Mthln)
Dim Ty$:            Ty = A.ShtTy
Dim Mdy$:          Mdy = A.ShtMdy
Dim Mthn$:        Mthn = A.Nm
                W1MthDr = AddAyAp(MdnDr, Ix + 1, Mdy, Ty, Mthn, Mthln)
End Function

'==ZZZ: the core subroutines.  Usually they are private. IF not, make one by adding ZZZ.  Put them at end will most easy to access.  They should not many, 3 or less.  They can have W?.
Private Function ZZZMthcDrs(MthDrs As Drs) As Drs
Dim Dy()
    Dim IxL%, IxMdn%, IxPjn%: AsgIx MthDrs, "L Mdn Pjn", IxL, IxMdn, IxPjn
    Dim Dic As New Dictionary
    Dim Dr: For Each Dr In Itr(MthDrs.Dy)
        Dim Pjn$: Pjn = Dr(IxPjn)
        Dim Mdn$: Mdn = Dr(IxMdn)
        Dim Src$(): Src = W1Src(Dic, Pjn, Mdn)
        Dim Bix&: Bix = Dr(IxL) - 1
        Dim Eix&: Eix = MthEix(Src, Bix)
        Dim Mthl$: Mthl = JnCrLf(AwBE(Src, Bix, Eix))
        PushI Dr, Mthl
        PushI Dy, Dr
    Next
ZZZMthcDrs = AddColByNewDy(MthDrs, "Mthl", Dy)
End Function

Private Function W1Src(ODic As Dictionary, Pjn$, Mdn$) As String() ' Ret :Src from @Dic by key @Pjn.@Mdn or add one to @Dic if new
Dim K$: K = Pjn & "." & Mdn
If ODic.Exists(K) Then
    W1Src = ODic(K)
Else
    Dim M As CodeModule: Set M = Pj(Pjn).VBComponents(Mdn).CodeModule
    ODic.Add K, Src(M)
End If
End Function
