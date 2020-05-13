Attribute VB_Name = "MxVbStrSeed"
Option Explicit
Option Compare Text
Const CNs$ = "Vb.Dta"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxVbStrSeed."
Public Const EmEpPjKd$ = "EiPk *Fba *Fxa"

Enum EmPjKd
EiPkFba
EiPkFxa
End Enum
Function Expandss$(ExpandStr$)
Expandss = JnSpc(Expand(ExpandStr))
End Function

Function Expand(ExpandStr$) As String()
Dim S$, Sy$(), L$
L = ExpandStr
S = ShfTerm(L)
Dim I: For Each I In SyzSS(L)
    PushI Expand, Replace(I, "*", S)
Next
End Function

Function EmSyPjKd() As String()
EmSyPjKd = Expand(EmEpPjKd)
End Function

Function EmssPjKd$()
EmssPjKd = Expandss(EmEpPjKd)
End Function

Function DistPthiP$()
DistPthiP = DistPthi(SrcPthP)
End Function

Function DistPthi$(SrcPth$)
':DistPthi: :Pthi ! #Dist-Pthi#
DistPthi = Pthi(DistPth(SrcPth))
End Function

Function DistPth$(SrcPth$)
ChkIsSrcPth SrcPth, "DistPth"
DistPth = EnsPth(RmvExt(Fdr(ParPth(SrcPth))) & ".dist")
End Function

Sub ChkIsSrcPth(Pth$, Fun$)
If Not IsSrcPth(Pth) Then Thw Fun, "Given @Pth is not a Src path.  (A SrcPth should under a fdr of name [.src])", "Given-Pth", Pth
End Sub

Function DistFbaiP$()
DistFbaiP = DistFbai(SrcPthP)
End Function

Function DistiFbazP$(P As VBProject)
DistiFbazP = DistFbai(SrcPth(Pjf(P)))
End Function

Function DistPthP$() 'Distribution Path
DistPthP = DistPth(SrcPthP)
End Function

Function DistFbai$(SrcPth$)
DistFbai = DistPthi(SrcPth) & DistFn(SrcPth, EiPkFba)
End Function

Function DistFn$(SrcPth$, Kd As EmPjKd)
DistFn = RmvExt(PjfzSrcPth(SrcPth)) & ExtzPjKd(Kd)
End Function

Function PjfzSrcPth$(SrcPth$)
PjfzSrcPth = RmvFstChr(UpNFdr(SrcPth, 2))
End Function

Function ExtzPjKd$(Kd As EmPjKd)
Dim O$
Select Case True
Case Kd = EiPkFba: O = ".accdb"
Case Kd = EiPkFxa: O = ".xlsa"
Case Else: EnmEr "ExtzPjKd", "EmPjKd", EmssPjKd, CInt(Kd)
End Select
End Function

Function DistFxai$(SrcPth$)
DistFxai = DistPthi(SrcPth) & DistFn(SrcPth, EiPkFxa)
End Function

Function DistFxaiP$()
DistFxaiP = DistFxai(SrcPthP)
End Function

Function DistFxaizP$(P As VBProject)
DistFxaizP = DistFxai(SrcPthzP(P))
End Function

Sub DistFxai__Tst()
Dim SrcPth1$
GoSub T0
Exit Sub
T0:
    SrcPth1 = SrcPthP
    Ept = "C:\Users\user\Documents\Projects\Vba\QLib\.Dist\QLib(002).xlam"
    GoTo Tst
Tst:
    Act = DistFxai(SrcPth1)
    C
    Return
End Sub
