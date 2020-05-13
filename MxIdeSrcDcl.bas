Attribute VB_Name = "MxIdeSrcDcl"
Option Compare Text
Option Explicit
Const CNs$ = "Ide.Dcl"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcDcl."
Public Const DclItmss$ = "Option Const Type Enum Dim "

Function HasDclCdl(M As CodeModule, DclCdl$) As Boolean
HasDclCdl = HasSubStr(Dcll(M), DclCdl)
End Function

Function DclItmAyzDimLin(DimLin) As String()
Dim L$: L = DimLin
If Not ShfPfx(L, "Dim ") Then Exit Function
DclItmAyzDimLin = SplitCommaSpc(L)
End Function

Function DclNy(DcitmAy$()) As String()
Dim Dcitm: For Each Dcitm In Itr(DcitmAy)
    PushI DclNy, Dcn(Dcitm)
Next
End Function

Function CdLyzL(Ln) As String()
Dim L$: L = Trim(Ln)
If L = "" Then Exit Function
If FstChr(L) = "'" Then Exit Function
CdLyzL = SyzTrim(Split(Ln, ":"))
End Function

Private Sub ClnSrc__Tst(): Brw ClnSrc(SrczP(CPj)): End Sub
Function IxyzPrvCdLn&(Src$(), Fm)
Dim O&: For O = Fm - 1 To 0 Step -1
    If IsCdLn(Src(O)) Then IxyzPrvCdLn = O: Exit Function
Next
IxyzPrvCdLn = -1
End Function
Function DclDicP() As Dictionary
Set DclDicP = DclDiczP(CPj)
End Function

Function DclDiczP(P As VBProject) As Dictionary
If P.Protection = vbext_pp_locked Then Set DclDiczP = New Dictionary: Exit Function
Dim C As VBComponent, M As CodeModule
Set DclDiczP = New Dictionary
For Each C In P.VBComponents
    Set M = C.CodeModule
    Dim D$(): D = Dcl(M)
    If Si(D) = 0 Then
        DclDiczP.Add MdDn(M), D
    End If
Next
End Function

Function DclP() As String()
DclP = DclzP(CPj)
End Function

Function DclzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy DclzP, Dcl(C.CodeModule)
Next
End Function

Function DclItr(M As CodeModule)
':DclItr: :LinItr #Dcl-Ln-Itr#
Asg Itr(Dcl(M)), DclItr
End Function

Function DcllzSrc$(Src$())
':Dcll: :Lines ! comes fm a module
DcllzSrc = JnCrLf(DclzSrc(Src))
End Function

Function HasDclLin(M As CodeModule, DclLin$) As Boolean
Dim J&: For J = 1 To M.CountOfDeclarationLines
    If M.Lines(J, 1) = DclLin Then HasDclLin = True: Exit Function
Next
End Function

Function DclzSrc(Src$()) As String()
If Si(Src) = 0 Then Exit Function
Dim N&: N = NDclLinzS(Src)
If N <= 0 Then Exit Function
DclzSrc = FstNEle(Src, N)
End Function

Function MdlzN$(P As VBProject, Mdn)
MdlzN = Mdl(MdzP(P, Mdn))
End Function

Function MdlAy(P As VBProject, Mdny$()) As String()
Dim N: For Each N In Itr(Mdny)
    PushI MdlAy, MdlzN(P, N)
Next
End Function

Function Mdl$(M As CodeModule)
Dim Cnt&: Cnt = M.CountOfLines
If Cnt = 0 Then Exit Function
Mdl = M.Lines(1, Cnt)
End Function

Function MdlzLc$(M As CodeModule, Lc As Lcnt)
With Lc
If Lc.Lno <= 0 Then Exit Function
If Lc.Cnt <= 0 Then Exit Function
If Lc.Lno > M.CountOfLines Then Exit Function
MdlzLc = M.Lines(Lc.Lno, Lc.Cnt)
End With
End Function

'**Dcl
Private Sub Dcl__Tst()
Dim M As CodeModule
'GoSub ZZ
GoSub T1
Exit Sub
T1:
    Set M = Md("MxVbRunThw")
    Ept = ""
    GoTo Tst
Tst:
    Act = Dcl(M)
    C
    Return
ZZ:
    Dim O$(), Cmp As VBComponent
    For Each Cmp In CPj.VBComponents
        PushNB O, Dcl(Cmp.CodeModule)
    Next
    VcLinesy O
End Sub
Function DclM() As String()
DclM = Dcl(CMd)
End Function

'**Dcll
Function DcllP()
Dim O$()
Dim C As VBComponent: For Each C In CPj.VBComponents
    PushI O, Dcll(C.CodeModule)
Next
DcllP = JnCrLf(O)
End Function
Function CDcll$():                       CDcll = JnCrLf(CDcl):       End Function
Function Dcll$(M As CodeModule)
Dim N%: N = NDcll(M)
If N > 0 Then Dcll = M.Lines(1, N)
End Function

'**NDcll
Private Sub NDcll__Tst()
Dim O$()
Dim C As VBComponent: For Each C In CPj.VBComponents
    PushI O, NDcll(C.CodeModule) & " " & C.Name
Next
BrwAy O
End Sub
Function NDcll%(M As CodeModule)
Dim S$(): S = Src(M)
NDcll = Mrmkix(S, FstMthix(S, M.CountOfDeclarationLines + 1)) - 1
End Function

'**Dcl
Function CDcl() As String():              CDcl = Dcl(CMd):           End Function
Function Dcl(M As CodeModule) As String(): Dcl = SplitCrLf(Dcll(M)): End Function

Function DclAy(P As VBProject) As Variant()
Dim C As VBComponent: For Each C In P.VBComponents
    PushI DclAy, Dcl(C.CodeModule)
Next
End Function
