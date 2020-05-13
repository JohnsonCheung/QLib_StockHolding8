Attribute VB_Name = "MxIdeSrc"
Option Explicit
Option Compare Text
Const CNs$ = "Src"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrc."
#If Doc Then
':IndtSrc:
#End If
Function SrcV() As String()
SrcV = SrczV(CVbe)
End Function
#If False Then
'AA
#End If

Private Sub RmvFalseSrc__Tst(): Brw RmvFalseSrc(SrcM): End Sub
Function RmvFalseSrc(Src$()) As String()
Dim InFalse As Boolean, IsFalseLn As Boolean, IsEndIfLn As Boolean
Dim L: For Each L In Itr(Src)
    IsFalseLn = L = "#If False Then"
    IsEndIfLn = L = "#End If"
    Select Case True
    Case IsFalseLn And InFalse:   Thw CSub, "Impossible to have InFalse=True and IsFalseLn=true"
    Case IsFalseLn:               InFalse = True
    Case IsEndIfLn And InFalse:   InFalse = False
    Case InFalse:
    Case Else:                    PushI RmvFalseSrc, L
    End Select
Next
End Function
Function Src(M As CodeModule) As String()
Src = SplitCrLf(Srcl(M))
End Function

Function SrcM() As String()
SrcM = SplitCrLf(Srcl(CMd))
End Function

Function SrczM(M As CodeModule) As String()
SrczM = SplitCrLf(Srcl(M))
End Function

Function Srcl$(M As CodeModule) '#Src-Lines# :Lines
If M.CountOfLines > 0 Then Srcl = M.Lines(1, M.CountOfLines)
End Function

Function SrczP(P As VBProject) As String()
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
For Each C In P.VBComponents
    PushIAy SrczP, Src(C.CodeModule)
Next
End Function

Function SrczV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy SrczV, SrczP(P)
Next
End Function

Function SrclM$()
SrclM = Srcl(CMd)
End Function

Function SrclP$()
SrclP = SrclzP(CPj)
End Function

Function SrcP() As String()
SrcP = SrczP(CPj)
End Function

Function SrclzP$(P As VBProject)
SrclzP = JnCrLf(SrczP(P))
End Function

Function VbItmEix&(Src$(), Bix, Itmn$)
If Bix < 0 Then VbItmEix = -1: Exit Function
Dim N$: N = "End " & Itmn
If HasSubStr(Src(Bix), N) Then VbItmEix = Bix: Exit Function
Dim O%: For O = Bix To UB(Src)
    If HasPfx(LTrim(Src(O)), N) Then VbItmEix = O: Exit Function
Next
VbItmEix = -1
End Function

Function ShfLHS$(OLin$)
Dim L$:                   L = OLin
Dim IsSet As Boolean: IsSet = ShfTermX(L, "Set")
Dim S$:                       If IsSet Then S = "Set "
Dim LHS$:               LHS = ShfDotNm(L)
If FstChr(L) = "(" Then
    LHS = LHS & QuoBkt(BetBkt(L))
    L = AftBkt(L)
End If
If Not ShfPfx(L, " = ") Then Exit Function
ShfLHS = S & LHS & " = "
OLin = L
End Function

Function ShfLRHS(OLin$) As Variant()
Dim L$:     L = OLin
Dim LHS$: LHS = ShfLHS(L)
With Brk1(L, "'")
    Dim RHS$:  RHS = .S1
              OLin = "   ' " & .S2
End With
ShfLRHS = Array(LHS, RHS)
End Function

'**IndtSrc
Sub IndtSrcP__Tst(): VcAy IndtSrcP: End Sub
Function IndtSrcP() As String(): IndtSrcP = IndtSrczP(CPj): End Function
Function IndtSrczP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy IndtSrczP, IndtSrczM(C.CodeModule)
Next
End Function
Function IndtSrczM(M As CodeModule) As String()
PushI IndtSrczM, Mdn(M)
PushIAy IndtSrczM, AmAddPfx(Src(M), " ")
End Function


