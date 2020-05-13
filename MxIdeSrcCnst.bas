Attribute VB_Name = "MxIdeSrcCnst"
Option Explicit
Option Compare Text
Const CNs$ = "Cnst"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcCnst."
'Enum XoCnst
'    XiMdn
'    XiIsPrv
'    XiCnstn
'    XiTyChr
'    XiCnstv
'End Enum

Function CnstLyP() As String()
CnstLyP = CnstLyzP(CPj)
End Function

Function CnstLyzP(P As VBProject) As String()
CnstLyzP = CnstLy(SrczP(P))
End Function

Sub CnstNyP__Tst()
Vc CnstNyP
End Sub

Function CnstNyP() As String()
CnstNyP = CnstNy(SrczP(CPj))
End Function

Function CnstNy(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB CnstNy, Cnstn(L)
Next
End Function
Function Cnstn$(Ln)
Cnstn = NmAftTerm(RmvMdy(Ln), "Const")
End Function

Function CnstLno%(M As CodeModule, Cnstn$, Optional IsPrvOnly As Boolean)
CnstLno = CnstIx(Src(M), Cnstn, IsPrvOnly) + 1
End Function

Function CnstLLnzSrc(Src$(), CnstNm$) As LLn
Dim J&: For J = 0 To UB(Src)
    Dim L$: L = Src(J)
    If Cnstn(L) = CnstNm Then
        L = Contln(Src, J)
        CnstLLnzSrc = LLn(J + 1, L)
        Exit Function
    End If
Next
End Function

Function CnstLLn(M As CodeModule, Cnstn$) As LLn
CnstLLn = CnstLLnzSrc(Dcl(M), Cnstn)
End Function

Private Sub HasCnstn__Tst()
Debug.Assert HasCnstn(CMd, "CMod")
End Sub

Function HasCnstn(M As CodeModule, Cnstn$) As Boolean
HasCnstn = CnstLno(M, Cnstn) = 0
End Function

Function HasCnstnzLin(Ln, N$) As Boolean
HasCnstnzLin = Cnstn(Ln) = N
End Function

Function ShfTermCnst(OLin$) As Boolean
ShfTermCnst = ShfTermX(OLin, "Const")
End Function

Function ShfCnst(OLin$) As Boolean
ShfCnst = ShfTerm(OLin) = "Const"
End Function

Function IsCnstLinzPfx(L, CnstnPfx$) As Boolean
Dim Ln$: Ln = RmvMdy(L)
If Not ShfTermCnst(Ln) Then Exit Function
IsCnstLinzPfx = HasPfx(L, CnstnPfx)
End Function

Private Sub IsStrCnstLin__Tst()
Dim O$()
Dim L: For Each L In SrczP(CPj)
    If IsStrCnstLin(L) Then PushI O, L
Next
Brw O
End Sub

Function IsStrCnstLin(Ln) As Boolean
Dim L$: L = Ln
ShfMdy L
If Not ShfTermX(L, "Const") Then Exit Function
If ShfNm(L) = "" Then Exit Function
IsStrCnstLin = FstChr(L) = "$"
End Function

Function IsLnCnst(Ln) As Boolean
Dim L$: L = Ln
ShfMdy L
If Not ShfTermX(L, "Const") Then Exit Function
If ShfNm(L) = "" Then Exit Function
IsLnCnst = True
End Function

Function IsCnstNmLin(L, CnstNm$) As Boolean
IsCnstNmLin = Cnstn(L) = CnstNm
End Function

Function CnstIx&(Src$(), CnstNm, Optional IsPrvOnly As Boolean)
Dim O&
Dim L: For Each L In Itr(Src)
    If Cnstn(L) = CnstNm Then
        Select Case True
        Case IsPrvOnly And HasPfx(L, "Public "): CnstIx = -1
        Case Else:                              CnstIx = O
        End Select
        Exit Function
    End If
    O = O + 1
Next
CnstIx = -1
End Function

Function CnstLinAy(Src$()) As String()
Dim Ix&, L: For Each L In Itr(Src)
    If IsStrCnstLin(L) Then PushI CnstLinAy, Contln(Src, Ix)
    Ix = Ix + 1
Next
End Function

Function CnstLinAyP() As String()
CnstLinAyP = CnstLinAy(SrczP(CPj))
End Function

Private Sub CnstLy__Tst()
Brw CnstLy(SrczP(CPj))
End Sub

Function CnstLy(Src$()) As String()
Dim Ix&: For Ix = 0 To UB(Src)
    If IsLnCnst(Src(Ix)) Then PushI CnstLy, Contln(Src, Ix)
Next
End Function
