Attribute VB_Name = "MxIdeSrcInf"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcInf."
Function SrczMdn(Mdn) As String()
SrczMdn = Src(Md(Mdn))
End Function
Sub TopRmIxSrcFm__Tst()
ZZ:
Dim ODy()
    Dim Src$(): Src = SrczMdn("IdeSrcLin")
    Dim Dr(), Lx&
    Dim J%, IsMth$, RmkLx$, Ln, I
    For Each I In Src
        Ln = I
        IsMth = ""
        RmkLx = ""
        If IsMthln(Ln) Then
            IsMth = "*Mth"
            RmkLx = Mrmkix(Src, Lx)

        End If
        Dr = Array(IsMth, RmkLx, Ln)
        Push ODy, Dr
        Lx = Lx + 1
    Next
BrwDrs DrszFF("Mth RmkLx Ln", ODy)
End Sub

Function ZZSrc() As String()
ZZSrc = Src(Md("IdeSrc"))
End Function

Property Get ZZSrcln()
ZZSrcln = "Sub IsMthln()"
End Property

Sub AsgMthDr(MthDr, OMdy$, OTy$, ONm$, OPrm$, ORet$, OLinRmk$, OLines$, OMrmk$)
AsgAy MthDr, OMdy, OTy, ONm, OPrm, ORet, OLinRmk, OLines, OMrmk
End Sub


Private Sub VbExmLy__Tst()
Vc VbExmLy(SrczP(CPj))
End Sub

Function VbExmLy(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLnVbExmRmk(L) Then PushI VbExmLy, L
Next
End Function

Function IsLnVbExmRmk(Ln) As Boolean
':VbExmRmkLn: :Ln #Vb-Exclaimation-Rmk-Line# ! It is a rmk Ln fst-non-spc-chr is ['] and nxt is [!]
Dim L$: L = LTrim(Ln)
If Not ShfPfx(L, "'") Then Exit Function
L = LTrim(L)
If Not ShfPfx(L, "!") Then Exit Function
IsLnVbExmRmk = True
End Function

Function ExmRmkl$(VbExmRmk$())
Dim O$()
Dim L: For Each L In Itr(VbExmRmk)
    PushI O, ExmRmk(L)
Next
ExmRmkl = JnCrLf(O)
End Function

Function ExmRmk$(VbExmRmkLn)
Const CSub$ = CMod & "ExmRmk"
Dim L$: L = LTrim(VbExmRmkLn)
If Not ShfPfx(L, "'") Then Thw CSub, "Given VbExmRmkLn does not have Fst-Non-Spc being [']", "VbExmRmkLn", VbExmRmkLn
L = LTrim(L)
If Not ShfPfx(L, "!") Then Thw CSub, "Given VbExmRmkLn does not have Snd-Non-Spc being [!]", "VbExmRmkLn", VbExmRmkLn
L = LTrim(L)
ExmRmk = Trim(L)
End Function

Function SrcByLcnt(M As CodeModule, A As Lcnt) As String()
SrcByLcnt = SplitCrLf(SrclByLcnt(M, A))
End Function

Function SrclByLcnt$(M As CodeModule, A As Lcnt)
SrclByLcnt = M.Lines(A.Lno, A.Cnt)
End Function

Function SrcwSngDblQ(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If HasSngDblQ(L) Then PushI SrcwSngDblQ, L
Next
End Function

Function NSrcLnP&()
NSrcLnP = NSrcLnzP(CPj)
End Function

Function NSrcLnzP&(P As VBProject)
Dim O&, C As VBComponent
If P.Protection = vbext_pp_locked Then Exit Function
For Each C In P.VBComponents
    O = O + C.CodeModule.CountOfLines
Next
NSrcLnzP = O
End Function
