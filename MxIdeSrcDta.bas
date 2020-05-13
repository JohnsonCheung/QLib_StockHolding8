Attribute VB_Name = "MxIdeSrcDta"
Option Compare Text
Option Explicit
Const CNs$ = "Src.Dta"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcDta."
Type MdSrc
    Mdn As String
    Src() As String
End Type
Type PjSrc
    Pjn As String
    Md() As MdSrc
End Type
Function PjSrcP() As PjSrc
PjSrcP = PjSrczP(CPj)
End Function

Function PjSrczP(P As VBProject) As PjSrc
PjSrczP.Pjn = P.Name
Dim C As VBComponent: For Each C In P.VBComponents
    PushMdSrc PjSrczP.Md, MdSrcFmMd(C.CodeModule)
Next
End Function

Sub PushMdSrc(O() As MdSrc, M As MdSrc)
Dim N&: N = MdSrcSi(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Function MdSrcSi&(A() As MdSrc)
On Error Resume Next
MdSrcSi = UBound(A) + 1
End Function

Function MdSrcUB&(A() As MdSrc)
MdSrcUB = MdSrcSi(A) - 1
End Function

Private Sub FmtPjSrc__Tst()
VcAy FmtPjSrc(PjSrczP(CPj)), "PjSrc_"
End Sub

Function FmtPjSrc(P As PjSrc) As String()
PushI FmtPjSrc, P.Pjn
Dim Ay() As MdSrc: Ay = P.Md
Dim J&: For J = 0 To MdSrcUB(Ay)
    PushIAy FmtPjSrc, AmAddPfxTab(FmtMdSrc(Ay(J)))
Next
End Function

Function FmtMdSrc(A As MdSrc) As String()
PushI FmtMdSrc, A.Mdn
PushIAy FmtMdSrc, AmAddPfxTab(A.Src)
End Function

Function MdSrcFmMd(M As CodeModule) As MdSrc
MdSrcFmMd.Mdn = Mdn(M)
MdSrcFmMd.Src = Src(M)
End Function
