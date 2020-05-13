Attribute VB_Name = "MxIdeMdOpRmk1"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdRmk1."
'--MthRmk
Function IsMrmk(Src) As Boolean
If IsEmpAy(Src) Then Exit Function
Dim L: For Each L In Src
    If Not IsRmkln(L) Then Exit Function
Next
IsMrmk = Las2Chr(LasEle(Src))
End Function

Private Sub MrmkDiP__Tst(): BrwDic MrmkDiP: End Sub
Private Sub MrmkDiz__Tst(): BrwDic MrmkDizM(Md("MxIdeMrmk")): End Sub
Function MrmkDiP() As Dictionary: Set MrmkDiP = MrmkDizP(CPj): End Function

Function MrmkDizP(P As VBProject) As Dictionary
Dim O As New Dictionary
Dim C As VBComponent: For Each C In P.VBComponents
    PushDic O, AddKpfx(MrmkDizM(C.CodeModule), C.Name & ".")
Next
Set MrmkDizP = O
End Function

Function MrmkDizM(M As CodeModule) As Dictionary
Set MrmkDizM = MrmkDi(Src(M))
End Function

Function MrmkDi(Src$()) As Dictionary
Dim L$
Dim I: For Each I In Mthixy(Src)
    L = Mrmkl(Src, I)
    If L <> "" Then MrmkDi.Add MthKnzL(L), L
Next
End Function


Function RmkBlkzM(M As CodeModule, RLno&) As String()
If RLno = 0 Then Exit Function
Dim J&, L$, O$()
For J = RLno To M.CountOfLines
    L = M.Lines(J, 1)
    If Not IsVrmkLn(L) Then Exit For
    PushI O, L
Next
RmkBlkzM = O
End Function
