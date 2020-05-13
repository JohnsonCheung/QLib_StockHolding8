Attribute VB_Name = "MxIdeMdEmp"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdEmp."

Function IsEmpMd(M As CodeModule) As Boolean
If M.CountOfLines = 0 Then IsEmpMd = True: Exit Function
Dim J&: For J = 1 To M.CountOfLines
    If Not IsNonSrcLn(M.Lines(J, 1)) Then Exit Function
Next
IsEmpMd = True
End Function

Sub RmvEmpCMd()
RmvEmpMd CPj
End Sub
Sub RmvEmpMd(P As VBProject)
Dim N: For Each N In Itr(EmpMdNy(P))
    RmvCmp P.VBComponents(N)
Next
End Sub

Function EmpCMdNy() As String()
EmpCMdNy = EmpMdNy(CPj)
End Function

Function EmpMdNyP() As String()
EmpMdNyP = EmpMdNy(CPj)
End Function

Function EmpMdNy(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If IsEmpMd(C.CodeModule) Then
        PushI EmpMdNy, C.Name
    End If
Next
End Function

Private Sub IsEmpMd__Tst()
Dim M As CodeModule
'GoSub T1
'GoSub T2
GoSub T3
Exit Sub
T3:
    Debug.Assert IsEmpMd(Md("Dic"))
    Return
T2:
    Set M = Md("Module2")
    Ept = True
    GoTo Tst
T1:
    '
    Dim T$, P As VBProject
        Set P = CPj
        T = TmpNm
    '
'    Set M = PjAddMd(P, T)
    Ept = True
    GoSub Tst
    DltCmpzPjn P, T
    Return
Tst:
    Act = IsEmpMd(M)
    C
    Return
End Sub

Function IsSrcEmp(A$()) As Boolean
Dim L
For Each L In Itr(A)
    If Not IsNonSrcLn(L) Then Exit Function
Next
IsSrcEmp = True
End Function

Function EmpMdNyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy EmpMdNyzV, EmpMdNy(P)
Next
End Function

Function NoMthMdNyP() As String()
NoMthMdNyP = NoMthMdNyzP(CPj)
End Function

Function NoMthMdNyzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If IsNoMthMd(C.CodeModule) Then PushI NoMthMdNyzP, C.Name
Next
End Function

Function IsNoMthMd(M As CodeModule) As Boolean
Dim J&: For J = M.CountOfDeclarationLines + 1 To M.CountOfLines
    If IsMthln(M.Lines(J, 1)) Then Exit Function
Next
IsNoMthMd = True
End Function
