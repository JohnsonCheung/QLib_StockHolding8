Attribute VB_Name = "MxIdeMdOpRmk"
Option Explicit
Option Compare Text
Const CNs$ = "Md"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdRmk."

Function UnRmkMd(M As CodeModule) As Boolean
Debug.Print "UnRmk " & M.Parent.Name,
If Not IsRmkdMd(M) Then
    Debug.Print "No need"
    Exit Function
End If
Debug.Print "<===== is unmarked"
Dim J%, L$
For J = 1 To M.CountOfLines
    L = M.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    M.ReplaceLine J, Mid(L, 2)
Next
UnRmkMd = True
End Function

Function IsRmkdSrc(Src$()) As Boolean
Dim L: For Each L In Itr(Src)
    If Not IsRmkln(L) Then Exit Function
Next
IsRmkdSrc = True
End Function

Function RmkdMdNyP() As String()
RmkdMdNyP = RmkdMdNy(CPj)
End Function

Function RmkdMdNy(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If IsRmkdMd(C.CodeModule) Then PushI RmkdMdNy, C.Name
Next
End Function
Function IsRmkdM() As Boolean
IsRmkdM = IsRmkdMd(CMd)
End Function

Function IsRmkdMd(M As CodeModule) As Boolean
Dim J%, L$
For J = 1 To M.CountOfLines
    If Left(M.Lines(J, 1), 1) <> "'" Then Exit Function
Next
IsRmkdMd = True
End Function

Sub RmkM()
RmkMd CMd
End Sub

Sub RmkAllMd()
Dim I, Md As CodeModule
Dim NRmk%, Skip%
For Each I In CPj.VBComponents
    If Md.Name <> "LibIdeRmkMd" Then
        If RmkMd(CvMd(I)) Then
            NRmk = NRmk + 1
        Else
            Skip = Skip + 1
        End If
    End If
Next
Debug.Print "NRmk"; NRmk
Debug.Print "SKip"; Skip
End Sub

Function RmkMd(M As CodeModule) As Boolean
Debug.Print "Rmk " & M.Parent.Name,
If IsRmkdMd(M) Then
    Debug.Print " No need"
    Exit Function
End If
Debug.Print "<============= is remarked"
Dim J%
For J = 1 To M.CountOfLines
    M.ReplaceLine J, "'" & M.Lines(J, 1)
Next
RmkMd = True
End Function

Sub UnRmk()
UnRmkMd CMd
End Sub

Sub UnRmkAll()
Dim NUnRmk%, NSkip%
Dim C As VBComponent: For Each C In CPj.VBComponents
    If UnRmkMd(C.CodeModule) Then
        NUnRmk = NUnRmk + 1
    Else
        NSkip = NSkip + 1
    End If
Next
Inf CSub, "NUnRmk", NUnRmk
Inf CSub, "NSkip", NSkip
End Sub

Sub UnRmkByLcntAy(M As CodeModule, A() As Lcnt)
Dim J&: For J = 0 To LcntUB(A)
    UnRmkByLCnt M, A(J)
Next
End Sub
Sub ChkMdLcntRmkd(M As CodeModule, A As Lcnt)
Stop
End Sub
Sub UnRmkByLCnt(M As CodeModule, A As Lcnt)
ChkMdLcntRmkd M, A
Dim J%: For J = A.Lno To A.Lno + A.Cnt - 1
    Dim L$: L = M.Lines(J, 1)
    M.ReplaceLine J, Mid(L, 2)
Next
End Sub

Sub RmkLn(M As CodeModule, Lno)
M.ReplaceLine Lno, "'" & M.Lines(Lno, 1)
End Sub
Sub RmkByLcnt(M As CodeModule, A As Lcnt)
Dim J&: For J = A.Lno To A.Lno + A.Cnt - 1
    RmkLn M, J
Next
End Sub
