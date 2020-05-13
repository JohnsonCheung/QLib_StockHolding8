Attribute VB_Name = "MxIdeMdInpCLibOrNs"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CNs$ = "Md3Cnst"
Const CMod$ = CLib & "MxIdeMdInpCLibOrNs."
Sub InpCLibP()
InpCLibzP CPj
End Sub

Private Sub InpCLibzP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    If InpCLibzM(C.CodeModule) Then
        Stop
        Exit Sub
    End If
Next
End Sub

Sub InpCLibM()
InpCLibzM CMd
End Sub

Private Function InpCLibzM(M As CodeModule) As Boolean
If CLibv(Dcl(M)) <> "" Then Exit Function
If Not IsMd(M.Parent) Then Exit Function
Static LasCModv$
Dim V$::
Again:
    V = InputBox("Input CLibv: (FstChr must be [Q])", "For Md: " & M.Name, LasCModv)
    If V = "" Then InpCLibzM = True: Exit Function
    If FstChr(V) <> "Q" Then
        MsgBox "FstChr must be [Q]", vbCritical
        GoTo Again
    End If
EnsCnst M, CLibLin(V)
LasCModv = V
End Function

Sub InpCNs()
InpCNszP CPj
End Sub

Private Sub InpCNszP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    If InpCNszM(C.CodeModule) Then Exit Sub
Next
End Sub

Sub InpCNsM()
InpCNszM CMd
End Sub

Private Function InpCNszM(M As CodeModule) As Boolean
If CNsv(Dcl(M)) <> "" Then Exit Function
Static LasCNsv$
Dim V$
    V = InputBox("Input CNsv: ", "For Md: " & M.Name, LasCNsv)
    If V = "" Then InpCNszM = True: Exit Function
EnsCnst M, CNsLin(V)
LasCNsv = V
End Function
