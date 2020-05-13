Attribute VB_Name = "MxIdePj"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdePj."

Function CvPj(I) As VBProject
Set CvPj = I
End Function

Function IsPjn(A) As Boolean
IsPjn = HasEle(PjnyV, A)
End Function

Function Pj(Pjn) As VBProject
Set Pj = CVbe.VBProjects(Pjn)
End Function

Function PjfStrP$()
PjfStrP = LineszFt(PjfP)
End Function

Function Pjp$(P As VBProject)
':Pjp: :Pth #Pj-Pth#
Pjp = Pth(Pjf(P))
End Function

Function PjfnP$()
PjfnP = Pjfn(CPj)
End Function

Function Pjfn$(P As VBProject)
Pjfn = Fn(Pjf(P))
End Function

Function PjfnAyV() As String()
PjfnAyV = PjfnAyzV(CVbe)
End Function

Function PjfnAyzV(A As Vbe) As String()
PjfnAyzV = FnAyzFfnAy(PjfyzV(A))
End Function

Function PjfzM$(M As CodeModule)
PjfzM = Pjf(PjzM(M))
End Function

Function Pjf$(P As VBProject)
Pjf = PjfzP(P)
End Function

Function PjfzP$(P As VBProject)
On Error GoTo X
PjfzP = P.FileName
Exit Function
X: Debug.Print FmtQQ("Cannot get Pjf for Pj(?). Err[?]", P.Name, Err.Description)
End Function

Function PjFnn$(P As VBProject)
PjFnn = Fnn(Pjf(P))
End Function

Function MdzP(P As VBProject, Mdn) As CodeModule
Set MdzP = P.VBComponents(Mdn).CodeModule
End Function

Sub ActPj(P As VBProject)
Set P.Collection.Vbe.ActiveVBProject = P
End Sub

Function IsProtectvvInf(P As VBProject) As Boolean
Const CSub$ = CMod & "IsProtectvvInf"
If Not IsProtect(P) Then Exit Function
InfLn CSub, FmtQQ("Skip protected Pj{?)", P.Name)
IsProtectvvInf = True
End Function

Function IsProtect(P As VBProject) As Boolean
IsProtect = P.Protection = vbext_pp_locked
End Function

Function FstMd(P As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In P.VBComponents
    If IsMd(CvCmp(Cmp)) Then
        Set FstMd = Cmp.CodeModule
        Exit Function
    End If
Next
End Function

Function FstMod(P As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In P.VBComponents
    If IsMod(Cmp) Then
        Set FstMod = Cmp.CodeModule
        Exit Function
    End If
Next
End Function

Function IsFbaPj(P As VBProject) As Boolean
IsFbaPj = IsFba(Pjf(P))
End Function
