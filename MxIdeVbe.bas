Attribute VB_Name = "MxIdeVbe"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeVbe."
Function CvVbe(A) As Vbe
Set CvVbe = A
End Function

Function PjzV(A As Vbe, Pjn$) As VBProject
Set PjzV = A.VBProjects(Pjn)
End Function

Function IsSavP() As Boolean: IsSavP = CPj.Saved: End Function
Sub ChkPjSav(P As VBProject)
If Not P.Saved Then Raise "Pj is not saved"
End Sub
Function PjzPjf(Vbe As Vbe, Pjf) As VBProject
Dim I As VBProject
For Each I In Vbe.VBProjects
    If PjfzP(I) = Pjf Then Set PjzPjf = I: Exit Function
Next
End Function

Sub SavVbe(A As Vbe)
Dim P As VBProject
For Each P In A.VBProjects
    SavPj P
Next
End Sub

Function PjfyV() As String()
PjfyV = PjfyzV(CVbe)
End Function

Function PjfyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushNB PjfyzV, Pjf(P)
Next
End Function

Function PjnyV() As String()
PjnyV = PjnyzV(CVbe)
End Function

Function PjnyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushI PjnyzV, P.Name
Next
End Function

Function SrtRptV() As String()
SrtRptV = SrtRptzV(CVbe)
End Function

Function HasBarzV(A As Vbe, BarNm) As Boolean
HasBarzV = HasItn(A.CommandBars, BarNm)
End Function

Function HasPj(A As Vbe, Pjn$) As Boolean
HasPj = HasItn(A.VBProjects, Pjn)
End Function

Function HasPjfzV(A As Vbe, Pjf) As Boolean
Dim P As VBProject
For Each P In A.VBProjects
    If PjfzP(P) = Pjf Then HasPjfzV = True: Exit Function
Next
End Function

Function SrtRptzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy SrtRptzV, SrtRptzP(P)
Next
End Function

Private Sub VbeFunPfx__Tst()
'D Vbe_MthPfx(CVbe)
End Sub

Private Sub MthnyzV__Tst()
Brw MthnyzV(CVbe)
End Sub
