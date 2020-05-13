Attribute VB_Name = "MxXlsRgCmt"
Option Explicit
Option Compare Text
Const CNs$ = "Xls"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsRgCmt."

Sub SetCmtzAy(Rg() As Range, Cmt$())
Dim J%: For J = 0 To MinUB(Rg, Cmt)
    SetCmt Rg(J), Cmt(J)
Next
End Sub

Private Sub SetCmt__Tst()
Dim R As Range: Set R = A1zWs(CWs)
SetCmt R, "lskdfjsdlfk"
End Sub

Function HasCmt(R As Range) As Boolean
HasCmt = Not IsNothing(R.Comment)
End Function

Sub SetCmt(R As Range, Cmt$)
If Not HasCmt(R) Then
    R.AddComment.Text Cmt
    Exit Sub
End If
Dim C As Comment: Set C = R.Comment
If C.Text = Cmt Then Exit Sub
C.Text Cmt
End Sub


Function CvCmt(A) As Comment
Set CvCmt = A
End Function

Function CmtAyzRg(A As Range) As Comment()
Dim C As Comment: For Each C In WszRg(A).Comments
    Dim CmtRg As Range: Set CmtRg = C.Parent
    If HasCell(A, CmtRg) Then PushObj CmtAyzRg, C
Next
End Function
Sub DltCmtzRg(A As Range)
Dim Cmt() As Comment: Cmt = CmtAyzRg(A)
Dim I: For Each I In Itr(Cmt)
    CvCmt(I).Delete
Next
End Sub
