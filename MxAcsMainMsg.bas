Attribute VB_Name = "MxAcsMainMsg"
Option Explicit
Option Compare Text
Const CLib$ = "QAcs."
Const CMod$ = CLib & "MxAcsMainMsg."
'Assume there is Application.Forms("Main").Msg (TextBox)
'MMsg means Main.Msg (TextBox)
Sub ClrMainMsg()
Dim M As TextBox: Set M = MainMsgBox
If Not IsNothing(M) Then M.Value = ""
End Sub

Sub SetMainMsgzQnm(QryNm)
SetMainMsg "Running query: (" & QryNm & ")...."
End Sub

Sub SetMainMsg(Msg$)
On Error Resume Next
SetTBox MainMsgBox, Msg
End Sub

Property Get MainMsgBox() As Access.TextBox
On Error Resume Next
Set MainMsgBox = MainFrm.Controls("Msg")
End Property

Property Get MainFrm() As Access.Form
On Error Resume Next
Set MainFrm = Access.Forms("Main")
End Property

Sub SetTBox(A As Access.TextBox, Msg$)
Dim CrLf$, B$
If A.Value <> "" Then CrLf = vbCrLf
B = LineszLasN(A.Value & CrLf & Now & " " & Msg, 5)
A.Value = B
DoEvents
End Sub
