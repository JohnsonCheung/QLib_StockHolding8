Attribute VB_Name = "MxAcsSts"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "MxAcsSts."
Private Msg$()

Private Sub Sts__Tst()
Dim J%: For J = 0 To 10
    Sts J
Next
Stop
End Sub
Sub StsQry(QNm$): Sts "Running query " & QNm & "....": End Sub
Sub StsLnk(T): Sts "Linking " & T & " ........": End Sub
Sub ShwSts(): BrwAy RevAy(Msg): End Sub
Sub Sts(Sts)
Dim A$: A = Now & " " & Sts
PushS Msg, A
RfhBtn
Debug.Print A
End Sub

Sub ClrSts()
Erase Msg
RfhBtn
End Sub


Private Sub RfhBtn()
Dim B As Access.CommandButton: Set B = Btn
If IsNothing(B) Then Exit Sub
B.Caption = Si(Msg) & " Msgs"
B.Requery
End Sub

Private Function Btn() As Access.CommandButton
Dim F As Access.Form: Set F = CFrm: If IsNothing(F) Then Exit Function
If Not HasCtl(F, "CmdMsg") Then Exit Function
Dim C As Control: Set C = F.Controls("CmdMsg")
If TypeName(C) <> "CommandButton" Then Raise "MxSts.Btn: Has a control with Name CmdMsg, but it is not a CommandButton, but[" & TypeName(C.Object) & "]"
Set Btn = F.Controls("CmdMsg")
End Function

