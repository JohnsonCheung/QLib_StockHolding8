Attribute VB_Name = "MxVbUI"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "InterAct"
Const CMod$ = CLib & "MxVbUI."

Function CfmYes(Msg$) As Boolean
CfmYes = UCase(InputBox(Msg)) = "YES"
End Function

Sub PromptCnl(Optional Msg = "Should cancel and check")
If MsgBox(Msg, vbOKCancel) = vbCancel Then Stop
End Sub

Sub Done()
MsgBox "Done"
End Sub
Function Start(Optional Msg$ = "Start?", Optional Tit$ = "Start?") As Boolean
Start = MsgBox(Replace(Msg, "|", vbCrLf), vbQuestion + vbYesNo + vbDefaultButton1, Tit) = vbYes
End Function

Function Cfm(Msg$, Optional Tit$ = "Please confirm", Optional NoAsk As Boolean) As Boolean
If NoAsk Then Cfm = True: Exit Function
Cfm = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton1) = vbYes
End Function
