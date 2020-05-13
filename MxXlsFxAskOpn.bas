Attribute VB_Name = "MxXlsFxAskOpn"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "MxXlsFxAskOpn."
Function AskOpnFxAy(Fxy$()) As Boolean
Dim Fx: For Each Fx In Itr(Fxy)
    If NoFfn(Fx) Then Exit Function
Next
Dim A$: A = JnCrLf(Fxy)
Select Case MsgBox("File exists:" & A & vbLf & vbLf & _
    "[Yes] = Re-generate and over-write" & vbLf & _
    "[No] = Open existing file" & vbLf & _
    "[Cancel] = Cancel", vbYesNoCancel + vbDefaultButton2 + vbQuestion, "Generate file.")
Case VbMsgBoxResult.vbNo:
    OpnFxy Fxy
    AskOpnFxAy = True
    Exit Function
End Select
End Function
