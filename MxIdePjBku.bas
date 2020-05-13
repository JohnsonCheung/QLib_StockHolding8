Attribute VB_Name = "MxIdePjBku"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdePjBku."

Sub BkuP(Optional Msg$ = "Bku")
BkuzP CPj, Msg
End Sub

Sub BkuzP(P As VBProject, Optional Msg$ = "Bku")
BkuFfn Pjf(P), Msg
End Sub

Sub BrwBkPthP()
BrwPth BkPthzP(CPj)
End Sub

Function BkPthzP$(P As VBProject)
BkPthzP = BkPth(Pjf(P))
End Function

Function BkPjfy() As String()
BkPjfy = BkFfnAy(PjfP)
End Function

Function LasBkPjf$()
LasBkPjf = LasBkFfn(PjfP)
End Function
