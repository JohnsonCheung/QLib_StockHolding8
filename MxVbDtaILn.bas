Attribute VB_Name = "MxVbDtaILn"
Option Explicit
Option Compare Text

Function LyzILny(L() As ILn) As String()
Dim J%: For J = 0 To ILnUB(L)
    PushI LyzILny, L(J).Ln
Next
End Function
