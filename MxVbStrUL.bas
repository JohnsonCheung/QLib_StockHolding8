Attribute VB_Name = "MxVbStrUL"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrUL."

Function UL(S, Optional ULChr$ = "=") As String()
PushI UL, S
PushI UL, ULn(S, ULChr)
End Function

Sub PushUL(O$(), S$, Optional ULChr$ = "=")
PushIAy O, UL(S, ULChr)
End Sub

Function ULinzLines$(Lines$, Optional ULinChr$ = "-")
ULinzLines = Lines & vbCrLf & Dup("-", WdtzLines(Lines))
End Function

Function ULn$(S, Optional ULnChr$ = "=")
ULn = String(Len(S), ULnChr)
End Function
