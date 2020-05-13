Attribute VB_Name = "MxVbAyCvToPrimAy"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbAyCvToPrimAy."

Function DteAyzSy(Sy$()) As Date()
Dim I: For Each I In Sy
    PushI DteAyzSy, I
Next
End Function

Function DblAyzSy(Sy$()) As Double()
Dim I: For Each I In Sy
    PushI DblAyzSy, I
Next
End Function
