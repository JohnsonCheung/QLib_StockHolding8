Attribute VB_Name = "MxDtaDaDyTfm"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDtaDaDyTfm."

Function DyzAyV(Ay, V) As Variant()
Dim I: For Each I In Itr(Ay)
    PushI DyzAyV, Array(I, V)
Next
End Function

Function DyzVAy(V, Ay) As Variant()
Dim I: For Each I In Itr(Ay)
    PushI DyzVAy, Array(V, I)
Next
End Function
