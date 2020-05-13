Attribute VB_Name = "JMxAm"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "JMxAm."
#If False Then
Function AmTrim(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmTrim, Trim(I)
Next
End Function

Function AmAddPfx(Ay, Pfx$) As String()
Dim I: For Each I In Itr(Ay)
    PushS AmAddPfx, Pfx & I
Next
End Function

Function AmAddPfxTab(Ay) As String()
AmAddPfxTab = AmAddPfx(Ay, vbTab)
End Function

#End If
