Attribute VB_Name = "JMxAyOp"
Option Compare Text
Const CMod$ = CLib & "JMxAyOp."
#If False Then
Option Explicit
Function ChkDup$(Ay, Optional ItmNm = "item")
Dim Dup: Dup = AwDup(Ay)
If Si(Dup) > 0 Then ChkDup = "Following " & ItmNm & " are duplicated:" & vbCrLf & TabAy(Dup) & vbCrLf
End Function


Function HasEleFm(Ay, Ele, Fm&) As Boolean
Dim J&: For J = Fm To UB(Ay)
    If Ay(J) = Ele Then HasEleFm = True: Exit Function
Next
End Function


Function AddAyPfx(Ay, Pfx$) As String()
Dim I: For Each I In Itr(Ay)
    PushS AddAyPfx, Pfx & I
Next
End Function
'---------------------------
Function SyzAy(Ay) As String()
If IsSy(Ay) Then SyzAy = Ay: Exit Function
Dim I: For Each I In Itr(Ay)
    PushI SyzAy, I
Next
End Function



#End If
