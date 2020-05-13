Attribute VB_Name = "MxVbDtaAp"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Dta"
Const CMod$ = CLib & "MxVbDtaAp."

Function AvzAy(Ay) As Variant()
If IsAv(Ay) Then AvzAy = Ay: Exit Function
Dim I: For Each I In Itr(Ay)
    Push AvzAy, I
Next
End Function

Function Av(ParamArray Ap()) As Variant()
Av = Ap
End Function

Function SyzAv(AvOf_Itm_or_Ay) As String()
Dim I: For Each I In Itr(AvOf_Itm_or_Ay)
    If IsArray(I) Then
        PushIAy SyzAv, I
    Else
        PushI SyzAv, I
    End If
Next
End Function

Function SyzAvNB(AvOf_Itm_or_Ay_NB()) As String()
Dim I: For Each I In Itr(AvOf_Itm_or_Ay_NB)
    If IsArray(I) Then
        PushNBAy SyzAvNB, I
    Else
        PushNB SyzAvNB, I
    End If
Next
End Function

Function SyzAp(ParamArray ApOf_Itm_Or_Ay()) As String()
Dim Av(): Av = ApOf_Itm_Or_Ay
SyzAp = SyzAv(Av)
End Function

Function SyzApNB(ParamArray ApOf_Itm_Or_Ay()) As String()
Dim Av(): Av = ApOf_Itm_Or_Ay
SyzApNB = SyzAvNB(Av)
End Function

Function Sy(ParamArray ApOf_Itm_Or_Ay()) As String()
Dim Av(): Av = ApOf_Itm_Or_Ay
Sy = SyzAv(Av)
End Function

Function SyNB(ParamArray ApOf_Itm_Or_Ay_NB()) As String()
Dim Av(): Av = ApOf_Itm_Or_Ay_NB
SyNB = SyzAvNB(Av)
End Function

Function IntAyzLngAy(LngAy&()) As Integer()
Dim I
For Each I In Itr(LngAy)
    PushI IntAyzLngAy, I
Next
End Function

Function IntAyzSS(IntSS$) As Integer()
Dim I
For Each I In Itr(SyzSS(IntSS))
    PushI IntAyzSS, I
Next
End Function

Function SyzAy(Ay) As String()
If IsSy(Ay) Then SyzAy = Ay: Exit Function
SyzAy = IntozAy(EmpSy, Ay)
End Function

Function SyzAyNB(Ay) As String()
Dim I: For Each I In Itr(Ay)
    If I <> "" Then PushI SyzAyNB, I
Next
End Function

Function IntozAyMap(Into, Ay, Map$)
IntozAyMap = IntozItrm(Into, Itr(Ay), Map)
End Function
