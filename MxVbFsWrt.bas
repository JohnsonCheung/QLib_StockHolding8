Attribute VB_Name = "MxVbFsWrt"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbFsWrt."
Const CNs$ = "Fs"
Function AppStr$(S, Ft)
Dim Fno%: Fno = FnoA(Ft)
Print #Fno, S;
Close #Fno
AppStr = Ft
End Function

Function WrtStr$(S, Ft, Optional OvrWrt As Boolean)
If OvrWrt Then DltFfnIf Ft
Dim Fno%: Fno = FnoO(Ft)
Print #Fno, S;
Close #Fno
WrtStr = Ft
End Function

Sub WrtAy(Ay, Ft$, Optional OvrWrt As Boolean)
If OvrWrt Then DltFfnIf Ft
Dim Fno%: Fno = FnoO(Ft)
Dim I: For Each I In Itr(Ay)
    Print #Fno, I
Next
Close #Fno
End Sub
