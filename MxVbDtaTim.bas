Attribute VB_Name = "MxVbDtaTim"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "InterAct"
Const CMod$ = CLib & "MxVbDtaTim."
Private M$, Beg As Date
Public NoStamp As Boolean
Sub TimBeg(Optional Msg$ = "Time")
If M <> "" Then TimEnd
M = Msg
Beg = Now
End Sub
Sub TimEnd(Optional Halt As Boolean)
Debug.Print M & " " & DateDiff("S", Beg, Now) & "(s)"
If Halt Then Stop
End Sub
Sub TimFun(FunNN)
Dim B!, E!, F
For Each F In Termy(FunNN)
    B = Timer
    Run F
    E = Timer
    Debug.Print F, "<-- Run"; E - B
Next
End Sub

Private Sub TimFun__Tst()
TimFun "TimFunA TimFunB"
End Sub

Sub TimFunA()
Dim J&, I&
For J = 0 To 100
    For I = 0 To 100
        Debug.Print I
    Next
Next
End Sub
Sub TimFunB()
Dim J&, I&
For J = 0 To 100
    For I = 0 To 100
        Debug.Print I
    Next
Next
End Sub
