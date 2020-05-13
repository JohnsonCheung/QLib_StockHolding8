Attribute VB_Name = "MxVbStrLbl123"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Fmt"
Const CMod$ = CLib & "MxVbStrLbl123."
Private Sub Lbl123__Tst()
Dmp Lbl123(543)
End Sub

Function Lbl123(L) As String()
'Lbl123: :Ly ! #Lbl-123# is Ly of 0-3 Ln depend of L, where L is bet 0 and 999
Const CSub$ = CMod & "Lbl123"
'@L : :Num 1 to 999, else Thw
'Ret  : :Lbl123: is sy-of-1-to-3 ele.  Las-ele is dig no 0, Las-2nd-if-any is ten-dig, Las-3rd-if-any is hundred-dig
If Not IsBet(L, 1, 999) Then Thw CSub, "Length should be bet 1 999", "Length", L
PushNB Lbl123, Lbl123HundredLin(L)
PushNB Lbl123, Lbl123TenLin(L)
PushI Lbl123, Lbl123DigLin(L)
End Function

Function Lbl123DigLin$(L)
Const C$ = "1234567890"
Dim N&: N = (L \ 10) + 1
Lbl123DigLin = Left(Dup(C, N), L)
End Function

Function Lbl123TenLin$(L)
If L < 9 Then Exit Function
Dim O$()
    PushI O, Space(9)
    Dim J%: For J = 0 To (L \ 10)
        Dim C$: C = Right(CStr((J Mod 10) + 1), 1)
        PushI O, Dup(C, 10)
    Next
Lbl123TenLin = Left(Jn(O), L)
End Function

Function Lbl123HundredLin$(L)
If L < 99 Then Exit Function
Dim O$()
    PushI O, Space(99)
    Dim J%: For J = 0 To (L \ 100)
        Dim C$: C = Right(CStr((J Mod 10) + 1), 1)
        PushI O, Dup(C, 100)
    Next
Lbl123HundredLin = Left(Jn(O), L)
End Function
