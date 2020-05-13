Attribute VB_Name = "MxVbStrSplit"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrSplit."

Function SplitComma(S) As String()
SplitComma = Split(S, ",")
End Function

Function SplitCommaSpc(S) As String()
SplitCommaSpc = Split(S, ", ")
End Function

Function Ly(Lines) As String()
Ly = SplitCrLf(Lines)
End Function

Function SplitCrLf(S) As String()
SplitCrLf = Split(Replace(S, vbCr, ""), vbLf)
End Function

Function SplitTab(S) As String()
SplitTab = Split(S, vbTab)
End Function

Function SplitDot(S) As String()
SplitDot = Split(S, ".")
End Function

Sub SplitPosy__Tst()
Dim S$, Posy%()
GoSub T1
Exit Sub
T1:
    S = "1234567890"
    Posy = IntAy(3, 7, 10)
    Ept = Sy("12", "456", "89", "")
    GoTo Tst
Tst:
    Act = SplitPosy(S, Posy)
    C
    Return
End Sub

Function SplitPosy(S, Posy%()) As String()
If Si(Posy) = 0 Then PushI SplitPosy, S: Exit Function
Dim PrvP%, L%
Dim P: For Each P In Posy
    L = P - PrvP - 1
    PushI SplitPosy, Mid(S, PrvP + 1, L)
    PrvP = P
Next
PushI SplitPosy, Mid(S, PrvP + 1)
End Function

Function SplitColon(S) As String()
SplitColon = Split(S, ":")
End Function

Function SplitSemi(S) As String()
SplitSemi = Split(S, ";")
End Function

Function SplitSpc(S) As String()
SplitSpc = Split(S, " ")
End Function

Function SplitSsl(S) As String()
SplitSsl = Split(RplDblSpc(Trim(S)), " ")
End Function

Function SplitVBar(S) As String()
SplitVBar = CvSy(Split(S, "|"))
End Function

Function LyzLinesy(Linesy$()) As String()
Dim L: For Each L In Itr(Linesy)
    PushIAy LyzLinesy, SplitCrLf(L)
Next
End Function
