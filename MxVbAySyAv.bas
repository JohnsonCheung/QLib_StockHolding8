Attribute VB_Name = "MxVbAySyAv"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbAySyAv."

Function DrzSyAv(SyAv(), R&, USy%) As String()
Dim C%: For C = 0 To USy
    PushI DrzSyAv, SyAv(C)(R)
Next
End Function

Sub ChkIsSyAv(SyAv(), Optional Fun$ = "ChkIsSyAv")
If IsSyAv(SyAv) Then Exit Sub
Thw Fun, "Given SyAv is having Sy in each of element"
End Sub

Sub ChkIsSamSiSyAv(SyAv(), Optional Fun$ = "ChkIsSyAv")
If IsSamSiSyAv(SyAv) Then Exit Sub
Thw Fun, "Given SyAv does not have same size for each of the element"
End Sub

Function IsSyAv(Av()) As Boolean
Dim Sy: For Each Sy In Itr(Av)
    If Not IsSy(Sy) Then Exit Function
Next
IsSyAv = True
End Function

Function IsSamSiSyAv(SyAv()) As Boolean
Dim Fst As Boolean: Fst = True
Dim S&
Dim Sy: For Each Sy In Itr(SyAv)
    If Fst Then
        Fst = False
        S = Si(Sy)
    Else
        If S <> Si(Sy) Then Exit Function
    End If
Next
IsSamSiSyAv = True
End Function
