Attribute VB_Name = "MxIdeSrcDblQStr"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeSrcVstr."

Function TakDqStr$(S)
If FstChr(S) <> """" Then Exit Function
Dim P%: P = EndPos(2, S, 0)
If P = 0 Then Stop: Exit Function
TakDqStr = Replace(Mid(S, 2, P - 2), """""", """")
End Function

Function EndPos%(Fm%, S, Lvl%)
LoopTooMuch CSub, Lvl
Dim P%: P = InStr(Fm, S, """"): If P = 0 Then Exit Function
If Mid(S, P + 1, 1) <> """" Then EndPos = P: Exit Function
EndPos = EndPos(P + 2, S, Lvl + 1)
End Function
Private Sub TakDqStr__Tst()
Dim S$
'GoSub T1
GoSub T2
'GoSub T3
Exit Sub
T1: S = """aa""": Ept = "aa":       GoTo Tst
T2: S = """aa""""""": Ept = "aa""": GoTo Tst
T3: S = """aa""": Ept = "aa":       GoTo Tst
Tst: Act = TakDqStr(S): Debug.Assert Act = Ept: Return
End Sub

Private Sub RmvVstr__Tst()
Dim Ln$
GoSub Z
Exit Sub
Z:
    Ln = "aa""""aa""""": Ept = "aa""bb"
    Ept = "aa""""bb"
    GoTo Tst
Tst:
    Act = RmvVstr(Ln)
    C
    Return
End Sub

Private Sub RmvVstrzS__Tst()
VcAy RmvVrmk(RmvVstrzS(SrczP(CPj)))
End Sub

Function RmvVstrzS(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushI RmvVstrzS, RmvVstr(L)
Next
End Function

Function RmvVstr$(Ln) ':Ln #Rmv-VbStr#
If IsVrmkLn(Ln) Then RmvVstr = Ln: Exit Function
Dim O$, L$, S$, J%
L = Ln
Do
    If L = "" Then RmvVstr = O: Exit Function
    O = O & ShfBef(L, vbDblQ)
    If FstChr(L) <> vbDblQ Then RmvVstr = O & L: Exit Function
    S = ShfDqStr(L): If S <> "" Then O = O & vb2DblQ
    LoopTooMuch CSub, J
Loop
X: RmvVstr = O
End Function

Function ShfDqStr$(OLin)
'Assume FstChr(OLin) is vbDblQ
Dim E%: E = EndvbquoPos(OLin, 2): If E = 0 Then Thw "ShfDqStr", "Given @Ln [Assume fstChr is vbDblQ] has no EndDblQPos", "Ln BegvbquoPos", OLin, 1
OLin = Mid(OLin, E + 1)
ShfDqStr = Mid(OLin, 2, E - 2)
End Function

Private Function EndvbquoPos%(S, BegvbquoPos%)
Dim B%: B = BegvbquoPos
Dim P%, J%
Do
    P = InStr(B, S, vbDblQ): If P = 0 Then Exit Function
    If Mid(S, P + 1, 1) <> vbDblQ Then EndvbquoPos = P: Exit Function
    B = P + 2
    LoopTooMuch "EndvbquoPos", J
Loop

End Function
