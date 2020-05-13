Attribute VB_Name = "MxVbAySyAdd"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbAySyAdd."

Function AddSyMsg(Sy$(), MsgStrOrSy) As String()
AddSyMsg = AddSy(Sy, CvSy(MsgStrOrSy))
End Function

Function AddSyAp(A$(), ParamArray SyAp()) As String()
AddSyAp = A
Dim IsY: For Each IsY In SyAp
    If Not IsSy(IsY) Then Raise "Some of ele in @SyAp is not Sy, but [" & TypeName(IsY) & "]"
    PushIAy AddSyAp, IsY
Next
End Function

Function AddSyStr(A$(), Str$) As String()
AddSyStr = AddAyEle(A, Str)
End Function

Private Sub AyzGp__Tst()
Dim Ay(), N%
GoSub T0
Exit Sub
T0:
    Ay = Array(1, 2, 3, 4, 5, 6)
    Ept = Array(Array(1, 2, 3, 4, 5), Array(6))
    N = 5
    GoTo Tst
Tst:
    Act = AyzGp(Ay, N%)
    C
    Return
End Sub

Function AyzGp(Ay, N%) As Variant()
Dim NEle&: NEle = Si(Ay): If NEle = 0 Then Exit Function
Dim Emp: Emp = Ay: Erase Emp
Dim M: M = Emp
Dim V, GpI%, Ix%: For Each V In Itr(Ay)
    PushI M, V
    GpI = GpI + 1
    If GpI = N Then
        GpI = 0
        PushI AyzGp, M
        M = Emp
    End If
Next
If Si(M) > 0 Then PushI AyzGp, M
End Function

Function AddAyAp(Ay, ParamArray Itm_or_AyAp())
Const CSub$ = CMod & "AddAyAp"
Dim Av(): Av = Itm_or_AyAp
If Not IsArray(Ay) Then Thw CSub, "Ay must be array", "Ay-TypeName", TypeName(Ay)
AddAyAp = Ay
Dim I: For Each I In Av
    If IsArray(I) Then
        PushIAy AddAyAp, I
    Else
        PushI AddAyAp, I
    End If
Next
End Function

Function AyzMap(Ay, MapFun$) As Variant()
Dim X
For Each X In Itr(Ay)
    Push AyzMap, Run(MapFun, X)
Next
End Function

Function AddAyEle(Ay, Ele)
AddAyEle = Ay
Push AddAyEle, Ele
End Function

Function AddEleAy(Ele, Ay)
Dim O: O = Ay: Erase O
Push O, Ele
PushAy O, Ay
AddEleAy = O
End Function
Function AddEleSy(Ele, Sy$()) As String()
Dim O$()
PushI O, Ele
PushIAy O, Sy
AddEleSy = O
End Function

Function AddAvItm(Ay, Ele) As Variant()
AddAvItm = AvzAy(Ay)
Push AddAvItm, Ele
End Function


Function AyzInc(Ay, Optional N& = 1)
AyzInc = NwAy(Ay)
Dim X
For Each X In Itr(Ay)
    PushI AyzInc, X + N
Next
End Function

Private Sub AddAy__Tst()
Dim Ay1(), Ay2()
GoSub T1
Exit Sub
T1:
    Ay1 = Array(1, 2, 2, 2, 4, 5)
    Ay2 = Array(2, 2)
    Ept = Array(1, 2, 2, 2, 4, 5, 2, 2)
    GoTo Tst
Tst:
    Act = AddAy(Ay1, Ay2)
    C
    Return
End Sub

Private Sub AmAddPfx__Tst()
Dim Sy$(), Pfx$
GoSub T1
Exit Sub
T1:
    Sy = SyzSS("1 2 3 4")
    Pfx = "* "
    Ept = SyzAp("* 1", "* 2", "* 3", "* 4")
    GoTo Tst
Tst:
    Act = AmAddPfx(Sy, Pfx)
    C
    Return
End Sub

Private Sub AmAddPfxSfx__Tst()
Dim Sy$(), Act$(), Sfx$, Pfx$, Exp$()
Sy = SyzAp(1, 2, 3, 4)
Pfx = "* "
Sfx = "#"
Exp = SyzAp("* 1#", "* 2#", "* 3#", "* 4#")
GoSub Tst
Exit Sub
Tst:
Act = AmAddPfxSfx(Sy, Pfx, Sfx)
Debug.Assert IsEqAy(Act, Exp)
Return
End Sub

Function TabLines$(L$)
TabLines = JnCrLf(AmAddPfx(SplitCrLf(L), vbTab))
End Function
Function TabSy(Sy$()) As String()
TabSy = AmAddPfx(Sy, vbTab)
End Function

Private Sub AmAddSfx__Tst()
Dim Sy$(), Sfx$
Sy = SyzSS("1 2 3 4")
Sfx = "#"
Ept = SyzSS("1# 2# 3# 4#")
GoSub Tst
Exit Sub
Tst:
    Act = AmAddSfx(Sy, Sfx)
    C
    Return
End Sub
