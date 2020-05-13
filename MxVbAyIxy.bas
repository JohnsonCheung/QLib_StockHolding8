Attribute VB_Name = "MxVbAyIxy"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ay"
Const CMod$ = CLib & "MxVbAyIxy."

Private Sub AsgIx__Tst()
Dim Drs As Drs, FF$, A%, B%, C%, EA%, EB%, Ec%
GoSub T1
Exit Sub
T1:
    Drs.Fny = SyzSS("A B C")
    FF = "C B A"
    EA = 0
    EB = 1
    Ec = 2
    GoTo Tst
Tst:
    AsgIx Drs, FF, A, B, C
    Debug.Print A = EA
    Debug.Print B = EB
    Debug.Print C = Ec
    Return
End Sub

Function IxyzFnyFF(Fny$(), FF$) As Integer()
Dim F: For Each F In FnyzFF(FF)
    PushI IxyzFnyFF, IxzAy(Fny, F, ThwEr:=EiThwEr)
Next
End Function

Sub AsgIxzFny(Fny$(), FF$, ParamArray OIxAp())
Dim Ix, J%: For Each Ix In IxyzFnyFF(Fny, FF)
    OIxAp(J) = Ix
    J = J + 1
Next
End Sub

Sub AsgIx(A As Drs, FF$, ParamArray OIxAp())
Dim Ix, J%: For Each Ix In IxyzFnyFF(A.Fny, FF)
    OIxAp(J) = Ix
    J = J + 1
Next
End Sub

Function Ele(Ay, Ix)
If IsBet(Si(Ay), 0, Ix) Then Ele = Ay(Ix)
End Function
Function EleIxy(Ay, Ele) As Long()
Dim J&
Dim V: For Each V In Itr(Ay)
    If V = Ele Then PushI EleIxy, J
    J = J + 1
Next
End Function

Function EleIx&(Ay, Ele, Optional Bix& = 0)
EleIx = IxzAy(Ay, Ele, Bix)
End Function

Function IxzAy&(Ay, Itm, Optional Bix& = 0, Optional ThwEr As EmThw)
Const CSub$ = CMod & "IxzAy"
Dim J&: For J = Bix To UB(Ay)
    If Ay(J) = Itm Then IxzAy = J: Exit Function
Next
If ThwEr = EiThwEr Then
    Thw CSub, "Itm not found in Ay", "Itm Si(Ay) Ay", Itm, Si(Ay), Ay
End If
IxzAy = -1
End Function

Function IxyzAlwEmp(Ay, SubAy) As Long()
Dim HasNegIx As Boolean, Ix&
Dim I: For Each I In Itr(SubAy)
    Ix = IxzAy(Ay, I, , EiNoThw)
    PushI IxyzAlwEmp, Ix
Next
End Function

Function Cxy(Fny$(), SubFny$()) As Integer()
Const CSub$ = CMod & "Cxy"
Dim F: For Each F In Itr(SubFny)
    Dim I%: I = IxzAy(Fny, F)
    If I = -1 Then
        Thw CSub, "Ele in @SubFny is not found in @Fny", "Ele SubFny Fny", F, SubFny, Fny
    End If
    PushI Cxy, I
Next
End Function

Function Ixy(Ay, SubAy) As Long()
Const CSub$ = CMod & "Ixy"
Dim O&(): O = IxyzAlwEmp(Ay, SubAy)
Dim Ix: For Each Ix In Itr(O)
    If Ix = -1 Then
        Thw CSub, "Negative index", "Ay SubAy Ixy", Ay, SubAy, Ixy
    End If
Next
Ixy = O
End Function

Function IntIxy(Ay, SubAy) As Integer()
Dim J&, U&
For J = 0 To UB(SubAy)
    PushI IntIxy, IxzAy(Ay, SubAy(J))
Next
End Function

Function IxyzDup(Ay) As Long()
Dim A As Dictionary: Set A = Aet(AwDup(Ay))
If IsEmpAet(A) Then Exit Function
Dim J&
For J = 0 To UB(Ay)
    If A.Exists(Ay(J)) Then PushI IxyzDup, J
Next
End Function
Function IxItrExl(IxUB&, ExlRxy&())
IxItrExl = Itr(MinusAy(NwIxy(IxUB), ExlRxy))
End Function

Function NwIxy(IxUB&) As Long()
If IxUB < 0 Then Exit Function
Dim O&(): ReDim O(IxUB)
Dim J&: For J = 0 To IxUB
    O(J) = J
Next
NwIxy = O
End Function
