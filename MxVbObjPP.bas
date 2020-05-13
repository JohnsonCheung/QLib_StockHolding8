Attribute VB_Name = "MxVbObjPP"
Option Compare Text
Option Explicit
Const CLib$ = "QItrObj."
Const CMod$ = CLib & "MxVbObjPP."
':PP: :Prpp-PP$ #Spc-Separated-Prpp# ! Each ele is a Prpp
':Prpp: :Dotn   #Prp-Pth# ! Prp-Pth-of-an-Object
Function OyAdd(Oy1, Oy2)
Dim O: O = Oy1
PushObjAy O, Oy2
OyAdd = O
End Function

Function FstzOyEq(Oy, Prpp, V)
Set FstzOyEq = FstObjByEq(Itr(Oy), Prpp, V)
End Function

Function AvzOyP(Oy, Prpp) As Variant()
AvzOyP = IntozOyP(EmpAv, Oy, Prpp)
End Function

Function IntozOyP(Into, Oy, Prpp)
Dim O: O = Into: Erase O
Dim Obj: For Each Obj In Itr(Oy)
    Push O, Opv(Obj, Prpp)
Next
IntozOyP = O
End Function

Function IntAyzOyP(Oy, Prpp) As Integer()
IntAyzOyP = IntozOyP(EmpIntAy, Oy, Prpp)
End Function

Function SyzOyP(Oy, Prpp) As String()
Stop
SyzOyP = IntozOyP(EmpSy, Oy, Prpp)
End Function

Function OyeNothing(Oy)
OyeNothing = NwAy(Oy)
Dim Obj As Object
For Each Obj In Oy
    If Not IsNothing(Obj) Then PushObj OyeNothing, Obj
Next
End Function

Function OywNmPfx(Oy, NmPfx$)
Dim Obj, O
O = Oy: Erase O
For Each Obj In Itr(Oy)
    If HasPfx(Obj.Name, NmPfx) Then PushObj O, Obj
Next
OywNmPfx = O
End Function

Function OywNm(Oy, B As WhNm)
Dim Obj, O
O = Oy: Erase O
For Each Obj In Itr(Oy)
    If HitNm(Obj.Name, B) Then PushObj OywNm, Obj
Next
End Function

Function OywPredXPTrue(Oy, XP$, P$)
Dim O, Obj As Object
O = Oy
Erase O
For Each Obj In Itr(Oy)
    If Run(XP, Obj, P) Then
        PushObj O, Obj
    End If
Next
OywPredXPTrue = O
End Function

Function FstzObj(Oy, Prpp$, V)
'Ret : Fst Obj in @Oy having @Prpp = @V
Dim Obj: For Each Obj In Itr(Oy)
    If Opv(Obj, Prpp) = V Then Asg Obj, FstzObj: Exit Function
Next
End Function
Function OyzItr(Itr) As Variant()
Dim O
For Each O In Itr
    PushObj OyzItr, O
Next
End Function
Function OywIn(Oy, Prpp, InAy)
Dim Obj As Object, O
If Si(Oy) = 0 Or Si(InAy) Then OywIn = Oy: Exit Function
O = Oy
Erase O
For Each Obj In Itr(Oy)
    If HasEle(InAy, Opv(Obj, Prpp)) Then PushObj O, Obj
Next
OywIn = O
End Function

Function LyzObjPP(Obj As Object, PP$) As String()
Dim Prpp: For Each Prpp In SyzSS(PP)
    PushI LyzObjPP, Prpp & " " & Opv(Obj, Prpp)
Next
End Function

Private Sub OyDrs__Tst()
'VisWs DrsNwWs(OyDrs(CurrentDb.TableDefs("Z_UpdSeqFld").Fields, "Name Type OrdinalPosition"))
End Sub

Private Sub OyP_Ay__Tst()
Dim CdPanAy() As CodePane
Stop
'CdPanAy = Oy(CPj.MdAy).PrpVy("CodePane", CdPanAy)
Stop
End Sub
Private Sub LyzObjPP__Tst()
Dim Obj As Object, PP$
GoSub T0
Exit Sub
T0:
    Set Obj = New DAO.Field
    PP = "Name Type Size"
    GoTo Tst
Tst:
    Act = LyzObjPP(Obj, PP)
    C
    Return
End Sub
