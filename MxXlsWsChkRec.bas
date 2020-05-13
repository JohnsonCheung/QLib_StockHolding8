Attribute VB_Name = "MxXlsWsChkRec"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "MxXlsWsChkRec."

Sub ChkWsRec(Fx$, W$, RecShdSy$())

End Sub

Sub ChkWsRecV2(Fx$, W$, RecShdSy$())
Dim O$(), Ay() As RecShd
Ay = RecShdAy(RecShdSy)
'PushIAy O, Blnk(Fx, W, ShdNonBlnkFny(Ay))
'PushIAy O, QLib_MxRecMsg.NAllIn(Fx, W, ShdAllInAy(Ay))
End Sub

Private Function ShdNonBlnkFny(A() As RecShd) As String()

End Function

Private Function ShdAllInAy(A() As RecShd) As FldVy()

End Function


Sub ChkWsRecV1(Fx$, W$, RecShdSy$())
Dim Shd() As RecShd: Shd = RecShdAy(RecShdSy)
Dim Fny$(): Fny = AwDis(FnyzRecShdAy(Shd))
Dim O$(), V() As FldVy: V = DisFldVyAyzFx(Fx, W, Fny)
Dim Vy()
Dim J%: For J = 0 To RecShdUB(Shd)
    Vy = FndVy(V, Shd(J).F)
    PushIAy O, RecEmsgzCol(Vy, Shd(J))
Next
ChkFxwEr Fx, W, O
End Sub

Private Sub ColValErMsgPerCol__Tst()
Dim ColVy(), Fldn$, Op As RecShdOp, Val$
GoSub Z1
Exit Sub
Z1:
    ColVy = Av("8601", "8701")
    Fldn = ""
    Op = ShdAllInLis
    Val = "8601"
    Ept = Sy()
    GoTo Tst
Tst:
    Dim A As RecShd
    A = RecShdAllIn(Fldn, Op, Val)
    Act = RecEmsgzCol(ColVy, A)
    C
    Return
End Sub

Private Function RecEmsgzCol(ColVy(), A As RecShd) As String()
':RecEmsg: :Emsg ! #ColVy-Emsg#
':Emsg :Sy ! #Er-Msg# If no error, return EmpSy, if error, return the error message.  So the parameter should be enough data to test iser and report error msg
':Msg :Sy  ! #Msg# It takes parameter to generate non-empty-Sy
Dim Ts() As ColErTy: Ts = ColErTyAy(ColVy, A)
Dim Es() As RecErDta, E As RecErDta
Dim Ty: For Each Ty In Itr(Ts)
    PushIAy RecEmsgzCol, RecEmsg(E)
Next
End Function

Private Function CvColErTy(T) As ColErTy
CvColErTy = T
End Function

Private Function ColErTyAy(Vy(), A As RecShd) As ColErTy()
'For each cur-ColVy and cur-RecShd, how many dif err type (ColErTy-NoEr SomBlnk MisVal SomNotIn)
Dim Op As RecShdOp: Op = A.Op
Select Case True
Case Op = ShdAllInLis:
    Dim T As ColErTy
Case Op = ShdNBlnk
    If HasEle(Vy, "") Then PushI ColErTyAy, T
Case Else: ' ThwEmEr CSub, "RecShdOp", Op, RecShdOpss
End Select
End Function


Private Function RecShdOpAy() As String()
Static Y$(): If Si(Y) = 0 Then Y = Split(RecShdOpss)
RecShdOpAy = Y
End Function

Private Function SampRecShdSy() As String()

End Function

Private Sub SampRecShdSy_Res()
#If False Then
chkWsDta Fx Wsn
  ShdNoNB
ChkWsCol
  
ShdNoNB Material Plnt
ShdInLis Plnt 8601 8701
ShdInTbl [Plnt
ShdInTbc
#End If
End Sub

Private Sub ShdAllInEmsgzForSy__Tst()
Dim A$(): A = ShdAllInEmsgzForSy(EmpSy)

End Sub

Function ShdAllInEmsgzForSy(ForColSy$(), ParamArray InSyap()) As String()

Dim ShdBeVal: For Each ShdBeVal In Av
Next
End Function

Function ShdNBEmsglzForFxwc$(Fx$, W$, C$)
'If HasBlnkzFxwc(Fx, W, C) Then ShdNBEmsglzForFxwc = ColBlnkMsgl(C)
End Function
