Attribute VB_Name = "MxVbDtaItr"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Obj"
Const CMod$ = CLib & "MxVbDtaItr."

Function LinsszItr$(Itr)
LinsszItr = JnSpc(AvzItr(Itr))
End Function

Function TermLnzItr$(Itr)
TermLnzItr = Termln(AvzItr(Itr))
End Function

Function ItrClnAy(Itr)
If Itr.Count = 0 Then Exit Function
Dim X
For Each X In Itr
    ItrClnAy = Array(X)
    Exit Function
Next
End Function

Function NItpTrue(Itr, BoolPrpNm)
Dim O&, X
For Each X In Itr
    If CallByName(X, BoolPrpNm, VbGet) Then
        O = O + 1
    End If
Next
NItpTrue = O
End Function

Function FstItm(Itr)
Dim X: For Each X In Itr
    Asg X, _
        FstItm
    Exit Function
Next
Set FstItm = Nothing
End Function

Function FstItmPredXP(Ay, XP$, P$)
Dim X: For Each X In Ay
    If Run(XP, X, P) Then
        Asg FstItmPredXP, _
            X
        Exit Function
    End If
Next
End Function

Function FstObjByEq(Itr, PrpNm, V)
'Ret : fst ele in @Itr with its prpOf-@Prpp eq to @V
Dim Obj: For Each Obj In Itr
    If Opv(Obj, PrpNm) = V Then Set FstObjByEq = Obj: Exit Function
Next
Set FstObjByEq = Nothing
End Function

Function FstObjByNm(Itr, Nm$) 'Return first element in Itr with its PrpNm=Nm being true
Set FstObjByNm = FstObjByEq(Itr, "Name", Nm)
End Function

Function FstObjByTruePrp(Itr, TruePrp$)
'Ret : fst Obj in @Itr wi its @TruePrp being true
Set FstObjByTruePrp = FstObjByEq(Itr, TruePrp, True)
End Function

Function HasItn(Itr, Nm) As Boolean
Dim Obj: For Each Obj In Itr
    If Opv(Obj, "Name") = Nm Then HasItn = True: Exit Function
Next
End Function

Function HasItp(Itr, P, V) As Boolean
Dim Obj: For Each Obj In Itr
    If Opv(Obj, P) = V Then HasItp = True: Exit Function
Next
End Function

Function HasItrTruePrp(Itr, Prpp) As Boolean
Dim I
For Each I In Itr
    If Opv(CvObj(I), Prpp) Then HasItrTruePrp = True: Exit Function
Next
End Function

Function AvzItrm(Itr, Map$) As Variant()
AvzItrm = IntozItrm(EmpAv, Itr, Map)
End Function

Private Sub AvzItp__Tst()
Vc AvzItp(CPj.VBComponents, "CodeModule.CountOfLines")
End Sub

Function MaxzItp(Itr, Prpp)
Dim O, Obj: For Each Obj In Itr
    O = Max(O, Opv(Obj, Prpp))
Next
MaxzItp = O
End Function

Function NyzItr(Itr) As String()
NyzItr = Itn(Itr)
End Function

Function NyzItrEq(Itr, Prpp, V) As String()
Dim Obj: For Each Obj In Itr
    If Opv(Obj, Prpp) = V Then PushI NyzItrEq, Objn(Obj)
Next
End Function
Function NyzOy(Oy) As String()
NyzOy = Itn(Itr(Oy))
End Function

Function VyzItrP(Itr, Prpp) As Variant()
Dim Obj: For Each Obj In Itr
    Push VyzItrP, Opv(Obj, Prpp)
Next
End Function
Function FstItn$(Itr)
Dim I: For Each I In Itr
    FstItn = Objn(I)
Next
End Function

Function ItnWhPrpNB(Itr, Prpp) As String()
Dim I: For Each I In Itr
    If Opv(I, Prpp) <> "" Then PushI ItnWhPrpNB, Objn(I)
Next
End Function

Function ItnWhPrpBlnk(Itr, Prpp) As String()
Dim I: For Each I In Itr
    If Opv(I, Prpp) = "" Then PushI ItnWhPrpBlnk, Objn(I)
Next
End Function

Function Itn(Itr) As String()
Dim I: For Each I In Itr
    PushI Itn, Objn(I)
Next
End Function

Function HasTruePrp(Itr, Prpp) As Boolean
Dim I: For Each I In Itr
    If Opv(I, Prpp) Then HasTruePrp = True: Exit Function
Next
End Function


Function IwEq(Itr, Prpp, V)
IwEq = ItrClnAy(Itr)
Dim Obj: For Each Obj In Itr
    If Opv(Obj, Prpp) = V Then PushObj IwEq, Obj
Next
IwEq = Obj
End Function




Function NIwEq&(Itr, Prpp, V)
Dim O&, Obj: For Each Obj In Itr
    If Opv(Obj, Prpp) = V Then O = O + 1
Next
NIwEq = O
End Function

Function PrpNy(Itr) As String()
PrpNy = Itn(Itr.Properties)
End Function

Function ItrzLines(Lines$)
Asg Itr(SplitCrLf(Lines$)), ItrzLines
End Function

Function NItr&(Itr)
Dim O&, V
For Each V In Itr
    O = O + 1
Next
NItr = O
End Function

Function ItrzAy(Ay)
ItrzAy = Itr(Ay)
End Function

Function ItwNm(Itr, Nm)
Dim O: For Each O In Itr
    If O.Name = Nm Then Asg O, ItwNm: Exit Function
Next
End Function

Sub Itr__Tst()
'Dim I: Set I = Itr(Array(1)) ' This will break
'Dim I: I = Itr(Array())      ' This will break
Dim I
Asg Itr(Array()), I         'The will not break
Asg Itr(Array(1)), I        'This will not break
Stop
End Sub
Function Itr(Ay)
If Si(Ay) = 0 Then Set Itr = New Collection Else Itr = Ay
End Function

Function IsAllEmpItr(Itr) As Boolean
Dim I: For Each I In Itr
    If IsEmpty(I) Then Exit Function
Next
IsAllEmpItr = True
End Function
Function SrtItr(Itr)
SrtItr = Itr(QSrt(AvzItr(Itr)))
End Function
