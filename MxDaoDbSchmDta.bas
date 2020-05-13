Attribute VB_Name = "MxDaoDbSchmDta"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDaoDbSchmDta."

Private Function SchmDta(A As SchmSrc) As SchmDta
With SchmDta
'    .Sk = SqyCrtSkzSchmSrc(A)
'    .Pkq = SqyCrtPk(PkTnyzsmsTblAy(A.Tbl))
'    Set .TblDesDi = TblDesDi(A.TblDes)
'    Set .TblFDesDi = TFqDesDi(A)
'    .TdAy = TdAy(A)
End With
End Function

Private Function TFqDesDi(A As SchmSrc)

End Function
Private Function PkTnyzsmsTblAy(A() As SmsTbl) As String()

End Function
Private Function EFqEsDi() As Dictionary

End Function
Private Function TblDesDi(A() As SmsTblDes) As Dictionary

End Function

Private Function TblFDesDi(F() As smdFldDes, TF() As smdTFDes) As Dictionary

End Function

Private Function PkTny(T_T$(), T_FnyAy()) As String()
Dim J%: For J = 0 To UB(T_T)
    Dim Fny$(): Fny = T_FnyAy(J)
    If T_T(J) & "Id" = Fny(0) Then PushI PkTny, T_T(J)
Next
End Function

Function IxzLikssAy%(Itm, LikssAy$())
Dim I%, Likss: For Each Likss In LikssAy
    If ItmInLikAy(Itm, SyzSS(Likss)) Then IxzLikssAy = I: Exit Function
    I = I + 1
Next
IxzLikssAy = -1
End Function

Function ItmInLikAy(Itm, LikAy$()) As Boolean
Dim Lik: For Each Lik In LikAy
    If Itm Like Lik Then ItmInLikAy = True: Exit Function
Next
End Function

Private Function DiFqE(AllFny$(), EF_E$(), EF_FldLikAy()) As Dictionary
Set DiFqE = New Dictionary
Dim F: For Each F In AllFny
    Stop
    Dim Ix%: 'Ix = IxzLikssAy(F, EF_FldLikAy)
    Dim E$: E = EF_E(Ix)
    DiFqE.Add F, E
Next
End Function

Private Function TFqEsDi(A As SchmSrc) As Dictionary
Dim AllFny$(): 'AllFny = AwDis(AyzAyOfAy(T_Fny))
Dim FqE As Dictionary: 'Set FqE = DiFqE(AllFny, EF_E, EF_FldLikAy)
Dim EqEs As Dictionary: ' Set EqEs = DiczAy2(E_E, E_EleStr)
Dim FqEs As Dictionary: Set FqEs = ChnDic(FqE, EqEs)
Set TFqEsDi = DiwAy(FqEs, AllFny)
End Function

Private Function TdAy(A As SchmSrc) As DAO.TableDef()
Dim Tny$(), FnyAy(), TFqEsDi As Dictionary
Dim T, J%: For Each T In Itr(Tny)
    Dim Fny$(): Fny = FnyAy(J)
    PushObj TdAy, Td_(T, FnyAy(J), TFqEsDi)
    J = J + 1
Next
End Function

Private Function Td_(T, Fny, TFqEsDi As Dictionary) As DAO.TableDef
Set Td_ = TdzTFdAy(T, FdAy_(T, CvSy(Fny), TFqEsDi))
End Function

Private Function FdAy_(T, Fny$(), TFqEsDi As Dictionary) As DAO.Field()
Dim F: For Each F In Fny
    PushObj FdAy_, FdzEleStr(F, TFqEsDi(T & "." & F))
Next
End Function

Private Sub CrtSchm__Tst()
Dim D As Database, Schm$()
GoSub T1
Exit Sub

T1:
    Set D = TmpDb
    Schm = SampSchm(1)
    GoTo Tst
Tst:
    CrtSchm D, Schm
    Return
End Sub

Private Function SmdSkzS(A As SchmSrc) As SmdSk()
Dim M As SmsTbl
Dim J%: For J = 0 To SmsTblUB(A.Tbl)
    M = A.Tbl(J)
    If Si(M.SkFny) > 0 Then
        PushSmdSk SmdSkzS, SmdSk(M.Tbn, M.SkFny)
    End If
Next
End Function
Private Sub PushSmdSk(O() As SmdSk, M As SmdSk)

End Sub
Private Function SmdSk(Tbn, SkFny$()) As SmdSk
End Function

Private Function Sk(A() As SmsTbl) As TFny()

End Function

Private Function HasPk(A As SmsTbl) As Boolean
HasPk = A.Tbn & "Id" = A.Fny(0)
End Function
Function SkFnyAy(A As SchmSrc) As Variant()
Stop
End Function


Function TnyzWiSk(A As SchmSrc) As String()
Stop
End Function
Function TnyzsmsTblAy(A() As SmsTbl) As String()
Dim J%: For J = 0 To SmsTblUB(A)
    PushS TnyzsmsTblAy, A(J).Tbn
Next
End Function
Function TnyzWiPk(A As SchmSrc) As String()
Dim M As SmsTbl
Dim J%: For J = 0 To SmsTblUB(A.Tbl)
    M = A.Tbl(J)
    If HasPkzsmsTbl(M) Then PushI TnyzWiPk, M.Tbn
Next
End Function
Function HasPkzsmsTbl(A As SmsTbl) As Boolean
HasPkzsmsTbl = A.Tbn & "Id" = A.Fny(0)
End Function
Sub AsgBrk1Dotlny(Dotlny$(), OBefDot$(), OAftDot$())
Dim U&: U = UB(Dotlny)
ReDim OBefDot(U)
ReDim OAftDot(U)
Dim DotNm, J&: For Each DotNm In Itr(Dotlny)
    With Brk1Dot(DotNm)
        OBefDot(J) = .S1
        OAftDot(J) = .S2
    End With
    J = J + 1
Next
End Sub

Private Function T_Fny(Tny$(), FssAy$()) As Variant()
Dim Fss, J%: For Each Fss In Itr(FssAy)
    PushI T_Fny, SyzSS(Replace(Fss, "*", Tny(J)))
Next
End Function
