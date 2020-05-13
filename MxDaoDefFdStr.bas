Attribute VB_Name = "MxDaoDefFdStr"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDefFdStr."
Public Const SSoStdEle$ = "CrtDte Pk Fk Ty Nm Dte Amt Att"
Public Const DaoTynn$ = "Boolean Byte Integer Int Long Single Double Char Text Memo Attachment" ' used in TzPFld
Function FdzStdFld(F, Optional T) As DAO.Field2
Dim E$: E = ElezStdFld(F, T)
Set FdzStdFld = FdzStdEle(F, E)
End Function

Function IsStdFld(F) As Boolean
IsStdFld = ElezStdFld(F) <> ""
End Function

Function FdzStdEle(F, E) As DAO.Field2
Set FdzStdEle = F & " " & StdEleStr(E)
End Function

Function StdEleStr$(E)
Const CSub$ = CMod & "StdEleStr"
Dim O$
Select Case E
Case "CrtDte"
Case "Dte"
Case "Pk"
Case "Fk"
Case "Ty"
Case "Nm"
Case "Dte"
Case "Amt"
Case "Att"
Case Else: Thw CSub, "Given Ele is not std", "E", E
End Select
StdEleStr = O
End Function

Function ElezStdFld$(F, Optional T)
Dim R2$, R3$
R2 = Right(F, 2)
R3 = Right(F, 3)
Dim O$
Select Case True
Case F = "CrtDte":  O = "CrtDte"
Case T & "Id" = F:  O = "Pk"
Case R2 = "Id":     O = "Fk"
Case R2 = "Ty":     O = "Ty"
Case R2 = "Nm":     O = "Nm"
Case R3 = "Dte":    O = "Dte"
Case R3 = "Amt":    O = "Amt"
Case R3 = "Att":    O = "Att"
End Select
ElezStdFld = O
End Function

Function StdEleAy() As String()
StdEleAy = SyzDik(DiStdEleqEleStr)
End Function

Function DiStdEleqEleStr() As Dictionary
Static X As Boolean, Y As Dictionary
If Not X Then
    X = True
    Set Y = New Dictionary
    Y.Add "Id", ""
    Y.Add "*Id", ""
    Y.Add "*Id", ""
    Y.Add "*Id", ""
    Y.Add "*Id", ""
    Y.Add "*Id", ""
    Y.Add "*Id", ""
End If
Set DiStdEleqEleStr = Y
End Function

Function FdStr$(A As DAO.Field2)
Dim D$, R$, Z$, VTxt$, VRul, E$, S$
If A.Type = DAO.DataTypeEnum.dbText Then S = " TxtSz=" & A.Size
If A.DefaultValue <> "" Then D = "Dft=" & A.DefaultValue
If A.Required Then R = "Req"
If A.AlloZZeroLength Then Z = "AlZZLen"
If A.Expression <> "" Then E = "Epr=" & A.Expression
If A.ValidationRule <> "" Then VRul = "VRul=" & A.ValidationRule
If A.ValidationText <> "" Then VTxt = "VTxt=" & A.ValidationText
FdStr = TmlzAp(A.Name, ShtDaoTy(A.Type), R, Z, VTxt, VRul, D, E, IIf((A.Attributes And DAO.FieldAttributeEnum.dbAutoIncrField) <> 0, "Auto", ""))
End Function

Function FdzFdStr1(FdStr) As DAO.Field2
Const CSub$ = CMod & "FdzFdStr1"
Dim N$, S$ ' #Fldn and #EleStr
Dim O As DAO.Field2
AsgBrkSpc FdStr, N, S
Select Case True
Case S = "Boolean":  Set O = FdzBool(N)
Case S = "Byte":     Set O = FdzByt(N)
Case S = "Integer", S = "Int": Set O = FdzInt(N)
Case S = "Long":     Set O = FdzLng(N)
Case S = "Single":   Set O = FdzSng(N)
Case S = "Double":   Set O = FdzDbl(N)
Case S = "Currency": Set O = FdzCur(N)
Case S = "Char":     Set O = FdzChr(N)
Case HasPfx(S, "Text"): Set O = FdzTxt(N, BetBkt(S))
Case S = "Memo":     Set O = FdzMem(N)
Case S = "Attachment": Set O = FdzAtt(N)
Case S = "Time":     Set O = FdzTim(N)
Case S = "Date":     Set O = FdzDte(N)
Case Else: Thw CSub, "Invalid FdStr", "Nm Spec vdt-DaoTynn, N, S, DaoTynn"
End Select
Set FdzFdStr1 = O
End Function

Function FdzFdStr(FdStr$) As DAO.Field2
Dim F$, ShtTy$, Req As Boolean, AlZZLen As Boolean, Dft$, VTxt$, VRul$, TxtSz As Byte, Epr$
Dim L$: L = FdStr
Dim Vy(): Vy = ShfVy(L, EleLblss)
AsgAy Vy, _
    F, ShtTy, Req, AlZZLen, Dft, VTxt, VRul, TxtSz, Epr
Set FdzFdStr = Fd( _
    F, DaoTyzShtTy(ShtTy), Req, TxtSz, AlZZLen, Epr, Dft, VRul, VTxt)
End Function

Function FdStrAyFds(A As DAO.Fields) As String()
Dim F As DAO.Field
For Each F In A
    PushI FdStrAyFds, FdStr(F)
Next
End Function

Function FdzStr(FdStr$) As DAO.Field2
End Function

Private Sub FdzFdStr__Tst()
Dim Act As DAO.Field2, Ept As DAO.Field2, mFdStr$
mFdStr = "AA Int Req AlZZLen Dft=ABC TxtSz=10"
Set Ept = New DAO.Field
With Ept
    .Type = DAO.DataTypeEnum.dbInteger
    .Name = "AA"
    '.AlloZZeroLength = False
    .DefaultValue = "ABC"
    .Required = True
    .Size = 2
End With
GoSub Tst
Exit Sub
Tst:
    Set Act = FdzFdStr(mFdStr)
    If Not IsEqFd(Act, Ept) Then
        D FmtMsgNap("Act", "FdStr", FdStr(Act))
        D FmtMsgNap("Ept", "FdStr", FdStr(Ept))
    End If
    Return
End Sub

Private Sub FdzFdStr1__Tst()
Dim FdStr$
FdStr = "Txt Req Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']"
GoSub Tst
Exit Sub
Tst:
    Set Act = FdzFdStr(FdStr)
    Stop
    Return
End Sub


Function FdStrAy(D As Database, T) As String()
Dim F, Td As DAO.TableDef
Set Td = D.TableDefs(T)
For Each F In Fny(D, T)
    PushI FdStrAy, FdStr(Td.Fields(F))
Next
End Function

Function FdStrzF$(D As Database, T, F$)
FdStrzF = FdStr(FdzF(D, T, F$))
End Function
