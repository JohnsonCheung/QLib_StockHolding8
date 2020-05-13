Attribute VB_Name = "MxVbAyBool"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ay"
Const CMod$ = CLib & "MxVbAyBool."
Enum eBoolOp: eOpEq: eOpNe: eOpAnd: eOpOr: End Enum
Enum eEqNe:  eEqnNe: eEqnEq: End Enum
Enum eOrAnd: eOraOr: eOraAnd: End Enum
Public Const eBoolOpSS$ = "OpEq OpNe OpAnd OpOr"
Public Const eEqnSS$ = "EqnNe EqnEq"  'Eqn:Cml #Eq-Ne#
Public Const eOraSS$ = "OraAnd OraOr" 'Ora:Cml #Or-And#
#If Doc Then
'Ebr enum-member  It use together with eXXX to form a EbrStr.  EbrStr is
'Booy boolean-array
'Bool boolean
'Op Operator
'EbrStr Enum-member-string.  It means the string for any enum-member.  The string will not have the prefix e
'Ebn    Enum-Member-name.    It explicit equal to the name defined in Enum.
'Ebn-start-e  Enum member name started with e
'Enmn-start-e Enum name started with e
#End If

Function CvOrAndStr(eOrAndStr$) As eOrAnd
Select Case True
Case eOrAndStr = "And": CvOrAndStr = eOraAnd
Case eOrAndStr = "Or": CvOrAndStr = eOraOr
Case Else: Thw CSub, "eOrAndStr error", "eOrAndStr", eOrAndStr
End Select
End Function
Function CvEqNeStr(eEqNeStr$) As eEqNe
Select Case True
Case eEqNeStr = "Eq": CvEqNeStr = eEqnEq
Case eEqNeStr = "Ne": CvEqNeStr = eEqnNe
Case Else: Thw CSub, "eEqNeSStr error", "eEqNeStr", eEqNeStr
End Select
End Function

Function EvlBooly(Booly() As Boolean, Op As eOrAnd) As BoolOpt
Select Case True
Case Op = eOraAnd: EvlBooly = SomBool(IsAllT(Booly))
Case Op = eOraOr:  EvlBooly = SomBool(IsSomT(Booly))
Case Else: Imposs CSub
End Select
End Function

Function IsOrAndStr(S$) As Boolean
Select Case True
Case S = "And", S = "Or": IsOrAndStr = True
End Select
End Function

Function IsEqNeStr(S$)
Select Case True
Case S = "Eq", S = "Ne": IsEqNeStr = True
End Select
End Function

Function Booy(TrueStr$) As Boolean() ' ele will be true if the chr of @TrueStr = T which is ignCas
Dim J%: For J = 1 To Len(TrueStr)
    PushI Booy, Mid(TrueStr, J, 1) = "T"
Next
End Function

Function eBoolOp(EbrBoolOp$) As eBoolOp: eBoolOp = IxzAy(eBoolOpSy, EbrBoolOp): End Function

Function IfStrT$(B As Boolean, TrueStr$) 'If @B is true, return @TrueStr else return ""
If B Then IfStrT = TrueStr
End Function
Function IfStrF$(B As Boolean, FalseStr$) 'If @B is false, return @FalseStr else return ""
If B = False Then IfStrF = FalseStr
End Function

Function IsAllF(A() As Boolean) As Boolean
Dim I
For Each I In A
    If I Then Exit Function
Next
IsAllF = True
End Function

Function IsAllT(A() As Boolean) As Boolean
Dim I
For Each I In A
    If Not I Then Exit Function
Next
IsAllT = True
End Function

Function IsStrAndOr(A$) As Boolean
Select Case UCase(A)
Case "AND", "OR": IsStrAndOr = True
End Select
End Function

Function IsSomF(A() As Boolean) As Boolean
Dim B: For Each B In A
    If Not B Then IsSomF = True: Exit Function
Next
End Function

Function IsSomT(A() As Boolean) As Boolean
Dim B: For Each B In A
    If B Then IsSomT = True: Exit Function
Next
End Function

Function IsEbrEqNe(eEqNeEle$) As Boolean: IsEbrEqNe = HasEle(eEqnSS, eEqNeEle): End Function

Function CvBoolOp(eBoolOpStr$) As eBoolOp
CvBoolOp = IxzAy(eBoolOpSy, eBoolOpStr)
End Function

Function eBoolOpSy() As String()
Static X$(): If IsNothing(X) Then X = Sy(eBoolOpSS)
eBoolOpSy = X
End Function
