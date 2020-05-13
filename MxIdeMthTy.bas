Attribute VB_Name = "MxIdeMthTy"
Option Explicit
Option Compare Text
Const CNs$ = "Mth.Ln"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthTy."
Const C_Fun$ = "Function"
Const C_Prp$ = "Property"
Const C_Sub$ = "Sub"

Const C_Get$ = "Get"
Const C_Set$ = "Set"
Const C_Let$ = "Let"

Const C_PrpGet$ = C_Prp + " " + C_Get
Const C_PrpLet$ = C_Prp + " " + C_Let
Const C_PrpSet$ = C_Prp + " " + C_Set

Private Sub MthKd__Tst()
Dim A$
Ept = "Property": A = "Property Get": GoSub Tst
Ept = "Property": A = "Property Get":         GoSub Tst
Ept = "Property": A = " Property Get":        GoSub Tst
Ept = "Property": A = "Friend Property Get":  GoSub Tst
Ept = "Property": A = "Friend  Property Get": GoSub Tst
Ept = "":         A = "FriendProperty Get":   GoSub Tst
Exit Sub
Tst:
    Act = MthKd(A)
    C
    Return
End Sub

Function MthKd$(MthTy$)
Select Case MthTy
Case C_PrpGet: MthKd = C_Prp
Case C_PrpSet: MthKd = C_Prp
Case C_PrpLet: MthKd = C_Prp
Case C_Sub, C_Fun: MthKd = MthTy
End Select
End Function

Function MthKdzL$(Ln)
MthKdzL = MthKd(MthTy(Ln))
End Function

Property Get ShtMthKdAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy("Fun", "Sub", "Prp")
ShtMthKdAy = X
End Property

Property Get PrpTyAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy(C_Get, C_Set, C_Let)
PrpTyAy = X
End Property

Property Get MthTyAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy(C_Fun, C_Sub, C_PrpGet, C_PrpLet, C_PrpSet)
MthTyAy = X
End Property
Property Get ShtMthTyAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy("Fun", "Sub", "Get", "Set", "Let")
ShtMthTyAy = X
End Property

Property Get MthKdAy() As String()
Static X$()
If Si(X) = 0 Then X = Sy(C_Fun, C_Sub, C_Prp)
MthKdAy = X
End Property

Function MthTy$(Ln) 'One Of [Function Sub [Property Get] [Propery Let] [Property Set]]
MthTy = PfxzAySpc(RmvMdy(Ln), MthTyAy)
End Function

Private Sub MthTy__Tst()
Dim O$(), L
For Each L In SrczMdn("Fct")
    Push O, MthTy(CStr(L)) & "." & L
Next
BrwAy O
End Sub

Function MthTyzSht$(ShtMthTy)
Dim O$
Select Case ShtMthTy
Case "Get": O = "Property Get"
Case "Set": O = "Property Set"
Case "Let": O = "Property Let"
Case "Fun": O = "Function"
Case "Sub": O = "Sub"
Case Else:  Thw CSub, "Given ShtMthTy is invalid", "ShtMthTy Invalid-ShtMthTy", ShtMthTy, "Get Set Let Fun Sub"
End Select
MthTyzSht = O
End Function


Private Sub ShtMthTyzLin__Tst()
GoSub Z
Exit Sub
Z:
    Dim O$(), I, Ln
    For Each I In MthlnyV
        Ln = I
        PushI O, ShtMthTyzLin(Ln)
    Next
    Brw O
    Return
End Sub
Function ShtMthTyzLin(Ln)
ShtMthTyzLin = ShtMthTy(TakMthTy(RmvMdy(Ln)))
End Function

Function ShtMthKdzShtMthTy$(ShtMthTy$)
Dim O$
Select Case ShtMthTy
Case "Get": O = "Prp"
Case "Set": O = "Prp"
Case "Let": O = "Prp"
Case "Fun": O = "Fun"
Case "Sub": O = "Sub"
Case Else: O = "???"
End Select
ShtMthKdzShtMthTy = O
End Function

Function ShtMthKd$(MthKd)
Dim O$
Select Case MthKd
Case "Property": O = "Prp"
Case "Function": O = "Fun"
Case "Sub":      O = "Sub"
Case Else: O = "???"
End Select
ShtMthKd = O
End Function

Function MthKdzTy$(MthTy)
Select Case MthTy
Case "Function", "Sub": MthKdzTy = MthTy
Case "Property Get", "Property Let", "Property Set": MthKdzTy = "Property"
End Select
End Function

Function IsMthTy(S) As Boolean
IsMthTy = HasEle(MthTyAy, S)
End Function

Function IsPrpTy(S) As Boolean
IsPrpTy = HasEle(PrpTyAy, S)
End Function

Function VbMdyzSht$(ShtMdy)
Dim O$
Select Case ShtMdy
Case "Pub": O = "Public"
Case "Prv": O = "Private"
Case "Frd": O = "Friend"
Case ""
Case Else: Stop
End Select
VbMdyzSht = O
End Function
Function IsRetVal(ShtMthTy$) As Boolean
Select Case ShtMthTy
Case "Get", "Fun": IsRetVal = True
End Select
End Function
Function ShtMthTy$(MthTy)
Dim O$
Select Case MthTy
Case "Property Get": O = "Get"
Case "Property Set": O = "Set"
Case "Property Let": O = "Let"
Case "Function":     O = "Fun"
Case "Sub":          O = "Sub"
End Select
ShtMthTy = O
End Function
