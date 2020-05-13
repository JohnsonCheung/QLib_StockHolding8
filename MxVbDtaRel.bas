Attribute VB_Name = "MxVbDtaRel"
Option Compare Text
Option Explicit
Const CNs$ = "Dta"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbDtaRel."

Property Get SampRel() As Dictionary
Set SampRel = Rel(SampRelLy)
End Property

Property Get SampRelLy() As String()
BfrClr
BfrV "A B"
BfrV "B A"
SampRelLy = BfrLy
End Property

Property Get SamPubMthRel() As Dictionary
Set SamPubMthRel = Rel(SamPubMthRelLy)
End Property

Property Get SamPubMthRelLy() As String()
'SampMthRelLy = RelOf_MthSDNm_To_Mdn_V
End Property

Function CvRel(A) As Dictionary
Set CvRel = A
End Function

Property Get EmpRel() As Dictionary
Set EmpRel = New Dictionary
End Property

Function IsRel(A) As Boolean
IsRel = TypeName(A) = "Dictionary"
End Function

Function RelzVbl(RelVbl$) As Dictionary
Set RelzVbl = Rel(SplitVBar(RelVbl))
End Function

Function Rel(RelLy$()) As Dictionary
Dim O As New Dictionary
Dim L: For Each L In Itr(RelLy)
    PushRelLin O, L
Next
Set Rel = O
End Function

'---

Sub PushParChd(Rel As Dictionary, P, C)
If Rel.Exists(P) Then
    PushEle Rel(P), C
Else
    Rel.Add P, DiczKv(P, C)
End If
End Sub

Function ParNod(Rel As Dictionary, Par) As Dictionary
':ParNod: #Par-Node# ! It is a dictionary with Key is Parent and Value is ChdAet
If Rel.Exists(Par) Then Set ParNod = Rel(Par)
End Function
    
Sub PushRelLin(Rel As Dictionary, RelLin)
Dim Ay$(), P$, C
Ay = SyzSS(RelLin)
If Si(Ay) = 0 Then Exit Sub
P = Shf(Ay)
For Each C In Itr(Ay)
    PushParChd Rel, P, C
Next
End Sub

Function CycSetAy(Rel As Dictionary) As Dictionary()
End Function

Function IsCyc(Rel As Dictionary) As Boolean
Dim ChdAet: For Each ChdAet In Rel.Items
    Dim C: For Each C In CvAet(ChdAet).Keys
        If IsPar(Rel, C) Then IsCyc = True: Exit Function
    Next
Next
End Function

Function SrtRel(Rel As Dictionary) As Dictionary
Set SrtRel = New Dictionary
Dim P: For Each P In SrtItr(Rel.Keys)
    PushEle SrtRel, SrtDic(Rel(P))
Next
End Function

Function SwapParChd(Rel As Dictionary) As Dictionary
Set SwapParChd = New Dictionary
Dim P: For Each P In Rel.Keys
    Dim C: For Each C In CvDic(Rel(P)).Keys
        PushParChd SwapParChd, C, P
    Next
Next
End Function

Sub VcRel(Rel As Dictionary, Optional FnPfx$ = "Rel_")
BrwAy RelLy(Rel), FnPfx
End Sub

Sub BrwRel(Rel As Dictionary)
BrwAy RelLy(Rel)
End Sub

Sub LisRel(Rel As Dictionary)
DmpAy RelLy(Rel)
End Sub

Function CloneRel(Rel As Dictionary) As Dictionary
Set CloneRel = New Dictionary
Dim P: For Each P In Rel.Keys
    CloneRel.Add P, CloneAet(CvDic(Rel(P)))
Next
End Function

Sub DmpRel(Rel As Dictionary)
D RelLy(Rel)
End Sub

Function CycRel(Rel As Dictionary) As String()
':CycRel: :RelLinAy #Cyclic-Rel#
End Function

Function RelLy(Rel As Dictionary) As String()
Dim P: For Each P In Rel.Keys
    PushI RelLy, RelLin(Rel, P)
Next
End Function

Function IsEqRel(Rel1 As Dictionary, Rel2 As Dictionary) As Boolean
Stop '
'If Not IsEqItr(A.Rel.Keys, B.Rel.Keys) Then Exit Function
'Dim K
'For Each K In Rel_ParAet(A)
'    If Not Aet_IsEq(A.Rel(K), B.Rel(K)) Then Exit Function
'Next
'Rel_IsEq = True
End Function

Sub ChkRelEq(Rel1 As Dictionary, Rel2 As Dictionary, Optional Msg$ = "Two rel are diff", Optional N1$ = "Rel-B", Optional N2$ = "Rel-B")
Const CSub$ = CMod & "ChkRelEq"
If IsEqRel(Rel1, Rel2) Then Exit Sub
Dim O$()
PushI O, Msg
PushI O, FmtQQ("?-ParCnt(?) / ?-ParCnt(?)", N1, Rel1.Count, N2, Rel2.Count)
PushI O, N1 & " --------------------"
PushIAy O, RelLy(Rel1)
PushI O, N2 & " --------------------"
PushIAy O, RelLy(Rel2)
ChkNoEr O, CSub
End Sub

Function ErzRel$(Rel As Dictionary)
Const C$ = "Given Rel "
Select Case True
Case IsNothing(Rel): ErzRel = C & "is nothing"
Case Else
    Dim P: For Each P In Rel.Keys
        Dim V: V = Rel(P)
        If Not IsDic(V) Then
            ErzRel = "has parent[" & P & "] whose ChdAet is not a dictionary, but typename=[" & TypeName(V) & "]"
            Exit Function
        End If
    Next
End Select
End Function

Sub ChkRelVdt(Rel As Dictionary, Fun$)
Dim Er$: Er = ErzRel(Rel): If Si(Er) = 0 Then Exit Sub
Thw Fun, "Given Rel is not a valid", "Er", Er
End Sub

Function NItmzRel&(Rel As Dictionary)
NItmzRel = ItmAet(Rel).Count
End Function

Function IsLeaf(Rel As Dictionary, Itm) As Boolean
IsLeaf = True
If IsNoChdPar(Rel, Itm) Then Exit Function
If Not IsPar(Rel, Itm) Then Exit Function
IsLeaf = False
End Function

Function IsNoChdPar(Rel As Dictionary, P) As Boolean
If Not IsPar(Rel, P) Then Exit Function
IsNoChdPar = CvDic(Rel(P)).Count = 0
End Function

Function IsPar(Rel As Dictionary, P) As Boolean
IsPar = Rel.Exists(P)
End Function

Function ItmAet(Rel As Dictionary) As Dictionary
Set ItmAet = New Dictionary
PushItrzAet ItmAet, Rel.Keys
Dim IChdAet: For Each IChdAet In Rel.Items
    PushSet ItmAet, CvAet(IChdAet)
Next
End Function
Sub ChkNoCyc(Rel As Dictionary, Fun$)
If IsCyc(Rel) Then Thw Fun, "Given Rel is Cyc", "CycRel Rel", CycRel(Rel), RelLy(Rel)
End Sub
Function ItmAetInDpdOrd(Rel As Dictionary) As Dictionary
Const CSub$ = CMod & "ItmAetInDpdOrd"
'Return itms in Rel in dependant order. Throw er if there is cyclic
'Example: A B C D
'         C D E
'         E X
'Return: B D X E C A
ChkNoCyc Rel, CSub
Dim O As New Dictionary, J%, MRel As Dictionary, Leaves As Dictionary
Set MRel = CloneRel(Rel)
Do
    J = J + 1: If J > 1000 Then Thw CSub, "looping to much"
    Set Leaves = MRel.LeafSet
    If Leaves.IsEmp Then
        If MRel.NPar > 0 Then
            Thw CSub, "Cyclic relation is found so far.  No leaves but there is remaining Rel", _
            "Turn-Cnt [Orginal rel] [Dpd itm found] [Remaining relation not solved]", _
            J, RelLy(Rel), O.Ln, RelLy(MRel)
        End If
        Set ItmAetInDpdOrd = O
        Exit Function
    End If
    O.PushSet Leaves
    MRel.RmvAllLeaf
    O.PushSet MRel.NoChdPar
    MRel.RmvNoChdPar
Loop
Set ItmAetInDpdOrd = O
End Function

Function ParAetzRel(Rel As Dictionary) As Dictionary
Set ParAetzRel = AetzItr(Rel.Keys)
End Function

Function LeafSet(Rel As Dictionary) As Dictionary
Set LeafSet = New Dictionary
Dim Itm: For Each Itm In ItmAet(Rel).Keys
    If IsLeaf(Rel, Itm) Then PushEle LeafSet, Itm
Next
End Function

Function NoChdParAet(Rel As Dictionary) As Dictionary
Set NoChdParAet = New Dictionary
Dim P: For Each P In Rel.Keys
    If IsNoChdPar(Rel, P) Then NoChdParAet.Add P, Empty
Next
End Function

Sub ChktPar(Rel As Dictionary, Par, Fun$)
If IsPar(Rel, Par) Then Exit Sub
Thw Fun, "Given Par is not a parent", "Rel Par", RelLy(Rel), Par
End Sub

Function ChdAetzPar(Rel As Dictionary, P) As Dictionary
Set ChdAetzPar = Rel(P)
End Function

Function HasChd(Rel As Dictionary, P, C) As Boolean
If Not IsPar(Rel, P) Then Exit Function
HasChd = ChdAetzPar(Rel, P).Exists(C)
End Function

Function RelLin$(Rel As Dictionary, P)
':RelLin: :Linss #Rel-Ln# ! First Term is parent, rest is chd
If Not Rel.Exists(P) Then Exit Function
RelLin = P & " " & LinsszItr(Rel(P).Keys)
End Function

Function RmvChdAy&(ORel As Dictionary, P, ChdAy)
If Not IsPar(ORel, P) Then Exit Function
Dim C: For Each C In Itr(ChdAy)
    If RmvChd(ORel, P, C) Then
        RmvChdAy = RmvChdAy + 1
    End If
Next
End Function

Private Sub RmvChd__Tst()

End Sub

Function RmvChd(ORel As Dictionary, P, C) As Boolean
If Not HasChd(ORel, P, C) Then Exit Function
CvDic(ORel(P)).Remove C
RmvChd = True
End Function

Function ChdAet(Rel As Dictionary) As Dictionary
Set ChdAet = New Dictionary
Dim IChdAet: For Each IChdAet In Rel.Items
    PushSet ChdAet, CvAet(IChdAet)
Next
End Function

Function LeafAv(Rel As Dictionary) As Variant()
LeafAv = AvzAet(LeafSet(Rel))
End Function
Function RmvAllLeaf(ORel As Dictionary) As Boolean
Dim LeafAv1(): LeafAv1 = LeafAv(ORel): If Si(LeafAv1) = 0 Then Exit Function
Dim P: For Each P In ORel.Keys
    RmvChdAy ORel, P, LeafAv1
Next
RmvAllLeaf = True
End Function

Function RmvNoChdPar&(ORel As Dictionary)
Dim O&
Dim P: For Each P In ORel.Keys
    If IsNoChdPar(ORel, P) Then
        ORel.Remove P
        O = O + 1
    End If
Next
RmvNoChdPar = O
End Function

Property Get SampRel1() As Dictionary
Set SampRel1 = RelzVbl("B C D | D E | X")
End Property

Private Sub ItmAet__Tst()
Dim Act As Dictionary, Ept As Dictionary, Rel As Dictionary
Set Ept = AetzSS("A B C D E")
Set Rel = RelzVbl("A B C | B D E | C D")
GoSub Tst
Exit Sub
Tst:
    Set Act = ItmAet(Rel)
    C
    Return
End Sub

Private Sub ItmAetInDpdOrd__Tst()
Dim Act As Dictionary, Ept As Dictionary
Dim Rel1 As Dictionary
GoSub T1
'GoSub T2
Exit Sub
T1:
    Set Ept = AetzSS("C E X D B")
    Set Rel1 = RelzVbl("B C D | D E | X")
    GoSub Tst
    Return
'
T2:
    Dim X$()
    PushI X, "MVb"
    PushI X, "MIde MVb MXls MAcs"
    PushI X, "MXls MVb"
    PushI X, "MDao MVb MDta"
    PushI X, "MAdo MVb"
    PushI X, "MAdoX MVb"
    PushI X, "MApp  MVb"
    PushI X, "MDta  MVb"
    PushI X, "MTp   MVb"
    PushI X, "MSql  MVb"
    PushI X, "AStkShpCst MVb MXls MAcs"
    PushI X, "MAcs  MVb MXls"
    Set Rel1 = Rel(X)
    Set Ept = AetzSS("MVb MIde MXls MDao MAdo MAdoX MApp MDta MTp MSql AStkShpCst MAcs ")
    GoSub Tst
    Return
Tst:
    Set Act = ItmAetInDpdOrd(Rel1)
    If Not IsEqAet(Act, Ept) Then Stop
    Return
End Sub

Function SngChdParAet(Rel As Dictionary) As Dictionary
Set SngChdParAet = New Dictionary
Dim P: For Each P In Rel.Keys
    If ChdAetzPar(Rel, P).Count = 1 Then SngChdParAet.PushItm P
Next
End Function

Function RelzS12y(A() As S12) As Dictionary
Set RelzS12y = Rel(LyzS12y(A))
End Function
