Attribute VB_Name = "MxVbDtaDicOp"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Dic"
Const CMod$ = CLib & "MxVbDtaDicOp."

Function AddKpfx(A As Dictionary, Kpfx$) As Dictionary
Dim O As New Dictionary
Dim K: For Each K In O.Keys
    O.Add Kpfx & K, A(K)
Next
Set AddKpfx = O
End Function

Function CvDic(A) As Dictionary
Set CvDic = A
End Function

Function AetzDicKey(A As Dictionary) As Dictionary
Set AetzDicKey = AetzItr(A.Keys)
End Function

Function CvDicAy(A) As Dictionary()
CvDicAy = A
End Function

Function AddDicAy(A As Dictionary, Dy() As Dictionary) As Dictionary
Set AddDicAy = CloneDic(A)
Dim J%
For J = 0 To UB(Dy)
   PushDic AddDicAy, Dy(J)
Next
End Function

Function IupDic(A As Dictionary, By As Dictionary) As Dictionary 'Return New dictionary from A-Dic by Ins-or-upd By-Dic.  Ins: if By-Dic has key and A-Dic. _
Upd: K fnd in both, A-Dic-Val will be replaced by By-Dic-Val
Dim O As New Dictionary, K
For Each K In A.Keys
    If By.Exists(K) Then
        O.Add K, By(K)
    Else
        O(K) = By(K)
    End If
Next
Set IupDic = O
End Function

Function AddDicKeyPfx(A As Dictionary, Pfx) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add Pfx & K, A(K)
Next
Set AddDicKeyPfx = O
End Function

Sub DicAddOrUpd(A As Dictionary, K$, V, Sep$)
If A.Exists(K) Then
    A(K) = A(K) & Sep & V
Else
    A.Add K, V
End If
End Sub

Function DicAyKy(A() As Dictionary) As Variant()
Dim I
For Each I In Itr(A)
   PushNDupAy DicAyKy, CvDic(I).Keys
Next
End Function

Function CloneDic(A As Dictionary) As Dictionary
Set CloneDic = New Dictionary
Dim K: For Each K In A.Keys
    CloneDic.Add K, A(K)
Next
End Function

Function DrDicKy(A As Dictionary, Ky$()) As Variant()
Dim O(), I, J&
ReDim O(UB(Ky))
For Each I In Ky
    If A.Exists(I) Then
        O(J) = A(I)
    End If
    J = J + 1
Next
DrDicKy = O
End Function

Function DyzDotlny(Dotlny$()) As Variant()
Dim I, Ln
For Each I In Itr(Dotlny)
    Ln = I
    PushI DyzDotlny, SplitDot(Ln)
Next
End Function

Function IntersectDic(A As Dictionary, B As Dictionary) As Dictionary ' ret Sam Key Sam Val
Dim O As New Dictionary
If A.Count = 0 Then GoTo X
If B.Count = 0 Then GoTo X
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            O.Add K, A(K)
        End If
    End If
Next
X: Set IntersectDic = O
End Function

Function KeyAet(A As Dictionary) As Dictionary
Set KeyAet = AetzItr(A.Keys)
End Function

Function VzDicIfKyJn$(A As Dictionary, Ky, Optional Sep$ = vb2CrLf)
Dim O$(), K
For Each K In Itr(Ky)
    If A.Exists(K) Then
        PushI O, A(K)
    End If
Next
VzDicIfKyJn = Join(O, Sep)
End Function

Function SyzDicKy(Dic As Dictionary, Ky$()) As String()
Const CSub$ = CMod & "SyzDicKy"
Dim K
For Each K In Itr(Ky)
    If Dic.Exists(K) Then Thw CSub, "K of Ky not in Dic", "K Ky Dic", K, Ky, Dic
    PushI SyzDicKy, Dic(K)
Next
End Function

Function LineszDic$(A As Dictionary)
LineszDic = JnCrLf(FmtDiczKLines(A))
End Function

Function FmtDiczKLines(A As Dictionary) As String()
Dim K: For Each K In A.Keys
    Push FmtDiczKLines, LyzKLines(K, A(K))
Next
End Function

Function LyzKLines(K, Lines$) As String()
Dim Ly$(): Ly = SplitCrLf(Lines)
Dim J&: For J = 0 To UB(Ly)
    Dim Ln
        Ln = Ly(J)
        If FstChr(Ln) = " " Then Ln = "~" & RmvFstChr(Ln)
    Push LyzKLines, K & " " & Ln
Next
End Function

Function MgeDic(A As Dictionary, PfxSsl$, ParamArray DicAp()) As Dictionary
Dim Av(): Av = DicAp
Dim Ny$()
   Ny = SyzSS(PfxSsl)
   Ny = AmAddSfx(Ny, "@")
If Si(Av) <> Si(Ny) Then Stop
Dim Dy() As Dictionary
Dim D As Dictionary
   Dim J%
   For J = 0 To UB(Ny)
       Set D = Av(J)
       Push Dy, AddDicKeyPfx(A, Ny(J))
   Next
Set MgeDic = AddDicAy(A, Dy)
End Function

Sub BrwKSet(KSet As Dictionary)
BrwDrs DrszKSet(KSet)
End Sub

Function DrszKSet(KSet As Dictionary) As Drs
Dim K, Dy(), Sset As Dictionary, V
For Each K In KSet.Keys
    Set Sset = KSet(K)
    If Sset.Count = 0 Then
        PushI Dy, Array(K, "#EmpSet#")
    Else
        For Each V In Sset.Keys
            PushI Dy, Array(K, V)
        Next
    End If
Next
DrszKSet = DrszFF("K V", Dy)
End Function

Function HasKSet(KSet As Dictionary, K, Aet As Dictionary) As Boolean
'Fm KSet : KSet if a dictionary with value is Aet.
'Ret     : True if KSet has such Key-K and Val-Set-Aet  @@
If KSet.Exists(K) Then
    Dim ISet As Dictionary: Set ISet = KSet(K)
    HasKSet = ISet.IsEq(Aet)
End If
End Function

Function KSetzDif(KSet1 As Dictionary, KSet2 As Dictionary)
'Ret : KSet from KSet1 where not found in KSet2 (Not found means K is not found or K is found but V is dif @@
Set KSetzDif = New Dictionary
Dim K: For Each K In KSet1.Keys
    Dim V As Dictionary: Set V = KSet1(K)
    Dim Has As Boolean: Has = HasKSet(KSet2, K, V)
    If Not Has Then
        KSetzDif.Add K, V
    End If
Next
End Function

Function AddKv(ODic As Dictionary, K, V) As Boolean
If Not ODic.Exists(K) Then ODic.Add K, V: AddKv = True
End Function

Function RmvKey(ODic As Dictionary, K) As Boolean
If ODic.Exists(K) Then ODic.Remove K: RmvKey = True
End Function

Function MinusDic(A As Dictionary, B As Dictionary) As Dictionary
'Ret those Ele in A and not in B
If B.Count = 0 Then Set MinusDic = CloneDic(A): Exit Function
Dim O As New Dictionary, K
For Each K In A.Keys
   If Not B.Exists(K) Then O.Add K, A(K)
Next
Set MinusDic = O
End Function

Function DicSelIntozAy(A As Dictionary, Ky$()) As Variant()
Dim O()
Dim U&: U = UB(Ky)
ReDim O(U)
Dim J&
For J = 0 To U
   If Not A.Exists(Ky(J)) Then Stop
   O(J) = A(Ky(J))
Next
DicSelIntozAy = O
End Function

Function DicSelIntoSy(A As Dictionary, Ky$()) As String()
DicSelIntoSy = SyzAy(DicSelIntozAy(A, Ky))
End Function

Function SyzDik(A As Dictionary) As String()
SyzDik = SyzItr(A.Keys)
End Function

Function SwapKv(StrDic As Dictionary) As Dictionary
Set SwapKv = New Dictionary
Dim K: For Each K In StrDic.Keys
    SwapKv.Add StrDic(K), K
Next
End Function

Function KeyzLikAyDic(Dic As Dictionary, Itm$) As String()
Dim K, LikAy$()
For Each K In Dic.Keys
    LikAy = Dic(K)
    If HitLikAy(Itm, LikAy) Then
        KeyzLikAyDic = K
        Exit Function
    End If
Next
End Function

Function WbzDiNmqLines(DiNmqLines As Dictionary) As Workbook 'Assume each dic keys is name and each value is lines. _
create a new Wb with worksheet as the dic key and the lines are break to each cell of the sheet
Dim A As Dictionary: Set A = DiNmqLines
Dim K, ThereIsSheet1 As Boolean
Dim O As Workbook: Set O = NwWb
Dim Ws As Worksheet
For Each K In A.Keys
    If K = "Sheet1" Then
        ThereIsSheet1 = True
    Else
        Set Ws = O.Sheets.Add
        Ws.Name = K
    End If
    Ws.Range("A1").Value = SqvzLines(A(K))
Next
X: Set WbzDiNmqLines = O
End Function

Function DiAqCzOuter(DiAqB As Dictionary, DiBqC As Dictionary) As Dictionary
Dim A, B, C
Set DiAqCzOuter = New Dictionary
For Each A In DiAqB.Keys
    B = DiAqB(A)
    If DiBqC.Exists(B) Then
        DiAqCzOuter.Add A, C
    Else
        DiAqCzOuter.Add A, Empty
    End If
Next
End Function
Function DicAC(DicAB As Dictionary, DicBC As Dictionary) As Dictionary
Dim A, B, C
Set DicAC = New Dictionary
For Each A In DicAB.Keys
    B = DicAB(A)
    If DicBC.Exists(B) Then
        DicAC.Add A, DicBC(B)
    End If
Next
End Function

Function DicAzDifVal(A As Dictionary, B As Dictionary) As Dictionary
Set DicAzDifVal = New Dictionary
Dim K, V
For Each K In A.Keys
    If B.Exists(K) Then
        V = A(K)
        If V <> B(K) Then DicAzDifVal.Add K, V
    End If
Next
End Function
Sub SetKv(O As Dictionary, K, V)
If O.Exists(K) Then
    Asg O(K), _
        V
Else
    O.Add K, V
End If
End Sub

Sub PushDic(O As Dictionary, A As Dictionary)
Const CSub$ = CMod & "PushDic"
If IsNothing(O) Then
    Set O = CloneDic(A)
    Exit Sub
End If
Dim K
For Each K In A.Keys
    AddKv O, K, A(K)
Next
End Sub

Function ChnDic(DiAqB As Dictionary, DiBqC As Dictionary) As Dictionary
Const CSub$ = CMod & "ChnDic"
':Ret :DiAqC  ! Thw Er if A->B and B not in fnd in DiBqC @@
Dim A As Dictionary: Set A = DiAqB
Dim B As Dictionary: Set B = DiBqC
Set ChnDic = New Dictionary
Dim KA: For Each KA In A.Keys
    If Not B.Exists(A(KA)) Then
        Thw CSub, "KeyA has ValB which not found in DiBqC", "KeyA ValB DiAqB DiBqC", KA, A(KA), FmtDic(DiAqB), FmtDic(DiBqC)
    End If
    ChnDic.Add KA, B(A(KA))
Next
End Function

Sub PushItmzDiT1qLy(A As Dictionary, K, Itm)
Dim M$()
If A.Exists(K) Then
    M = A(K)
    PushI M, Itm
    A(K) = M
Else
    A.Add K, Sy(Itm)
End If
End Sub

Sub ChktSuperAyDiT1qLy(A As Dictionary, Fun$)
If Not IsDiiSy(A) Then Thw Fun, "Given dictionary is not DiT1qLy, all key is string and val is Sy", "Give-Dictionary", FmtDic(A)
End Sub

Function AddSfxToVal(D As Dictionary, Sfx$) As Dictionary
Dim O As New Dictionary
Dim K: For Each K In D.Keys
    Dim V$: V = D(K) & Sfx
    O.Add K, V
Next
Set AddSfxToVal = O
End Function

Sub AddKvIfNB(ODic As Dictionary, K, IfNB_S$)
If IfNB_S = "" Then Exit Sub
AddKv ODic, K, IfNB_S
End Sub

Sub SetDicValAsItr(O As Dictionary)
Dim K: For Each K In O.Keys
    O(K) = Itr(O(K))
Next
End Sub

Function AddSnoToKey(A As Dictionary) As Dictionary
Dim O As New Dictionary, J&
Dim N%: N = NDig(A.Count)
Dim K: For Each K In A.Keys
    O.Add AliR(J, N) & " " & K, A(K)
    J = J + 1
Next
Set AddSnoToKey = O
End Function

Function SrtDic(A As Dictionary, Optional By As eOrd) As Dictionary
If A.Count = 0 Then Set SrtDic = New Dictionary: Exit Function
Dim O As New Dictionary
Dim Srt: Srt = QSrt(A.Keys, By)
Dim K: For Each K In Srt
   O.Add K, A(K)
Next
Set SrtDic = O
End Function
Function MgeSKy(A As Dictionary, B As Dictionary) As String()
MgeSKy = AwDis(AddSy(SKy(A), SKy(B)))
End Function

Function SyzDii(A As Dictionary) As String()
SyzDii = SyzItr(A.Items)
End Function

Function JnStrVy$(StrDic As Dictionary, Optional Sep$ = vb2CrLf)
JnStrVy = Jn(SVy(StrDic), Sep)
End Function
