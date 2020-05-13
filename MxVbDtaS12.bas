Attribute VB_Name = "MxVbDtaS12"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CNs$ = "S12"
Const CMod$ = CLib & "MxVbDtaS12."
Type S12: S1 As String: S2 As String: End Type 'Deriving(Ay Ctor Opt)
Type S12Opt: Som As Boolean: S12 As S12: End Type
Type S3: A As String: B As String: C As String: End Type
#If Doc Then
'Function S12(S1, S2) As S12
'With S12
'    .S1 = S1
'    .S2 = S2
'End With
'End Function
#End If

'**S12y-Op
Function RmvBlnkS2(S() As S12) As S12()
Dim J%: For J = 0 To S12UB(S)
    If Trim(S(J).S2) <> "" Then PushS12 RmvBlnkS2, S(J)
Next
End Function
'**Samp-S12
Function SampS12y() As S12()
Dim O() As S12
PushS12 O, S12("sldjflsdkjf", "lksdjf")
PushS12 O, S12("sldjflsdkjf", "lksdjf")
PushS12 O, S12("sldjf", "lksdjf")
PushS12 O, S12("sldjdkjf", "lksdjf")
SampS12y = O
End Function

Function S12zTRst(A As TRst) As S12
With S12zTRst
    .S1 = A.T
    .S2 = A.Rst
End With
End Function


Sub AsgS12(A As S12, O1, O2)
O1 = A.S1
O2 = A.S2
End Sub

Sub BrwS12y(A() As S12)
BrwAy FmtS12y(A)
End Sub

Private Sub S12yzDic__Tst()
Dim A As New Dictionary
A.Add "A", "BB"
A.Add "B", "CCC"
Dim Act() As S12
Act = S12yzDic(A)
Stop
End Sub

'**Fm-S12y
Function DiczS12y(A() As S12, Optional Sep$ = vbCrLf) As Dictionary
Set DiczS12y = New Dictionary
Dim J&: For J = 0 To S12UB(A)
    PushDicS12 DiczS12y, A(J), Sep
Next
End Function
Function SqzS12y(A() As S12, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Variant()
Dim N&: N = S12Si(A)
If N = 0 Then Exit Function
Dim O(), I, R&, J&
ReDim O(1 To N, 1 To 2)
R = 2
O(1, 1) = Nm1
O(1, 2) = Nm2
For J = 0 To N - 1
    With A(J)
        O(R, 1) = .S1
        O(R, 2) = .S2
        R = R + 1
    End With
Next
SqzS12y = O
End Function

'**To-S12y
Function S12yzAyab(A, B, Optional NoTrim As Boolean) As S12()
ChkSamSi A, B, , CSub
Dim U&, O() As S12
U = UB(A)
ReDim O(U)
Dim J&: For J = 0 To U
    O(J) = S12Trim(A(J), B(J), NoTrim)
Next
S12yzAyab = O
End Function
Function S12yzDic(D As Dictionary) As S12()
Dim K: For Each K In D.Keys
    PushS12 S12yzDic, S12(K, D(K))
Next
End Function
Function S12yzColonVbl(ColonVbl$) As S12()
Dim I: For Each I In SplitVBar(ColonVbl)
    PushS12 S12yzColonVbl, BrkBoth(I, ":")
Next
End Function
Function S12yzSySep(Sy$(), Sep$, Optional NoTrim As Boolean) As S12()
Dim O() As S12
Dim U&: U = UB(Sy)
ReDim O(U)
Dim J&: For J = 0 To U
    O(J) = Brk1(Sy(J), Sep, NoTrim)
Next
S12yzSySep = O
End Function

'**S12-Prp
Function S1y(A() As S12) As String()
Dim J&: For J = 0 To S12UB(A)
   PushI S1y, A(J).S1
Next
End Function
Function S2y(A() As S12) As String()
Dim J&: For J = 0 To S12UB(A)
   Push S2y, A(J).S2
Next
End Function

'**S12-Op
Function S12yzDif(A() As S12, B() As S12) As S12()
'Ret : Subset of @A.  Those itm in @A also in @B will be exl.
Dim J&: For J = 0 To S12UB(A)
    If Not HasS12(B, A(J)) Then
        PushS12 S12yzDif, A(J)
    End If
Next
End Function
Function IsEqS12(A As S12, B As S12) As Boolean
With A
    If .S1 <> B.S1 Then Exit Function
    If .S2 <> B.S2 Then Exit Function
End With
IsEqS12 = True
End Function
Function HasS12(A() As S12, B As S12) As Boolean
Dim J&: For J = 0 To S12Si(A)
    If IsEqS12(A(J), B) Then HasS12 = True: Exit Function
Next
End Function
Function AddS2Sfx(A() As S12, S2Sfx$) As S12()
Dim O() As S12: O = A
Dim J&: For J = 0 To S12UB(A)
    O(J).S2 = O(J).S2 & S2Sfx
Next
AddS2Sfx = O
End Function
Function S12Trim(S1, S2, Optional NoTrim As Boolean) As S12
S12Trim = S12(S1, S2)
If Not NoTrim Then S12Trim = TrimS12(S12Trim)
End Function
Function TrimS12(S As S12) As S12: TrimS12 = S12(Trim(S.S1), Trim(S.S2)): End Function
Function MapS1(A() As S12, Dic As Dictionary) As S12()
Const CSub$ = CMod & "MapS1"
Dim J&: For J = 0 To S12UB(A)
    Dim M As S12: M = A(J)
    If Not Dic.Exists(M.S1) Then
        Thw CSub, "Som S1 in [S12y] not found in [Dic]", "S1-not-found S12y Dic", M.S1, FmtS12y(A), FmtDic(Dic)
    End If
    M.S1 = Dic(M.S1)
    PushS12 MapS1, M
Next
End Function
Sub WrtS12y(A() As S12, Ft$, Optional OvrWrt As Boolean)
WrtAy LyzS12y(A), Ft, OvrWrt
End Sub
Function SwapS12y(A() As S12) As S12()
Dim O() As S12: O = A
Dim J&: For J = 1 To S12UB(A)
    O(J) = SwapS12(A(J))
Next
SwapS12y = O
End Function
Function SwapS12(A As S12) As S12
With SwapS12
    .S1 = A.S2
    .S2 = A.S1
End With
End Function
Sub PushS1S2(O() As S12, S1$, S2$, Optional NoTrim As Boolean)
PushS12 O, S12Trim(S1, S2, NoTrim)
End Sub
Function AddS1Pfx(A() As S12, S1Pfx$) As S12()
Dim J&: For J = 0 To S12UB(A)
    Dim M As S12: M = A(J)
    M.S1 = S1Pfx & M.S1
    PushS12 AddS1Pfx, M
Next
End Function
Sub PushS12Opt(O() As S12, M As S12)
If M.S1 <> "" Then PushS12 O, M
End Sub
Sub PushDicS12(O As Dictionary, M As S12, Optional Sep$ = vbCrLf)
With M
    If O.Exists(.S1) Then
        O(.S1) = O(.S1) & " " & O(.S2)
    Else
        O.Add .S1, .S2
    End If
End With
End Sub

'**S12-Dta
Function S12yzDrs(D As Drs, Optional CC$) As S12()
'Fm D  : ..@CC.. ! A drs with col-@CC.  At least has 2 col
'Fm CC :         ! if isBlnk, use fst 2 col
'Ret   :         ! fst col will be S1 and snd col will be S2 join with vbCrLf
Dim S1$(), S2() ' S2 is ay of sy
Dim I1%, I2%
    If CC = "" Then I1 = 0: I2 = 1 Else AsgIx D, CC, I1, I2
Dim Dr: For Each Dr In Itr(D.Dy)
    Dim A$, B$: A = Dr(I1): B = Dr(I2)
    Dim R&: R = IxzAy(S1, A, ThwEr:=EiNoThw)
    If R = -1 Then
        PushI S1, A
        PushI S2, Sy(B)
    Else
        PushI S2(R), B
    End If
Next
Dim J&: For J = 0 To UB(S1)
    PushS12 S12yzDrs, S12(S1(J), JnCrLf(S2(J)))
Next
End Function
Function DrszS12y(A() As S12, Optional FF$ = "S1 S2") As Drs
DrszS12y = DrszFF(FF, DyzS12y(A))
End Function
Function DrzS12(A As S12) As Variant(): DrzS12 = Array(A.S1, A.S2): End Function
Function DyzS12y(A() As S12) As Variant()
Dim J&: For J& = 0 To S12UB(A)
    PushI DyzS12y, DrzS12(A(J))
Next
End Function

Function S12(S1, S2) As S12
With S12
    .S1 = S1
    .S2 = S2
End With
End Function
Function AddS12(A As S12, B As S12) As S12(): PushS12 AddS12, A: PushS12 AddS12, B: End Function
Sub PushS12Ay(O() As S12, A() As S12): Dim J&: For J = 0 To S12UB(A): PushS12 O, A(J): Next: End Sub
Sub PushS12(O() As S12, M As S12): Dim N&: N = S12Si(O): ReDim Preserve O(N): O(N) = M: End Sub
Function S12Si&(A() As S12): On Error Resume Next: S12Si = UBound(A) + 1: End Function
Function S12UB&(A() As S12): S12UB = S12Si(A) - 1: End Function
Function S12Opt(Som, A As S12) As S12Opt: With S12Opt: .Som = Som: .S12 = A: End With: End Function
Function SomS12(A As S12) As S12Opt: SomS12.Som = True: SomS12.S12 = A: End Function

