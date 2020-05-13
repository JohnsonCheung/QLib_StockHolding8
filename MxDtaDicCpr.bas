Attribute VB_Name = "MxDtaDicCpr"
Option Compare Text
Option Explicit
Const CNs$ = "Cpr.Dic"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDicCpr."
Private Type DiCpr
    Kn As String
    H12 As String
    IsExlSam As Boolean
    AExcess As Dictionary
    BExcess As Dictionary
    ADif As Dictionary
    BDif As Dictionary
    Sam As Dictionary
End Type
Public Const SamKeyDifValFF$ = "Key ValA ValB"
Function FmtCprDic(A As Dictionary, B As Dictionary, Optional KeyNm12ss$ = "Key Fst Snd", Optional IsExlSam As Boolean) As String()
FmtCprDic = ZZFmtDiCpr(ZZDiCpr(A, B, KeyNm12ss, IsExlSam))
End Function

Private Sub BrwCprDic__Tst()
Dim A As Dictionary, B As Dictionary
Set A = DiczVbl("X AA|A BBB|A Lines1|A Line3|B Line1|B line2|B line3..")
BrwDic A
Stop
Set B = DiczVbl("X AA|C Line|D Line1|D line2|B Line1|B line2|B line3|B Line4")
BrwCprDic A, B
End Sub

Sub BrwCprDic(A As Dictionary, B As Dictionary, Optional NmOfKeyNm12$ = "Key Fst Snd", Optional IsExlSam As Boolean)
BrwAy ZZFmtDiCpr(ZZDiCpr(A, B, NmOfKeyNm12, IsExlSam))
End Sub

Private Function SamKvDi(A As Dictionary, B As Dictionary) As Dictionary
Set SamKvDi = New Dictionary
If A.Count = 0 Or B.Count = 0 Then Exit Function
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            SamKvDi.Add K, A(K)
        End If
    End If
Next
End Function

Private Function W1FmtDif(A As Dictionary, B As Dictionary, Kn$, H12$) As String()
'@H12:: :NN #Name-1-and-2# Use !AsgTRst to get 2 names
Const CSub$ = CMod & "W1FmtDif"
If A.Count <> B.Count Then Thw CSub, "Dic A & B should have same size", "Dic-A-Si Dic-B-Si", A.Count, B.Count
If A.Count = 0 Then Exit Function
Dim O$(), K, S1$, S2$, S() As S12, KK$
For Each K In A
    KK = K
    S1 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & ULinzLines(KK) & vbCrLf & A(K)
    S2 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & ULinzLines(KK) & vbCrLf & B(K)
    PushS12 S, S12(S1, S2)
Next
W1FmtDif = FmtS12y(S, H12:=H12)
End Function

Function SamKeyDifValDrs(A As Dictionary, B As Dictionary, Kn$, Nm1$, Nm2$) As Drs
ChkDiiIsStr A, CSub
ChkDiiIsStr B, CSub
ChkDicabSamKey A, B, CSub
Dim Dy()
    Dim K: For Each K In A.Keys
        PushI Dy, Array(K, A(K), B(K))
    Next
SamKeyDifValDrs = DrszFF(SamKeyDifValFF, Dy)
End Function

Private Function ZZDiCpr(A As Dictionary, B As Dictionary, KeyH12$, IsExlSam As Boolean) As DiCpr
Dim O As DiCpr
Dim L$: L = KeyH12
O.Kn = ShfTerm(L)
O.H12 = L
Set O.AExcess = MinusDic(A, B)
Set O.BExcess = MinusDic(B, A)
Set O.Sam = SamKvDi(A, B)
With W1DifDi2(A, B)
    Set O.ADif = .A
    Set O.BDif = .B
End With
ZZDiCpr = O
End Function

Private Function W1DifDi2(A As Dictionary, B As Dictionary) As Di2
Dim OA As New Dictionary, OB As New Dictionary
Dim K: For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) <> B(K) Then
            OA.Add K, A(K)
            OB.Add K, B(K)
        End If
    End If
Next
W1DifDi2 = Di2(OA, OB)
End Function

'--
Private Function ZZFmtDiCpr(A As DiCpr) As String()
Dim O$()
With A
    Dim Nm1$, Nm2$
    AsgTRst A.H12, Nm1, Nm2
    O = AddAyAp( _
        W1FmtExcess(.AExcess, .Kn, Nm1), _
        W1FmtExcess(.BExcess, .Kn, Nm2), _
        W1FmtDif(.ADif, .BDif, .Kn, .H12))
    If Not .IsExlSam Then
        O = AddAy(O, W1FmtSam(A.Sam, .Kn, Nm1, Nm2))
    End If
End With
ZZFmtDiCpr = O
End Function

Private Function W1FmtExcess(A As Dictionary, Kn$, Nm$) As String()
If A.Count = 0 Then Exit Function
Dim K, S1$, S2$, S() As S12
For Each K In A.Keys
    S1 = ULinzLines(CStr(K))
    S2 = A(K)
    PushS12 S, S12(S1, S2)
Next
PushIAy W1FmtExcess, Box(FmtQQ("!Er (?) has Excess", Nm))
PushAy W1FmtExcess, FmtS12y(S, H12:="Key " & Nm)
End Function

Private Function W1FmtSam(A As Dictionary, Kn$, Nm1$, Nm2$) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, S() As S12, KK$
For Each K In A.Keys
    KK = K
    PushS12 S, S12("*Same", K & vbCrLf & ULinzLines(KK) & vbCrLf & A(K))
Next
W1FmtSam = FmtS12y(S)
End Function
