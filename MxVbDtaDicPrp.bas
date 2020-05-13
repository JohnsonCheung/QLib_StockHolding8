Attribute VB_Name = "MxVbDtaDicPrp"
Option Explicit
Option Compare Text
Const CNs$ = "Vb.Dic"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbDtaDicPrp."

Function SKy(D As Dictionary) As String()
':SKy: :Sy ! #Str-Key-Array# it comes from the all keys of a Di
SKy = SyzItr(D.Keys)
End Function

Function Vy(D As Dictionary) As Variant()
Vy = AvzItr(D.Items)
End Function

Function SVy(D As Dictionary) As String()
':SVy: :Sy ! #Str-Value-Array-Of-a-Dictionary# it comes from the all Items of a Dic
SVy = SyzItr(D.Items)
End Function

Function VzDik(A As Dictionary, K)
If IsNothing(A) Then Exit Function
If A.Exists(K) Then Asg A(K), VzDik
End Function

Function VyzDikk(Dic As Dictionary, KK$) As Variant()
VyzDikk = VyzDiky(Dic, SyzSS(KK))
End Function

Function LineszLinesDic(LinesDic As Dictionary, Optional LinesSep$ = vbCrLf) ' Return the joined Lines from LinesDic
Dim O$(), I, Lines$
For Each I In LinesDic.Items
    PushI O, I
Next
LineszLinesDic = Jn(O, LinesSep)
End Function

Function AddPfxToKey(Pfx$, A As Dictionary) As Dictionary
Dim K
Set AddPfxToKey = New Dictionary
For Each K In A.Keys
    AddPfxToKey.Add Pfx & K, A(K)
Next
End Function

Function StrVy(A As Dictionary) As String()
StrVy = StrVyzK(A, QSrt(SKy(A)))
End Function

Function StrVyzK(A As Dictionary, Ky) As String()
Dim K: For Each K In Itr(Ky)
    PushI StrVyzK, JnCrLf(FmtV(A(K)))
Next
End Function

Function DicHasBlnkKey(A As Dictionary) As Boolean
If A.Count = 0 Then Exit Function
Dim K
For Each K In A.Keys
   If Trim(K) = "" Then DicHasBlnkKey = True: Exit Function
Next
End Function

Function DicHasK(A As Dictionary, K$) As Boolean
DicHasK = A.Exists(K)
End Function

Function DicHasKeyLvs(A As Dictionary, KeyLvs) As Boolean
DicHasKeyLvs = DicHasKy(A, SyzSS(KeyLvs))
End Function

Sub DicHasKeyssAss(A As Dictionary, KeySS$)
DicHasKyAss A, SyzSS(KeySS)
End Sub

Function DicHasKeySsl(A As Dictionary, KeySsl) As Boolean
DicHasKeySsl = A.Exists(SyzSS(KeySsl))
End Function

Function DicHasKy(A As Dictionary, Ky) As Boolean
Ass IsArray(Ky)
If Si(Ky) = 0 Then Stop
Dim K
For Each K In Ky
   If Not A.Exists(K) Then
       Debug.Print FmtQQ("Dix.HasKy: Key(?) is Missing", K)
       Exit Function
   End If
Next
DicHasKy = True
End Function

Sub DicHasKyAss(A As Dictionary, Ky)
Dim K
For Each K In Ky
   If Not A.Exists(K) Then Debug.Print K: Stop
Next
End Sub


Private Sub IsDikStr__Tst()
Dim A As Dictionary
GoSub T1
Exit Sub
T1:
    Set A = New Dictionary
    Dim J&
    For J = 1 To 10000
        A.Add J, J
    Next
    Ept = True
    GoSub Tst
    '
    A.Add 10001, "X"
    Ept = False
    GoTo Tst
Tst:
    Act = IsDikStr(A)
    C
    Return
End Sub

Function Tyny(Ay) As String(): Tyny = TynyzItr(Itr(Ay)): End Function
Function TynyzDic(A As Dictionary) As String(): TynyzDic = TynyzItr(A.Items): End Function
Function TynyzItr(Itr) As String(): Dim V: For Each V In Itr: PushI TynyzItr, TypeName(V): Next: End Function

Function VyzDiky(D As Dictionary, Ky) As Variant()
Const CSub$ = CMod & "VyzDicKy"
Dim K
For Each K In Itr(Ky)
    If Not D.Exists(K) Then Thw CSub, "Some K in given Ky not found in given Dic keys", "[K with error] [given Ky] [given dic keys]", K, AvzItr(D.Keys), Ky
    Push VyzDiky, D(K)
Next
End Function
Function DicwKy(D As Dictionary, Ky) As Dictionary
Set DicwKy = New Dictionary
Dim Vy(): Vy = VyzDiky(D, Ky)
Dim K, J&
For Each K In Itr(Ky)
    DicwKy.Add K, Vy(J)
    J = J + 1
Next
End Function

Function VyzDii(A As Dictionary) As Variant()
VyzDii = AvzItr(A.Items)
End Function

Function DicTy$(A As Dictionary)
Dim O$
Select Case True
Case IsDicEmp(A):   O = "EmpDic"
Case IsDiiStr(A):   O = "StrDic"
Case IsDiiLines(A): O = "LineszDic"
Case IsDiiSy(A):    O = "DiT1qLy"
Case Else:           O = "Dic"
End Select
End Function

Sub AddDicLin(ODic As Dictionary, DicLin$)
With BrkSpc(DicLin)
    ODic.Add .S1, .S2
End With
End Sub
Function AddDic(A As Dictionary, B As Dictionary) As Dictionary
Set AddDic = New Dictionary
PushDic AddDic, A
PushDic AddDic, B
End Function

Function DicAyzAp(ParamArray DicAp()) As Dictionary()
Const CSub$ = CMod & "DicAyzAp"
Dim Av(): Av = DicAp: If Si(Av) = 0 Then Exit Function
Dim I
For Each I In Av
    If Not IsDic(I) Then Thw CSub, "Some itm is not Dic", "TypeName-Ay", VbTyNy(Av)
    PushObj DicAyzAp, CvDic(I)
Next
End Function

Function DefDic(Ly$(), KK) As Dictionary
Const CSub$ = CMod & "DefDic"
Dim L, Aet As Dictionary, T1$, Rst$, O As New Dictionary
Set Aet = TermAet(KK)
If Aet.Exists("*Er") Then Thw CSub, "KK cannot have Term-*Er", "KK Ly", KK, Ly
For Each L In Ly
    AsgTRst L, T1, Rst
    If Aet.Exists(T1) Then
        PushItmzDiT1qLy O, T1, Rst
    Else
'        PushItmzDiT1qLy , O, L
    End If
    Set DefDic = O
Next
End Function

Function TynyzDiiKy(A As Dictionary, Ky) As String()
Dim K: For Each K In Itr(Ky)
    PushI TynyzDiiKy, TypeName(A(K))
Next
End Function

Function TynyzDii(A As Dictionary) As String()
TynyzDii = TynyzDiiKy(A, SKy(A))
End Function
