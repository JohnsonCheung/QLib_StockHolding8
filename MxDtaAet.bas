Attribute VB_Name = "MxDtaAet"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Dta.Aet"
Const CMod$ = CLib & "MxDtaAet."
Function EmpAet() As Dictionary
Set EmpAet = New Dictionary
End Function

Function CvAet(V) As Dictionary
Set CvAet = V
End Function

Function IsAet(V) As Boolean
Select Case True
Case TypeName(V) <> "Dictionary"
Case Not IsAllEmpItr(CvDic(V).Items)
Case Else: IsAet = True
End Select
End Function

Function AetzAp(ParamArray Ap()) As Dictionary
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
Set AetzAp = Aet(Av)
End Function

Function EmpColl() As VBA.Collection
Set EmpColl = New VBA.Collection
End Function
Function AddAetItr(Aet As Dictionary, Itr) As Dictionary
Dim I: For Each I In Itr
Next
End Function
Function AetzItr(Itr) As Dictionary
Set AetzItr = AddAetItr(EmpAet, Itr)
End Function

Function AetzT(Termln$) As Dictionary
Set AetzT = Aet(Termy(Termln))
End Function

Function AetzSS(SS$) As Dictionary
Set AetzSS = Aet(SyzSS(SS))
End Function

Function SrtAet(Aet As Dictionary) As Dictionary
Set SrtAet = Aet(QSrt(AvzAet(Aet)))
End Function

Function Aet(Ay) As Dictionary
Set Aet = EmpAet
PushAyzAet Aet, Ay
End Function

Function AetzItm(Itm) As Dictionary
Set AetzItm = EmpAet
PushEle AetzItm, Itm
End Function

Sub BrwAet(Aet As Dictionary, Optional FnPfx$ = "BrwAet_")
Brw Aet.Keys, FnPfx
End Sub

Function CloneAet(Aet As Dictionary) As Dictionary
Set CloneAet = New Dictionary
Dim K: For Each K In Aet.Keys
    CloneAet.Add K, Empty
Next
End Function

Function AddAet(Aet1 As Dictionary, Aet2 As Dictionary) As Dictionary
Set AddAet = CloneAet(Aet1)
Dim K: For Each K In Aet2.Keys
    PushEle AddAet, K
Next
End Function

Function AvzAet(Aet As Dictionary) As Variant()
AvzAet = AvzItr(Aet.Keys)
End Function

Sub DmpAet(Aet As Dictionary)
D Aet.Keys
End Sub

Function FstItmzAet(Aet As Dictionary)
Dim I: For Each I In Aet.Keys
    Asg I, FstItmzAet: Exit Function
Next
End Function

Function IsEmpAet(Aet As Dictionary) As Boolean
IsEmpAet = Aet.Count = 0
End Function

Function IsEqAet(Aet1 As Dictionary, Aet2 As Dictionary) As Boolean
If Aet1.Cnt <> Aet2.Cnt Then Exit Function
Dim K1: For Each K1 In Aet1.Keys
    If Not Aet2.Exists(K1) Then Exit Function
Next
IsEqAet = True
End Function

Function IsEqAetzInOrd(Aet1 As Dictionary, Aet2 As Dictionary) As Boolean
If Aet1.Count <> Aet2.Count Then Exit Function
IsEqAetzInOrd = IsEqAy(AvzAet(Aet1), AvzAet(Aet2))
End Function

Function LinzAet$(Aet As Dictionary)
LinzAet = JnSpc(AvzAet(Aet))
End Function

Function MinusAet(Aet1 As Dictionary, Aet2 As Dictionary) As Dictionary
Set MinusAet = New Dictionary
Dim E1: For Each E1 In Aet1.Keys
    If Not Aet2.Exists(E1) Then PushEle MinusAet, E1
Next
End Function

Sub PushSet(OAet As Dictionary, Aet As Dictionary)
PushItrzAet OAet, Aet.Keys
End Sub

Sub PushAyzAet(OAet As Dictionary, Ay)
Dim I: For Each I In Itr(Ay)
    PushEle OAet, I
Next
End Sub

Sub PushEle(OAet As Dictionary, Ele)
If Not OAet.Exists(Ele) Then OAet.Add Ele, Empty
End Sub

Sub PushItrzAet(Aet As Dictionary, Itr, Optional NoBlnkStr As Boolean)
Dim I
If NoBlnkStr Then
    For Each I In Itr
        If I <> "" Then
            PushEle Aet, I
        End If
    Next
Else
    For Each I In Itr
        PushEle Aet, I
    Next
End If
End Sub

Function RmvEle(Aet As Dictionary, Ele) As Dictionary
Set RmvEle = CloneAet(Aet)
If RmvEle.Exists(Ele) Then RmvEle.Remove Ele
End Function

Function SyzAet(Aet As Dictionary) As String()
SyzAet = SyzAy(Aet.Keys)
End Function

Function TermLnzAet$(Aet As Dictionary)
TermLnzAet = Termln(AvzAet(Aet))
End Function

Function FmtAet(A As Dictionary) As String()
Dim N%: N = NDig(A.Count)
Dim O$()
Dim K: For Each K In A.Keys
    Dim J%: J = J + 1
    PushI O, AliR(J, N) & " " & K
Next
End Function

Sub VcAet(Aet As Dictionary, Optional FnPfx$ = "VcAet_")
VcAy FmtAet(Aet), FnPfx
End Sub
