Attribute VB_Name = "MxVbDtaDicNw"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Dic"
Const CMod$ = CLib & "MxVbDtaDicNw."
Function DiczFt(Ft) As Dictionary
Set DiczFt = Dic(LyzFt(Ft))
End Function

Function DiT1qLy(TermLny$()) As Dictionary
Dim L$, I, T$, Ssl$
Dim O As New Dictionary
For Each I In Itr(TermLny)
    L = I
    AsgTRst L, T, Ssl
    If O.Exists(T) Then
        O(T) = AddAy(O(T), SyzSS(Ssl))
    Else
        O.Add T, SyzSS(Ssl)
    End If
Next
Set DiT1qLy = O
End Function

Function DiczLines(Lines$, Optional JnSep$ = vbCrLf) As Dictionary
Set DiczLines = Dic(SplitCrLf(Lines), JnSep)
End Function

Sub AppLinzToLinesDic(OLinesDic As Dictionary, K, Ln, Sep$)
If OLinesDic.Exists(K) Then
    OLinesDic(K) = OLinesDic(K) & Sep & Ln
Else
    OLinesDic.Add K, Ln
End If
End Sub

Function LyzLinesDicItems(LineszDic As Dictionary) As String()
Dim Lines$, I
For Each I In LineszDic.Items
    Lines = I
    PushIAy LyzLinesDicItems, SplitCrLf(Lines)
Next
End Function

Function DiczVkkLy(VkkLy$()) As Dictionary
Set DiczVkkLy = New Dictionary
Dim I, V$, Vkk$, K
For Each I In Itr(VkkLy)
    Vkk = I
    V = T1(Vkk)
    For Each K In SyzSS(RmvT1(Vkk))
        DiczVkkLy.Add K, V
    Next
Next
End Function

Function LyzDic(A As Dictionary) As String()
Dim K
For Each K In A.Keys
    PushI LyzDic, K & " " & A(K)
Next
End Function
Function DiczDrsCC(A As Drs, Optional CC$) As Dictionary
If CC = "" Then
    Set DiczDrsCC = DiczDyCC(A.Dy)
Else
    With BrkSpc(CC)
        Dim C1%: C1 = IxzAy(A.Fny, .S1)
        Dim C2%: C2 = IxzAy(A.Fny, .S2)
        Set DiczDrsCC = DiczDyCC(A.Dy, C1, C2)
    End With
End If
End Function

Function DiczDyCC(Dy(), Optional C1 = 0, Optional C2 = 1) As Dictionary
Set DiczDyCC = New Dictionary
Dim Dr
For Each Dr In Itr(Dy)
    DiczDyCC.Add Dr(C1), Dr(C2)
Next
End Function
Function DiczUniq(Ly$()) As Dictionary 'T1 of each Ly must be uniq
Set DiczUniq = New Dictionary
Dim I
For Each I In Itr(Ly)
    DiczUniq.Add T1(I), RmvT1(I)
Next
End Function

Function DiKqABC(AyOfSi26OrLess) As Dictionary
Const CSub$ = CMod & "DiKqABC"
'Ret : :DiKqABC: is a dic wi v running fm A-Z at most 26 ele.  The k is CStr fm @AyOfSi26OrLess-ele.
If Si(AyOfSi26OrLess) > 26 Then Thw CSub, "Si-@AyOfSi26OrLess cannot >26", "Si-@AyOfSi26OrLess", Si(AyOfSi26OrLess)
Dim O As New Dictionary
Dim V, J&: For Each V In Itr(AyOfSi26OrLess)
    V = CStr(V)
    If Not O.Exists(V) Then
        O.Add V, Chr(65 + J)
    End If
    J = J + 1
Next
Set DiKqABC = O
End Function

Function DiT1qLyItr(TRstLy$(), T1ss$) As Dictionary
'Fm TRstLy : T Rst             ! it is ly of [T1 Rst]
'Fm T1ss   : SS                ! it is a list T1 in SS fmt.
'Ret       : DicOf T1 to LyItr ! it will have sam of keys as (@T1ss nitm + 1).
'                              ! Each val is either :Ly or emp Vb.Collection if no such T1.  The :Ly will have T1 rmv.
'                              ! The las key is '*Er' and the val is :Ly or emp-vb.Collection.  The :Ly will have T1 incl.
Dim AmT1$(): AmT1 = SyzSS(T1ss)

Dim O As New Dictionary
Dim Er$()               ' The er Ln of @TRstLy
    Dim T1: For Each T1 In Itr(AmT1)  ' Put all T1 in @AmT1 to @O
        O.Add T1, EmpSy
    Next
    Dim T$, Rst$, L, Ly$(): For Each L In Itr(TRstLy) ' For each @TRstLy Ln put it to either @O or @Er
        AsgTRst L, T, Rst
        If O.Exists(T) Then
            Ly = O(T)
            PushI Ly, Rst
            O(T) = Ly       '<-- Put to @O
        Else
            PushI Er, L      '<-- Put to @Er
        End If
    Next
SetDicValAsItr O                '<-- for each dic val setting to itr
O.Add "*Er", Itr(Er)
Set DiT1qLyItr = O
End Function

Function DiczFnyDr(Fny$(), Dr) As Dictionary
Set DiczFnyDr = New Dictionary
Dim F, J%: For Each F In Fny
    DiczFnyDr.Add F, Dr(J)
    J = J + 1
Next
End Function

Function DiczKv(K, V) As Dictionary
Set DiczKv = New Dictionary
DiczKv.Add K, V
End Function

Function EmpDic() As Dictionary
Set EmpDic = New Dictionary
End Function

Function Dic(Ly$(), Optional JnSep$ = vbCrLf) As Dictionary
Set Dic = DiczS12y(S12y(Ly), JnSep)
End Function

Function DiczKyVy(Ky, Vy) As Dictionary
ChkSamSi Ky, Vy, CSub
Dim J&
Set DiczKyVy = New Dictionary
For J = 0 To UB(Ky)
    DiczKyVy.Add Ky(J), Vy(J)
Next
End Function

Function DiczSyab(A$(), B$(), Optional JnSep$ = vbCrLf) As Dictionary
ChkSamSi A, B, , CSub
Dim O As New Dictionary
Dim I, J&: For Each I In Itr(A)
    If O.Exists(I) Then
       O(I) = O(I) & JnSep & B(J)
    Else
        O.Add I, B(J)
    End If
    J = J + 1
Next
Set DiczSyab = O
End Function

Function DiczAy2(A, B) As Dictionary
ChkSamSi A, B, CSub
Dim N1&, N2&
N1 = Si(A)
N2 = Si(B)
If N1 <> N2 Then Stop
Set DiczAy2 = New Dictionary
Dim J&, X
For Each X In Itr(A)
    DiczAy2.Add X, B(J)
    J = J + 1
Next
End Function

Function DiKqIx(Ay) As Dictionary
Dim O As New Dictionary, J&
For J = 0 To UB(Ay)
    If Not O.Exists(Ay(J)) Then
        O.Add Ay(J), J
    End If
Next
Set DiKqIx = O
End Function

Function DiKqNum(Ay) As Dictionary
Dim O As New Dictionary, J&
For J = 0 To UB(Ay)
    If Not O.Exists(Ay(J)) Then
        O.Add Ay(J), J + 1
    End If
Next
Set DiKqNum = O
End Function


Function ValTyAy(A As Dictionary) As String()
Dim V: For Each V In A.Items
    PushI ValTyAy, TypeName(V)
Next
End Function
Function DiwAy(DiAqB As Dictionary, Ay) As Dictionary
'Ret : :DiAqB #SubSet-Of-Dic-By-Ay#
Set DiwAy = New Dictionary
Dim A: For Each A In Itr(Ay)
    DiwAy.Add A, DiAqB(A)
Next
End Function
