Attribute VB_Name = "MxVbAyAw"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Ay.Op"
Const CMod$ = CLib & "MxVbAyAw."
Enum eThwOutRgeEr: eNoThwOutRge: eThwOutRge: End Enum
Function AwBefEle(Ay, Ele)
Const CSub$ = CMod & "AwBefEle"
Dim O: O = Ay: Erase O
Dim I: For Each I In Itr(Ay)
    PushI O, I
    If I = Ele Then AwBefEle = O: Exit For
Next
Thw CSub, "No @Ele in @Ay", "Ele Ay", Ele, Ay
End Function

Function AwBet(Ay, FmEle, ToEle)
Dim O: O = NwAy(Ay)
Dim I: For Each I In Itr(Ay)
    If IsBet(I, FmEle, ToEle) Then
        Push O, I
    End If
Next
AwBet = O
End Function

Function AwDis(Ay, Optional IgnCas As Boolean)
AwDis = IntozItr(NwAy(Ay), CntDi(Ay).Keys)
End Function

Function AwDistAsI(Ay, Optional IgnCas As Boolean) As Integer()
AwDistAsI = CvIntAy(AwDis(Ay, IgnCas))
End Function

Function AwDistAsSy(Ay, Optional IgnCas As Boolean) As String()
AwDistAsSy = CvSy(AwDis(Ay, IgnCas))
End Function

Function AwDistT1(Ay) As String()
AwDistT1 = AwDis(AmT1(Ay))
End Function

Function AwDup(Ay, Optional C As VbCompareMethod = vbTextCompare)
Dim O: O = NwAy(Ay)
Dim D As Dictionary: Set D = CntDi(Ay, C)
Dim K: For Each K In D.Keys
    If D(K) > 1 Then PushI O, K
Next
AwDup = O
End Function

Function SywDup(Sy$(), Optional C As VbCompareMethod = vbTextCompare) As String()
SywDup = AwDup(Sy, C)
End Function

Function AwEQ(Ay, V)
Dim O: O = Ay: Erase O
Dim I: For Each I In Itr(Ay)
    If I = V Then PushI O, I
Next
AwEQ = O
End Function
Function AwBE(Ay, Bix, Eix, Optional Thw As eThwOutRgeEr)
Dim U&: U = UB(Ay)
If Thw = eThwOutRge Then
    ChkIsBet Bix, 0, U, CSub
    ChkIsBet Eix, 0, U, CSub
Else
    AwBE = Ay: Erase AwBE
    If Not IsBet(Bix, 0, U) Then Exit Function
    If Not IsBet(Eix, 0, U) Then Exit Function
End If
Dim J&: For J = Bix To Eix
    Push AwBE, Ay(J)
Next
End Function
Function AwBeiAsSy(Ay, B As Bei) As String(): AwBeiAsSy = AwBei(Ay, B):              End Function
Function AwBei(Ay, B As Bei):                     AwBei = AwBEThw(Ay, B.Bix, B.Eix): End Function
Function AwBEThw(Ay, Bix, Eix):                 AwBEThw = AwBE(Ay, Bix, Eix):        End Function

Function AwBeiy(Ay, B() As Bei)
Dim J%: For J = 0 To BeiUB(B)
    PushIAy AwBeiy, AwBei(Ay, B(J))
Next
End Function

Function AwBix(Ay, Bix)
AwBix = Ay: Erase AwBix
Dim J&: For J = Bix To UB(Ay)
    Push AwBix, Ay(J)
Next
End Function

Function AwNB(Ay) As String()
Dim I: For Each I In Itr(Ay)
    If Trim(I) <> "" Then PushI AwNB, I
Next
End Function

Function AwEix(Ay, Eix)
Dim U&: U = UB(Ay)
If Eix > U Then AwEix = Ay: Exit Function
Dim O: O = Ay
ReDim Preserve O(U)
AwEix = O
End Function


Function AwGT(Ay, V)
If Si(Ay) <= 1 Then AwGT = Ay: Exit Function
AwGT = NwAy(Ay)
Dim I: For Each I In Ay
    If I > V Then PushI AwGT, I
Next
End Function

Function AwInAet(Ay, Aet As Dictionary)
AwInAet = NwAy(Ay)
Dim I
For Each I In Itr(Ay)
    If Aet.Exists(I) Then Push AwInAet, I
Next
End Function

Function AwAftEle(Ay, Ele)
Const CSub$ = CMod & "AwAftEle"
Dim O: O = Ay: Erase O
Dim I, F As Boolean: For Each I In Itr(Ay)
    If F Then
        PushI O, I
    Else
        If I = Ele Then F = True
    End If
   
Next
Thw CSub, "No @Ele in @Ay", "Ele Ay", Ele, Ay
End Function

Function AwIsNm(Ay) As String()
Dim I: For Each I In Itr(Ay)
    If IsNm(I) Then PushI AwIsNm, I
Next
End Function

Function AwIxCnt(Ay, Ix, Cnt)
Dim J&
Dim O: O = Ay: Erase O
For J = 0 To Cnt - 1
    Push O, Ay(Ix + J)
Next
AwIxCnt = O
End Function

Function AwIxy(Ay, Ixy)  ' Array where elements at pointed by @Ixy allow empty if the Ix is outside range of @Ay
AwIxy = AwIxyzAlwEmp(Ay, Ixy)
End Function

Function AwIxyzAlwEmp(Ay, Ixy) ' Array where elements at pointed by @Ixy allow empty if the Ix is outside range of @Ay
If Si(Ixy) = 0 Then
    AwIxyzAlwEmp = NwAy(Ay)
    Exit Function
End If
Dim U&: U = UB(Ixy)
Dim O: O = NwAy(Ay)
Dim AU&: AU = UB(Ay)
ReDim Preserve O(U)
Dim Ix, J&
For Each Ix In Itr(Ixy)
    If IsBet(Ix, 0, AU) Then
        Asg Ay(Ix), O(J)
    End If
    J = J + 1
Next
AwIxyzAlwEmp = O
End Function

Function AwIxyzMust(Ay, Ixy) ' Array where elements at pointed by @Ixy which have Ix-ele must be in range of @Ay
Dim U&: U = UB(Ixy)
Const CSub$ = CMod & "AwIxyzMust"
If IsIxyOut(Ixy, UB(Ay)) Then Thw CSub, "Some element in Ixy is outsize Ay", "UB(Ay) Ixy", UB(Ay), Ixy
Dim O: O = NwAy(Ay)
Dim Ix
For Each Ix In Itr(Ixy)
    Push O, Ay(Ix)
Next
AwIxyzMust = O
End Function

Function AwKssAy(Ay, KssAy$()) As String()
Dim LikAy$(): LikAy = LikAyzKssAy(KssAy)
Dim S: For Each S In Itr(Ay)
    If HitLikAy(S, LikAy) Then PushI AwKssAy, S
Next
End Function

Function AwLasN(Ay, N)
Dim O, J&, I&, U&, Fm&, NewU&
U = UB(Ay)
If U < N Then AwLasN = Ay: Exit Function
O = Ay: Erase O
Fm = U - N + 1
For J = Fm To U
    Push O, Ay(J)
Next
AwLasN = O
End Function

Function AwLE(Ay, V)
If Si(Ay) <= 1 Then AwLE = Ay: Exit Function
AwLE = NwAy(Ay)
Dim I: For Each I In Ay
    If I <= V Then PushI AwLE, I
Next
End Function

Function AwLik(Ay, Lik) As String()
Dim I: For Each I In Itr(Ay)
    If I Like Lik Then PushI AwLik, I
Next
End Function

Function AwLikAy(Ay, LikAy$()) As String()
Dim I, Lik
For Each I In Itr(Ay)
    If HitLikAy(I, LikAy) Then PushI AwLikAy, I
Next
End Function

Function AwLikss(Ay, Likss$) As String()
AwLikss = AwLikAy(Ay, SyzSS(Likss))
End Function

Function AwLT(Ay, V)
If Si(Ay) = 1 Then AwLT = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I < V Then PushI O, I
Next
AwLT = O
End Function

Function AwMid(Ay, Fm, Optional L = 0)
AwMid = NwAy(Ay)
Dim J&
Dim E&
    Select Case True
    Case L = 0: E = UB(Ay)
    Case Else:  E = Min(UB(Ay), L + Fm - 1)
    End Select
For J = Fm To E
    Push AwMid, Ay(J)
Next
End Function

Function AwmRmvT1(Ly$(), T1) As String()
Dim L: For Each L In Itr(Ly)
    If ShfTerm(L) = T1 Then PushI AwmRmvT1, L
Next
End Function

Function AwWhNm(Ay, B As WhNm) As String()
Dim I: For Each I In Itr(Ay)
    If HitNm(I, B) Then PushI AwWhNm, I
Next
End Function

Function AwNmStr(Ay, WhNmStr$) As String()
AwNmStr = AwWhNm(Ay, WhNm(WhNmStr))
End Function

Function AwNm(Ay) As String()
Dim Nm: For Each Nm In Itr(Ay)
    If IsNm(Nm) Then PushI AwNm, Nm
Next
End Function

Function AwRx(Ay, Rx As RegExp) As String()
Dim I: For Each I In Itr(Ay)
    If Rx.Test(I) Then PushI AwRx, I
Next
End Function
Function AwPatn(Ay, Patn$) As String()
If Patn = "" Then AwPatn = Ay: Exit Function
AwPatn = AwRx(Ay, Rx(Patn))
End Function

Function AwPfx(Ay, Pfx) As String()
Dim S: For Each S In Itr(Ay)
    If HasPfx(S, Pfx) Then PushI AwPfx, S
Next
End Function

Function AwRmvEle(Ay, Ele)
AwRmvEle = NwAy(Ay)
Dim I
For Each I In Itr(Ay)
    If I <> Ele Then PushI AwRmvEle, I
Next
End Function

Function AwRmvT1(Ay, T1) As String()
AwRmvT1 = AmRmvT1(AwT1(Ay, T1))
End Function

Function AwRmvTT(Ay, T1, T2) As String()
AwRmvTT = AmRmvTT(AwTT(Ay, T1, T2))
End Function

Function AwSfx(Ay, Sfx$) As String()
Dim I
For Each I In Itr(Ay)
    If HasSfx(I, Sfx) Then PushI AwSfx, I
Next
End Function

Function AwSingleEle(Ay)
Dim O: O = Ay: Erase O
Dim CntDy(): CntDy = CntgDy(Ay, EiCntSng)
If Si(CntDy) = 0 Then
    AwSingleEle = O
    Exit Function
End If
Dim Dr
For Each Dr In CntDy
    If Dr(1) = 1 Then
        Push O, Dr(0)
    End If
Next
AwSingleEle = O
End Function

Function AwSkip(Ay, Optional SkipN& = 1)
Const CSub$ = CMod & "AwSkip"
If SkipN <= 0 Then AwSkip = Ay: Exit Function
Dim U&: U = UB(Ay) - SkipN: If SkipN < -1 Then Thw CSub, "Ay is not enough to skip", "Si-Ay SkipN", "Si(Ay),SKipN"
Dim O: O = Ay: Erase O
Dim J&: For J = SkipN To U
    Push O, Ay(J)
Next
AwSkip = O
End Function

Function AwSng(Ay)
AwSng = MinusAy(Ay, AwDup(Ay))
End Function

Function AwSngEle(Ay)
'Return Set of Element as array in {Ay} having 2 or more element
Dim O: O = NwAy(Ay)
Dim K, D As Dictionary
Set D = CntDi(Ay)
For Each K In D.Keys
    If D(K) = 1 Then PushI O, K
Next
End Function

Function AwSubStr(Ay, SubStr, Optional C As eCas) As String()
Dim I: For Each I In Itr(Ay)
    If HasSubStr(I, SubStr, C) Then
        PushI AwSubStr, I
    End If
Next
End Function

Function AwT1(Ay, T1) As String()
Dim L$, I
For Each I In Itr(Ay)
    L = I
    If HasT1(L, T1) Then
        PushI AwT1, L
    End If
Next
End Function

Function AwT1InAy(Ay, InAy) As String()
If Si(Ay) = 0 Then Exit Function
Dim O$(), L: For Each L In Ay
    If HasEle(InAy, T1(L)) Then Push O, L
Next
AwT1InAy = O
End Function

Function AwTT(Ay, T1, T2) As String()
Dim I, L$
For Each I In Itr(Ay)
    L = I
    If HasTT(L, T1, T2) Then PushI AwTT, L
Next
End Function

Function AwTTSelRst(Ay, T1, T2) As String()
Dim L$, I, X1$, X2$, Rst$
For Each I In Itr(Ay)
    L = I
    AsgTTRst L, X1, X2, Rst
    If X1 = T1 Then
        If X2 = T2 Then
            PushI AwTTSelRst, Rst
        End If
    End If
Next
End Function

Function AwPfxss(Ay, Pfxss$) As String()
Dim Pfxy$(): Pfxy = SyzSS(Pfxss)
Dim I: For Each I In Itr(Ay)
    If HasPfxy(I, Pfxy) Then PushI AwPfxss, I
Next
End Function

Function AwEndTrim(Ly$()) As String()
Dim O$(): O = Ly
Dim NBlnkUB&
    For NBlnkUB = UB(Ly) To 0 Step -1
        If LTrim(Ly(NBlnkUB)) <> "" Then Exit For
    Next
    If NBlnkUB = -1 Then Exit Function
ReDim Preserve O(NBlnkUB)
AwEndTrim = O
End Function
Function AwPred(Ay, Pred$)
Dim O: O = NwAy(Ay)
Dim I: For Each I In Itr(Ay)
    If Run(Pred, I) Then PushI O, I
Next
AwPred = O
End Function
