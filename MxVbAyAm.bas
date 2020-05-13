Attribute VB_Name = "MxVbAyAm"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ay.Op"
Const CMod$ = CLib & "MxVbAyAm."

Function AmAddIxPfx(Ay, Optional BegIx%) As String()
If BegIx < 0 Then AmAddIxPfx = CvSy(Ay): Exit Function
Dim L, J&, N%
J = BegIx
N = Len(CStr(UB(Ay) + J))
For Each L In Itr(Ay)
    PushI AmAddIxPfx, AliR(J, N) & ": " & L
    J = J + 1
Next
End Function

':FunPfx-Am: :FunPfx '#Ay-Map# Function will take an array plus 0 or more parameter and return an array with same number of elements.  The return array may be same or diff type.
Function AmRunAsAv(Ay, Fun$) As Variant()
Dim I: For Each I In Itr(Ay)
    PushI AmRunAsAv, Run(Fun, I)
Next
End Function
Private Sub AmFstCmlByRun__Tst()
Brw AmFstCmlByRun(SrczP(CPj))
End Sub
Function AmFstCmlByRun(Ay) As String()
AmFstCmlByRun = AmRunAsSy(Ay, "FstCml")
End Function
Function AmFstCmlByLoop(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmFstCmlByLoop, FstCml(I)
Next
End Function

Function AmRunAsSy(Ay, Fun$) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRunAsSy, Run(Fun, I)
Next
End Function
Function AmFstCml(Ay) As String()
AmFstCml = AmFstCmlByLoop(Ay)
End Function
Function AmAddPfx(Ay, Pfx) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmAddPfx, Pfx & I
Next
End Function

Function AmAddPfxTab(Ay) As String()
AmAddPfxTab = AmAddPfx(Ay, vbTab)
End Function

Function AmAddPfxSfx(Ay, Pfx, Sfx) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmAddPfxSfx, Pfx & I & Sfx
Next
End Function

Function AmAddSfx(Ay, Sfx) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmAddSfx, I & Sfx
Next
End Function

Function AmRev(Ay)
Dim O: O = Ay
Dim U&: U = UB(Ay)
Dim J&: For J = 0 To U
    O(J - U) = Ay(J)
Next
End Function

Function AmAft(Sy$(), Sep$) As String()
Dim S: For Each S In Itr(Sy)
    Push AmAft, Aft(S, Sep)
Next
End Function

Function AmAftRev(Sy$(), Sep$) As String()
Dim S: For Each S In Itr(Sy)
    Push AmAftRev, AftRev(S, Sep)
Next
End Function
Function AmAli(Ay, Optional W0%) As String()
Dim W%: If W0 <= 0 Then W = WdtzAy(Ay) Else W = W0
Dim S: For Each S In Itr(Ay)
    PushI AmAli, AliL(S, W)
Next
End Function

Function AmAliR(Ay, Optional W0%) As String()
Dim S$: If W0 <= 0 Then S = " "
Dim W%: If W0 <= 0 Then W = WdtzAy(Ay) Else W = W0
Dim I: For Each I In Itr(Ay)
    PushI AmAliR, AliR(I, W) & S
Next
End Function

Function AmBef(Sy$(), Sep$) As String()
Dim S: For Each S In Itr(Sy)
    Push AmBef, Bef(S, Sep)
Next
End Function

Function AmInc(NumAy, Optional By = 1)
Dim O: O = NumAy
Dim J&: For J = 0 To UB(O)
    O(J) = O(J) + By
Next
AmInc = O
End Function

Function AmRmv2Dash(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmv2Dash, Rmv2Dash(I)
Next
End Function

Function AmRmvLasChr(Ay) As String()
'Gen:AyFor RmvLasChr
Dim I
For Each I In Itr(Ay)
    PushI AmRmvLasChr, RmvLasChr(I)
Next
End Function

Function AmRmvPfx(Ay, Pfx$) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmvPfx, RmvPfx(I, Pfx)
Next
End Function

Function AmRpl(Ay, Fm$, By$, Optional Cnt& = 1, Optional C As eCas) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRpl, Replace(I, Fm, By, Count:=Cnt, Compare:=CprMth(C))
Next
End Function

Function AmRTrim(Sy$()) As String()
Dim S: For Each S In Itr(Sy)
    Push AmRTrim, RTrim(S)
Next
End Function

Function AmSyzSS(Ly$()) As Variant()
Dim L: For Each L In Itr(Ly)
    PushI AmSyzSS, SyzSS(L)
Next
End Function

Function AmT1(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmT1, T1(I)
Next
End Function

Function AmT2(Ay) As String()
Dim L: For Each L In Itr(Ay)
    PushI AmT2, T2(L)
Next
End Function

Function AmT3(Ay) As String()
Dim L: For Each L In Itr(Ay)
    PushI AmT3, T3(L)
Next
End Function

Function AmTab(Ay, Optional NTab% = 1) As String()
AmTab = AmAddPfx(Ay, TabN(NTab))
End Function

Function AmTrim(Ay) As String()
Dim S: For Each S In Itr(Ay)
    Push AmTrim, Trim(S)
Next
End Function

Function AmRmvT1(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmvT1, RmvT1(I)
Next
End Function

Function AmRmvTT(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmvTT, RmvTT(I)
Next
End Function

Function AmRmvSngQuo(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmvSngQuo, RmvSngQuo(I)
Next
End Function

Function AmRplStar(Ay, By) As String()
Dim I
For Each I In Itr(Ay)
    PushI AmRplStar, Replace(I, By, "*")
Next
End Function

Function AmRplT1(Ay, NewT1) As String()
AmRplT1 = AmAddPfx(AmRmvT1(Ay), NewT1 & " ")
End Function

Function AmTakBefDD(Sy$()) As String()
Dim I: For Each I In Itr(Sy)
    PushI AmTakBefDD, BefDD(I)
Next
End Function

Function AmTakAftDot(Sy$()) As String()
AmTakAftDot = AmTakAft(Sy, ".")
End Function

Function AmTakAft(Sy$(), Sep$) As String()
Dim I: For Each I In Itr(Sy)
    PushI AmTakAft, Aft(I, Sep)
Next
End Function

Function AmTakAftOrAll(Sy$(), Sep$) As String()
Dim I: For Each I In Itr(Sy)
    PushI AmTakAftOrAll, AftOrAll(I, Sep)
Next
End Function

Function AmTakBef(Sy$(), Sep$) As String() 'Return a Sy which is taking Bef-Sep from Given Sy
Dim I
For Each I In Itr(Sy)
    PushI AmTakBef, Bef(CStr(I), Sep)
Next
End Function

Function AmTakBefDot(Sy$()) As String()
AmTakBefDot = AmTakBef(Sy, ".")
End Function

Function AmTakBefOrAll(Sy$(), Sep$) As String()
Dim I
For Each I In Itr(Sy)
    Push AmTakBefOrAll, BefOrAll(CStr(I), Sep)
Next
End Function

Function AmBetBkt(Sy$()) As String()
Dim I: For Each I In Itr(Sy)
    PushI AmBetBkt, BetBkt(CStr(I))
Next
End Function

Function AmAliQuoSq(Fny$()) As String()
AmAliQuoSq = AmAli(AmQuoSq(Fny))
End Function

Function AmAliByWdty(Ay, Wdty%()) As String()
Dim S, J&: For Each S In Ay
    PushI AmAliByWdty, Ali(S, Wdty(J))
    J = J + 1
Next
End Function


Function AmQuo(Ay, QuoStr$) As String()
If Si(Ay) = 0 Then Exit Function
Dim U&: U = UB(Ay)
Dim Q1$, Q2$
    With BrkQuo(QuoStr)
        Q1 = .S1
        Q2 = .S2
    End With

Dim O$()
    ReDim O(U)
    Dim J&
    For J = 0 To U
        O(J) = Q1 & Ay(J) & Q2
    Next
AmQuo = O
End Function

Function AmQuoDbl(Ay) As String()
AmQuoDbl = AmQuo(Ay, vbDblQ)
End Function

Function AmQuoSng(Sy$()) As String()
AmQuoSng = AmQuo(Sy, vbSngQ)
End Function

Function AmQuoSq(Ay) As String()
AmQuoSq = AmQuo(Ay, "[]")
End Function

Function AmRmvFstChr(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmvFstChr, RmvFstChr(I)
Next
End Function

Function AmRmvFstNonLetter(Ay) As String() 'Gen:AyXXX
Dim I: For Each I In Itr(Ay)
    PushI AmRmvFstNonLetter, RmvFstNonLetter(I)
Next
End Function

Private Sub AmRplQ__Tst()
Dim Ay$(): Ay = SyzSS("Stm Bus L1 L2 L3 L4 Sku")
D AmRplQ(Ay, "#StkDay?")
End Sub

Function AmRplQ(Ay, QmrkStr$) As String()
Dim I: For Each I In Itr(Ay)
    PushS AmRplQ, RplQ(QmrkStr, I)
Next
End Function

Function AmRplMid(Ay, B As Bei, ByAy)
With Ay3FmAyBei(Ay, B)
AmRplMid = AddAyAp(.A, ByAy, .C)
End With
End Function
