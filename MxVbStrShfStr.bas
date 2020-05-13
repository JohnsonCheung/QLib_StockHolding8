Attribute VB_Name = "MxVbStrShfStr"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Str.Shf"
Const CMod$ = CLib & "MxVbStrShfStr."
Function ShfDotSeg$(OLn$)
ShfDotSeg = ShfBef(OLn, ".")
End Function

Function ShfBktStr$(OLn$)
If FstChr(OLn) = "(" Then
    ShfBktStr = BetBkt(OLn)
    OLn = AftBkt(OLn$)
End If
End Function

Function RmvLasChrzLis$(S, ChrLis$) ' Rmv fst chr if it is in ChrLis
If HasSubStr(ChrLis, LasChr(S)) Then
    RmvLasChrzLis = RmvLasChr(S)
Else
    RmvLasChrzLis = S
End If
End Function

Function TakChr$(S, ChrLis$) ' Ret fst chr if it is in ChrLis
If HasSubStr(ChrLis, FstChr(S)) Then TakChr = FstChr(S)
End Function

Function RmvFstChrzLis$(S, ChrLis$)
If HasSubStr(ChrLis, FstChr(S)) Then
    RmvFstChrzLis = RmvFstChr(S)
Else
    RmvFstChrzLis = S
End If
End Function
Function ShfChr$(OLn$, ChrList$)
Dim C$: C = TakChr(OLn, ChrList)
If C = "" Then Exit Function
ShfChr = C
OLn = Mid(OLn, 2)
End Function

Function ShfEq(OLn$) As Boolean
ShfEq = ShfTermX(OLn, "=")
End Function

Function ShfTy(OLn$) As Boolean
ShfTy = ShfTermX(OLn, "Ty")
End Function

Function ShfBetBkt$(OLn$) ' FstChr of @OLn should be (, remove the (....) as @OLn and return the text bet bkt, else thw error
If FstChr(OLn) <> "(" Then Thw CSub, "FstChr <> (", "OLn", OLn
Dim A$(): A = BrkBkt123(OLn)
If A(0) <> "" Then LgcEr "ShfBetBkt", "OLn is tested with ( at pos-1, and BrkBkt123 does not return A(0) as empty", "OLn", OLn
ShfBetBkt = A(1)
OLn = A(2)
End Function

Function ShfBkt(OLn$) As Boolean ' IF Fst2Chr of @OLn is (), Set @OLn with rmv 2 char and return true
ShfBkt = ShfPfx(OLn, "()")
End Function

Function ShfPfx(OLn$, Pfx$) As Boolean
If HasPfx(OLn, Pfx) Then
    OLn = RmvPfx(OLn, Pfx)
    ShfPfx = True
End If
End Function

Function ShfSfx(OLn$, Sfx$) As Boolean
If HasSfx(OLn, Sfx) Then
    OLn = RmvSfx(OLn, Sfx)
    ShfSfx = True
End If
End Function
Function ShfPfxy$(OLn$, Pfxy$())
Dim O$: O = PfxzAy(OLn, Pfxy): If O = "" Then Exit Function
ShfPfxy = O
OLn = RmvPfx(OLn, O)
End Function

Function ShfPfxyS$(OLn$, Pfxy$(), Optional BlnkVal$) ' Shf Pfx ay with space :: if @OLn one of pfx in @Pfxy + spc, return that pfx and remove that pfx+spc and LTrim into @Ln
Dim O$: O = PfxzAySpc(OLn, Pfxy): If O = "" Then ShfPfxyS = BlnkVal: Exit Function
ShfPfxyS = O
OLn = RmvPfxSpc(OLn, O)
End Function

Function ShfPfxSpc(OLn$, Pfx$) As Boolean
If HitPfxSpc(OLn, Pfx) Then
    OLn = LTrim(Mid(OLn, Len(Pfx) + 2))
    ShfPfxSpc = True
End If
End Function

Private Sub ShfBktStr__Tst()
Dim A$, Ept1$
A$ = "(O$()) As X": Ept = "O$()": Ept1 = " As X": GoSub Tst
Exit Sub
Tst:
    Act = ShfBktStr(A)
    C
    Ass A = Ept1
    Return
End Sub

Private Sub ShfPfx__Tst()
Dim O$: O = "AA{|}BB "
Ass ShfPfx(O, "{|}") = "AA"
Ass O = "BB "
End Sub
