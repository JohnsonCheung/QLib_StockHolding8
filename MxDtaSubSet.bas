Attribute VB_Name = "MxDtaSubSet"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaSubSet."

Function DeIn(A As Drs, C, InVy) As Drs
Const CSub$ = CMod & "DeIn"
Dim Ix&: Ix = IxzAy(A.Fny, C)
If Not IsArray(InVy) Then Thw CSub, "Given InVy is not an array", "Ty-InVy", TypeName(InVy)
Dim Dr, Dy(): For Each Dr In Itr(A.Dy)
    If Not HasEle(InVy, Dr(Ix)) Then
        PushI Dy, Dr
    End If
Next
DeIn = Drs(A.Fny, Dy)
End Function

Function DeRxy(A As Drs, Rxy&()) As Drs
Dim Dy(): Dy = A.Dy
Dim ODy()
Dim Rix: For Each Rix In IxItrExl(UB(A.Dy), Rxy)
    PushI ODy, Dy(Rix)
Next
DeRxy = Drs(A.Fny, ODy)
End Function

Function DeVap(D As Drs, CC$, ParamArray Vap()) As Drs
'Fm D : ..@CC..            ! to be selected.  It has col-@CC
'Fm Vap : #Val-Ay-of-Param ! to select what rec in @D to be selected
'Ret    : ..@D..           ! sam stru as @D.  Subset of @D.  @@
Dim Vy(): Vy = Vap
DeVap = DeVy(D, CC, Vy)
End Function

Function DeVy(D As Drs, CC$, Vy) As Drs
'Fm D  : ..@CC..            ! to be selected.  It has col-@CC
'Fm Vy : #Val-Ay-of-Param ! to select what rec in @D to be selected
'Ret   : ..@D..           ! sam stru as @D.  Subset of @D.  @@
Dim KeyDy(): KeyDy = SelDrs(D, CC).Dy
Dim Rxy&(): Rxy = RxyeDyVy(KeyDy, Vy)
Dim ODy(): ODy = AwIxy(D.Dy, Rxy)
DeVy = Drs(D.Fny, ODy)
End Function
Function Dw2Eq(D As Drs, C2$, V1, V2) As Drs
Dim A$, B$: AsgTRst C2, A, B
Dw2Eq = DwEq(DwEq(D, A, V1), B, V2)
End Function
Function DwNBlnk(D As Drs, C$) As Drs
DwNBlnk = DwNe(D, C, "")
End Function

Function Dw2EqE(D As Drs, C2$, V1, V2) As Drs
Dw2EqE = DrpColzDrsCC(Dw2Eq(D, C2, V1, V2), C2)
End Function

Function Dw2Patn(A As Drs, TwoC$, Patn1$, Patn2$) As Drs
Dim C1$, C2$: AsgBrkSpc TwoC, C1, C2
Dw2Patn = DwPatn(DwPatn(A, C1, Patn1), C2, Patn2)
End Function

Function Dw3Eq(D As Drs, C3$, V1, V2, V3) As Drs
Dim A$, B$, C$: AsgTTRst C3, A, B, C
Dw3Eq = DwEq(DwEq(DwEq(D, A, V1), B, V2), C, V3)
End Function

Function Dw3EqE(D As Drs, C3$, V1, V2, V3) As Drs
Dw3EqE = DrpColzDrsCC(Dw3Eq(D, C3, V1, V2, V3), C3)
End Function

Function DwColGt(A As Drs, C$, V) As Drs
Dim Dy(), Ix%, Fny$()
Fny = A.Fny
'Ix = Ixy(Fny, C)
DwColGt = Drs(Fny, DywColGt(A.Dy, Ix, V))
End Function

Function DwColNe(A As Drs, C$, V) As Drs
Dim Dy(), Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C)
DwColNe = Drs(Fny, DywColNe(A.Dy, Ix, V))
End Function

Function DwDup(D As Drs, DupFF$) As Drs
DwDup = DwRxy(D, DupRecRxyzFF(D, DupFF))
End Function

Function DwDupC(A As Drs, C$) As Drs
Dim Dup(): Dup = AwDup(ColzDrs(A, C))
DwDupC = DwIn(A, C, Dup)
End Function

Function DwBlnk(A As Drs, C$) As Drs
DwBlnk = DwEq(A, C, "")
End Function

Function DwNBlnkC(A As Drs, NBlnkC$) As Drs
DwNBlnkC = DwNe(A, NBlnkC, "")
End Function

Function DwCC2Eq(D As Drs, CC2$, EqV1, EqV2) As Drs
Dim C1$, C2$: AsgBrkSpc CC2, C1, C2
DwCC2Eq = DwEq(DwEq(D, C1, EqV1), C2, EqV2)
End Function

Function DwCC2EqExl(D As Drs, CC2$, EqV1, EqV2) As Drs
DwCC2EqExl = DrpColzDrsCC(DwCC2Eq(D, CC2, EqV1, EqV2), CC2)
End Function

Function DwEq(A As Drs, ByC$, Eq) As Drs
Dim Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, ByC, ThwEr:=EiThwEr)
DwEq = Drs(Fny, DywEq(A.Dy, Ix, Eq))
End Function

Function DwEqStr(A As Drs, C$, Str$) As Drs
If Str = "" Then DwEqStr = A: Exit Function
DwEqStr = DwEq(A, C, Str)
End Function

Function DwSubStr(A As Drs, C$, SubStr) As Drs
Dim Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C, ThwEr:=EiThwEr)
DwSubStr = Drs(Fny, DywSubStr(A.Dy, Ix, SubStr))
End Function

Function DwLik(A As Drs, C$, Lik) As Drs
Dim Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C, ThwEr:=EiThwEr)
DwLik = Drs(Fny, DywLik(A.Dy, Ix, Lik))
End Function

Function DwFalse(A As Drs, C$) As Drs
DwFalse = DwEq(A, C, False)
End Function
Function DwFalseExl(A As Drs, C$) As Drs
DwFalseExl = DrpColzDrsCC(DwFalse(A, C), C)
End Function

Function DwEqExl(A As Drs, C$, V) As Drs
DwEqExl = DrpColzDrsCC(DwEq(A, C, V), C)
End Function

Function DwEqSel(A As Drs, C$, V, Sel$) As Drs
DwEqSel = SelDrs(DwEq(A, C, V), Sel)
End Function

Function DwFFNe(A As Drs, F1, F2) As Drs 'FFNe = Two Fld Not Eq
Dim Fny$()
Fny = A.Fny
'DwFFNe = Drs(Fny, DyWhCCNe(A.Dy, Ixy(Fny, F1), Ixy(Fny, F2)))
End Function

Function DwFldEqV(A As Drs, F, Eqval) As Drs
'DwFldEqV = Drs(A.Fny, DyWh(A.Dy, Ixy(A.Fny, F), EqVal))
End Function

Function DwIn(A As Drs, C, InVy) As Drs
Dim Ix&: Ix = IxzAy(A.Fny, C)
DwIn = Drs(A.Fny, DywIn(A.Dy, Ix, InVy))
End Function

Function DwNe(A As Drs, C$, V) As Drs
DwNe = DwColNe(A, C, V)
End Function

Function DwNeSel(A As Drs, C$, V, Sel$) As Drs
DwNeSel = SelDrs(DwNe(A, C, V), Sel)
End Function

Function DwRxy(A As Drs, Rxy&()) As Drs
DwRxy = Drs(A.Fny, CvAv(AwIxy(A.Dy, Rxy)))
End Function

Function DwPatn(A As Drs, C$, Patn$) As Drs
If Patn = "" Then DwPatn = A: Exit Function
Dim R As RegExp: Set R = Rx(Patn)
Dim Ix%: Ix = IxzAy(A.Fny, C, ThwEr:=EiThwEr)
Dim Dy(), Dr: For Each Dr In Itr(A.Dy)
    If HitRx(Dr(Ix), R) Then PushI Dy, Dr
Next
DwPatn = Drs(A.Fny, Dy)
End Function

Function DwPfx(D As Drs, C$, Pfx) As Drs
DwPfx = Drs(D.Fny, DywPfx(D.Dy, IxzAy(D.Fny, C), Pfx))
End Function

Function DwTop(A As Drs, Optional NTop& = 50) As Drs
DwTop = Drs(A.Fny, CvAv(FstNEle(A.Dy, NTop)))
End Function

Function DwVap(D As Drs, CC$, ParamArray Vap()) As Drs
'Fm D : ..@CC..            ! to be selected.  It has col-@CC
'Fm Vap : #Val-Ay-of-Param ! to select what rec in @D to be returned
'Ret    : ..@D..           ! sam stru as @D.  Subset of @D.  @@
Dim Vy(): Vy = Vap
Dim KeyDy(): KeyDy = SelDrs(D, CC).Dy
Dim Rxy&(): Rxy = RxywDyVy(KeyDy, Vy)
Dim ODy(): ODy = AwIxy(D.Dy, Rxy)
DwVap = Drs(D.Fny, ODy)
End Function

Function DwEqFny(A As Drs, C$, V, SelFny$()) As Drs
DwEqFny = DrszSelFny(DwEq(A, C, V), SelFny)
End Function

Function DwInSel(A As Drs, C, InVy, Sel$) As Drs
DwInSel = SelDrs(DwIn(A, C, InVy), Sel)
End Function
