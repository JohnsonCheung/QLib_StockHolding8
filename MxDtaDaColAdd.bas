Attribute VB_Name = "MxDtaDaColAdd"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaColAdd."

Function AddColzDy(Dy(), ValToBeAddAsLasCol) As Variant()
'Ret : a new :Dy with a col of value all eq to @ValToBeAddAsLasCol at end
Dim O(): O = Dy
Dim ToU&
    ToU = NColzDy(Dy)
Dim J&, Dr
For Each Dr In Itr(O)
    ReDim Preserve Dr(ToU)
    Dr(ToU) = ValToBeAddAsLasCol
    O(J) = Dr
    J = J + 1
Next
AddColzDy = O
End Function

Function AddColzDyAv(Dy(), Av()) As Variant()
Dim O(): O = Dy
Dim ToU&
    ToU = NColzDy(Dy) + 1
Dim J&, Dr, I1%, I2%
I2 = ToU
I1 = I2 - 1
For Each Dr In Itr(O)
    ReDim Preserve Dr(ToU)
    PushAy Dr, Av
    O(J) = Dr
    J = J + 1
Next
AddColzDyAv = O
End Function

Function AddColzDyBy(Dy(), Optional ByNCol% = 1) As Variant()
Dim NewU&
    NewU = NColzDy(Dy) + ByNCol - 1
Dim O()
    Dim UDy&: UDy = UB(Dy)
    O = AyReSzU(O, UDy)
    Dim J&
    For J = 0 To UDy
        O(J) = AyReSzU(Dy(J), NewU)
    Next
AddColzDyBy = O
End Function

Function AddColzDyC(Dy(), C) As Variant()
If Si(Dy) = 0 Then Exit Function
Dim O(): O = AddColzDyBy(Dy)
    Dim UCol%: UCol = UB(Dy(0))
    Dim J&
    For J = 0 To UB(Dy)
       O(J)(UCol) = C
    Next
AddColzDyC = O
End Function

Function AddColzMap(A As Drs, NewFldEqFunQuoFmFldSsl$) As Drs
Dim NewFldVy(), FmVy()
Dim I, S$, NewFld$, Fun$, FmFld$
For Each I In SyzSS(NewFldEqFunQuoFmFldSsl)
    S = I
    NewFld = Bef(S, "=")
    Fun = IsBet(S, "=", "(")
    FmFld = BetBkt(S)
    FmVy = ColzDrs(A, FmFld)
    NewFldVy = AyzMap(FmVy, Fun)
    Stop '
Next
End Function

Function AddColzVy(A As Drs, ColNm$, FldVy) As Drs
Dim Fny$(): Fny = AddAyEle(A.Fny, ColNm)
Dim AtIx&: AtIx = UB(Fny)
Dim Dy(): Dy = AddColzDyFldVy(A.Dy, FldVy, AtIx)
AddColzVy = Drs(Fny, Dy)
End Function

Function CntColEq&(A As Drs, C$, V)
Dim I%: I = IxzAy(A.Fny, C)
Dim O&, Dr: For Each Dr In Itr(A.Dy)
    If Dr(I) = V Then O = O + 1
Next
CntColEq = O
End Function

Function CntColNe&(A As Drs, C$, V)
Dim I%: I = IxzAy(A.Fny, C)
Dim O&, Dr: For Each Dr In Itr(A.Dy)
    If Dr(I) <> V Then O = O + 1
Next
CntColNe = O
End Function
Function ColNoSng(A As Drs, C$) As Drs
'@A : has a column-C
'Ret   : sam stru as A and som row removed.  rmv row are its col C value is Single. @@
Dim Col(): Col = ColzDrs(A, C)
Dim Sng(): Sng = AwSng(Col)
ColNoSng = DeIn(A, C, Sng)
End Function
Function FstDrWhColEq(A As Drs, C$, Eqval) As Variant()
Dim Cix%: Cix = IxzAy(A.Fny, C)
Dim Rix&: Rix = RixWhColEq(A.Dy, Cix, Eqval)
FstDrWhColEq = A.Dy(Rix)
End Function

Function FstRecWhColEq(A As Drs, C$, Eqval) As Rec
FstRecWhColEq = Rec(A.Fny, FstDrWhColEq(A, C, Eqval))
End Function

Function RixWhColEq&(Dy, Cix, Eqval)
Dim R&
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(Cix) = Eqval Then RixWhColEq = R
    R = R + 1
Next
RixWhColEq = -1
End Function
Function HasColEq(A As Drs, C$, V) As Boolean
HasColEq = HasColEqzDy(A.Dy, IxzAy(A.Fny, C), V)
End Function

Function InsColzDyAv(Dy(), Av()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    PushI InsColzDyAv, AddAy(Av, Dr)
Next
End Function

Function InsColzDy(A(), V, Optional At& = 0) As Variant()
Dim Dr
For Each Dr In Itr(A)
    PushI InsColzDy, InsBef(Dr, V, At)
Next
End Function

Function InsColzDyV2(A(), V1, V2) As Variant()
InsColzDyV2 = InsColzDyAv(A, Av(V1, V2))
End Function

Function InsColzDyV3(Dy(), V1, V2, V3) As Variant()
InsColzDyV3 = InsColzDyAv(Dy, Av(V1, V2, V3))
End Function

Function InsColzDyV4(A(), V1, V2, V3, V4) As Variant()
InsColzDyV4 = InsColzDyAv(A, Av(V1, V2, V3, V4))
End Function

Function RmvPfxzDrs(A As Drs, C$, Pfx$) As Drs
Dim Dr, ODy(), J&, I%
ODy = A.Dy
I = IxzAy(A.Fny, C)
For Each Dr In Itr(A.Dy)
    Dr(I) = RmvPfx(Dr(I), Pfx)
    ODy(J) = Dr
    J = J + 1
Next
RmvPfxzDrs = Drs(A.Fny, ODy)
End Function

Function RxyeDyVy(Dy(), Vy) As Long()
'Fm Dy: ! to be selected if it ne to @Vy.  It has sam NCol as Si-Vy
'Fm Vy : ! to select @Dy
'Ret   : Rxy of @Dy if the rec ne @Vy
Dim Rix&, Dr: For Each Dr In Itr(Dy)
    If Not IsEqAy(Dr, Vy) Then PushI RxyeDyVy, Rix
    Rix = Rix + 1
Next
End Function

Function RxywDyVy(Dy(), Vy) As Long()
'Fm Dy: ! to be selected if it eq to @Vy.  It has sam NCol as Si-Vy
'Fm Vy : ! to select @Dy
'Ret   : Rxy of @Dy if the rec eq @Vy
Dim Rix&, Dr: For Each Dr In Itr(Dy)
    If IsEqAy(Dr, Vy) Then PushI RxywDyVy, Rix
    Rix = Rix + 1
Next
End Function

Function DwTopN(A As Drs, Optional N = 50) As Drs
If N <= 0 Then DwTopN = A: Exit Function
DwTopN = Drs(A.Fny, CvAv(FstNEle(A.Dy, N)))
End Function

Function ValzDrs(A As Drs, C$, V, ColNm$)
Const CSub$ = CMod & "ValzDrs"
Dim Dr, Ix%, IxRet%
Ix = IxzAy(A.Fny, C)
IxRet = IxzAy(A.Fny, ColNm)
For Each Dr In Itr(A.Dy)
    If Dr(Ix) = V Then
        ValzDrs = Dr(IxRet)
        Exit Function
    End If
Next
Thw CSub, "In Drs, there is no record with Col-A eq Value-B, so no Col-C is returened", "Col-A Value-B Col-C Drs-Fny Drs-NRec", C, V, ColNm, A.Fny, NReczDrs(A)
End Function
Function AddColByNewDy(A As Drs, AddFF$, NewDy()) As Drs
AddColByNewDy = Drs(AddSy(A.Fny, SyzSS(AddFF)), NewDy)
End Function

Function AddCol(A As Drs, C$, V) As Drs
Dim Dr, Dy()
For Each Dr In Itr(A.Dy)
    PushI Dr, V
    PushI Dy, Dr
Next
AddCol = AddColzFFDy(A, C, Dy)
End Function

Function AddColzDyCC(Dy(), V1, V2) As Variant()
AddColzDyCC = AddColzDyAv(Dy, Av(V1, V2))
End Function

Function AddColzDyC3(Dy(), V1, V2, V3) As Variant()
AddColzDyC3 = AddColzDyAv(Dy, Av(V1, V2, V3))
End Function

Function AddColz2(A As Drs, FF2$, C1, C2) As Drs
Dim Fny$(), Dy()
Fny = AddAy(A.Fny, Termy(FF2))
Dy = AddColzDyCC(A.Dy, C1, C2)
AddColz2 = Drs(Fny, Dy)
End Function

Function AddColz3(A As Drs, FF3$, C1, C2, C3) As Drs
Dim Fny$(), Dy()
Fny = AddAy(A.Fny, Termy(FF3))
Dy = AddColzDyC3(A.Dy, C1, C2, C3)
AddColz3 = Drs(Fny, Dy)
End Function

Function AddColzDyFldVy(Dy(), FldVy, AtIx&) As Variant()
Const CSub$ = CMod & "AddColzDyFldVy"
Dim Dr, J&, O(), U&
U = UB(FldVy)
If U = -1 Then Exit Function
If U <> UB(Dy) Then Thw CSub, "Row-in-Dy <> Si-FldVy", "Row-in-Dy Si-FldVy", Si(Dy), Si(FldVy)
ReDim O(U)

For Each Dr In Itr(Dy)
    If Si(Dr) > AtIx Then Thw CSub, "Some Dr in Dy has bigger size than AtIx", "DrSz AtIx", Si(Dr), AtIx
    ReDim Preserve Dr(AtIx)
    Dr(AtIx) = FldVy(J)
    O(J) = Dr
    J = J + 1
Next
AddColzDyFldVy = O
End Function
