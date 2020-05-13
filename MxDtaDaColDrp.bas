Attribute VB_Name = "MxDtaDaColDrp"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaColDrp."

Function DrpColzDrsCC(A As Drs, CC$) As Drs
DrpColzDrsCC = DrpColzDrsFny(A, SyzSS(CC))
End Function

Function DrpColzDyIxy(Dy(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
   Push DrpColzDyIxy, AeIxy(Dr, Ixy)
Next
End Function

Function DrpColzDrsFny(D As Drs, Fny$()) As Drs
Dim IxAll&(): IxAll = LngSno(UB(D.Fny))
Dim IxToExl&():      IxToExl = Ixy(D.Fny, Fny)
Dim IxSel&(): IxSel = MinusAy(IxAll, IxToExl)
Dim ODy(): ODy = SelDy(D.Dy, IxSel)
DrpColzDrsFny = Drs(MinusSy(D.Fny, Fny), ODy)
End Function
