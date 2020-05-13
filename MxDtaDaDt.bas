Attribute VB_Name = "MxDtaDaDt"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaDt."

Function CsyzDt(A As Dt) As String()
Dim Dy(): Dy = A.Dy
Push CsyzDt, JnComma(AmQuoDbl(A.Fny))
Dim Dr: For Each Dr In A.Dy
   PushI CsyzDt, Csl(Dr)
Next
End Function

Sub DmpDt(A As Dt)
DmpAy FmtDt(A)
End Sub

Function Dr(Dy(), R&) As Variant()
Dr = Dy(R)
End Function

Function DtDrpCol(A As Dt, CC$, Optional DtNm$) As Dt
DtDrpCol = DtzDrs(DrpColzDrsCC(DrszDt(A), CC), Dft(DtNm, A.DtNm))
End Function

Function DtReOrd(A As Dt, BySubFF$) As Dt
DtReOrd = DtzDrs(ReOrdCol(DrszDt(A), BySubFF), A.DtNm)
End Function

Function DtSelCol(A As Dt, CC$, Optional DtNm$) As Dt
DtSelCol = DtzDrs(SelDrs(DrszDt(A), CC), Dft(DtNm, A.DtNm))
End Function

Function DtByFF(DtNm$, FF$, Dy()) As Dt
DtByFF = Dt(DtNm, Ny(FF), Dy)
End Function

Property Get EmpDtAy() As Dt()
End Property

Function IsEmpDt(A As Dt) As Boolean
IsEmpDt = Si(A.Dy) = 0
End Function


Function NRowOfDt&(A As Dt)
NRowOfDt = Si(A.Dy)
End Function
