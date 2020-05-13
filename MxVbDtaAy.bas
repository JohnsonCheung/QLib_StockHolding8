Attribute VB_Name = "MxVbDtaAy"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxVbDtaAy."
Type StrColy: Coly() As Variant: End Type ' Each ele is Sy
Function StrColy(Coly()) As StrColy
With StrColy
    .Coly = Coly ' Each ele of @V should be Sy
End With
End Function
Function DtzAy(Ay, Optional Fldn$ = "Itm", Optional DtNm$ = "Ay") As Dt
Dim Dy(), J&
For J = 0 To UB(Ay)
    PushI Dy, Array(Ay(J))
Next
DtzAy = Dt(DtNm, Sy(Fldn), Dy)
End Function
