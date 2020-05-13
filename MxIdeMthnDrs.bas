Attribute VB_Name = "MxIdeMthnDrs"
Option Explicit
Option Compare Text
Const CNs$ = "Mth.Drs"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthnDrs."
Public Const MthnFF$ = "MdTy Mthn Mdn Mdy Ty"
Function MthnDrszM(M As CodeModule) As Drs
MthnDrszM = DwEq(SelDrs(MthnDrsP, MthnFF), "Mthn", Mdn(M))
End Function

Function MthnDrsP() As Drs
MthnDrsP = SelDrs(MthcDrsP, MthnFF)
End Function

Function MthnDrszV(V As Vbe) As Drs

End Function

Function MthnDrsV() As Drs
MthnDrsV = MthnDrszV(CVbe)
End Function

Function MthnDrsM() As Drs
MthnDrsM = MthnDrszM(CMd)
End Function
