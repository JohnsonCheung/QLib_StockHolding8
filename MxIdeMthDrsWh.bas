Attribute VB_Name = "MxIdeMthDrsWh"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeMthDrsWh."


Function PubFunDrszP(P As VBProject) As Drs
PubFunDrszP = SelDrs(Dw2Eq(MthcDrszP(P), "Mdy MdTy", "Pub", "Std"), PubMthFF)
End Function
