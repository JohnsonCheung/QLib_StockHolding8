Attribute VB_Name = "MxVbRunFun"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxVbRunFun."
Public Const PFunFF$ = MthFF
Function FunDrsP() As Drs
FunDrsP = DwEq(DwEq(CacMthDrsP, "MdTy", "Std"), "Ty", "Fun")
End Function

Function SubDrsP() As Drs
SubDrsP = DwEq(DwEq(CacMthDrsP, "MdTy", "Std"), "Ty", "Sub")
End Function

Function PSubP() As Drs
PSubP = DwEq(FunDrsP, "Ty", "Sub")
End Function

Function PFunDrsP() As Drs
PFunDrsP = DwEqExl(FunDrsP, "Ty", "Fun")
End Function

Function PFunDrsWhMthnPatn(P$) As Drs
PFunDrsWhMthnPatn = DwPatn(PFunDrsP, "Mthn", P)
End Function

Function PPrpDrs() As Drs
PPrpDrs = DwIn(PFunDrsP, "Ty", SyzSS("Get Let Set"))
End Function

Function PPrpDrsWiPm() As Drs
Dim A As Drs: A = AddMthColMthPm(PPrpDrsWiPm)
PPrpDrsWiPm = DwEqExl(A, "MthPm", "")
End Function
