Attribute VB_Name = "gzTmpFc"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzTmpFc."
'Sub T_Fc(): TmpFc_ByYM YM(19, 12): End Sub
Sub TmpFc_ByYM(A As YM)
RRoll WhereFc(A.Y, A.Y)
End Sub
Sub TmpFc_ByCoYM(A As CoYM)
RRoll WhereFcCo(A)
End Sub
Sub TmpFcStm(A As StmYM)
RRoll WhereFcStm(A)
End Sub

Private Sub RRoll(W$)
'Oup: Create [$Fc{L7}] from FcSku & , where {7} is Sku Stm Bus L1..4
'       Stm   = Co Stm         {Dta}
'       Bus   = Co Stm Bus     {Dta}
'       L1..4 = Co Stm PHL1..4 {Dta}
'       Sku   = Co Stm Sku     {Dta}
'       where {Dta} is M01..15
'     Select from FcSku into $FcSku and up to above 6 level
'     It is called by subr-Fc_Calc & subr-Fc_Export.
StsQry "Forecast"
DoCmd.SetWarnings False
Const Sum$ = _
"Sum(x.M01) As M01,Sum(x.M02) as M02," & _
"Sum(x.M03) As M03,Sum(x.M04) as M04," & _
"Sum(x.M05) As M05,Sum(x.M06) as M06," & _
"Sum(x.M07) As M07,Sum(x.M08) as M08," & _
"Sum(x.M09) As M09,Sum(x.M10) as M10," & _
"Sum(x.M11) As M11,Sum(x.M12) as M12," & _
"Sum(x.M13) As M13,Sum(x.M14) as M14," & _
"Sum(x.M15) As M15"

'== $FcSku
RunCQ "select Distinct Co,Stm,Sku,M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12,M13,M14,M15" & _
" Into [$FcSku]" & _
" from [FcSku] x" & W

'== #Sku
RunCQ "Select Sku,BusArea,PHL4 into [#Sku] from qSku_Main"

'== $Fc{6}
RunCQ "select Distinct Co,Stm,BusArea,             " & Sum & " Into [$FcBus] from FcSku x left join [#Sku] a on x.Sku=a.Sku" & W & " group by Co,Stm,BusArea"
RunCQ "Select Distinct Co,Stm,PHL4,                " & Sum & " Into [$FcL4]  from FcSku x left join [#Sku] a on x.Sku=a.Sku" & W & " group by Co,Stm,PHL4"
RunCQ "select Distinct Co,Stm,Left(PHL4,7) as PHL3," & Sum & " Into [$FcL3]  from [$FcL4] x group by Co,Stm,Left(PHL4,7)"
RunCQ "select Distinct Co,Stm,Left(PHL3,4) as PHL2," & Sum & " Into [$FcL2]  from [$FcL3] x group by Co,Stm,Left(PHL3,4)"
RunCQ "Select Distinct Co,Stm,Left(PHL2,2) as PHL1," & Sum & " Into [$FcL1]  from [$FcL2] x group by Co,Stm,Left(PHL2,2)"
RunCQ "select Distinct Co,Stm,                     " & Sum & " Into [$FcStm] from [$FcL1] x group by Co,Stm"

RunCQ "Drop Table [#Sku]"
End Sub
