Attribute VB_Name = "gzTmpScOH"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzTmpScOH."
'Sub T_ScOH():TmpScOH7_ByYmd Ymd(18, 12, 31):End Sub
Sub TmpScOH7_ByYmd(A As Ymd)
'Called by UpdTbPHStkDays7_FldStkDays_andRemSc
TmpScOH7ByBexp OHYmdBexp(A)
End Sub
Sub TmpScOH7_ByCoYmd(A As CoYmd)
TmpScOH7ByBexp CoOHYmdBexp(A)
End Sub

Sub TmpScOH7ByBexp(B$) 'Crt #ScOH{7} by @Wh
DoCmd.SetWarnings False
'Oup: $ScOH{7}
'Inp: OH = YY MM DD Co Sku | Btl
'Ref: qSku_Main    = Sku [Btl/SC]
'Ref:
'Oup: $ScOHSku  Co Stm Sku  | SC
'Oup: $ScOHL4   Co Stm PHL4 | SC
'Oup: $ScOHL3   Co Stm PHL3 | SC
'Oup: $ScOHL2   Co Stm PHL2 | SC
'Oup: $ScOHL1   Co Stm PHL1 | SC
'Oup: $ScOHBus  Co Stm BusArea | SC
'Oup: $ScOHStm  Co Stm      | SC

'$ScOHSku
RunCQ FmtQQ("Select Distinct Co,Sku,Sum(x.Btl) as Btl,CDbl(0) As SC into [$ScOHSku] from OH x Where ? Group by Co,Sku", B)
RunCQ "Alter Table [$ScOHSku] add Column Stm Text(1)"
RunCQ "Update [$ScOHSku] x inner join qSku_Main a on x.Sku=a.Sku set SC = Btl/[Btl/SC],x.Stm=a.Stm"
RunCQ "Alter Table [$ScOHSku] Drop Column Btl"

'Oup: $ScOHL4
'Fm : $ScOHSku
'Tmp: #A
'Ref: qSku_Main      -> Sku PLH4
RunCQ "Select Co,x.Sku,PHL4,SC Into [#A] from [$ScOHSku] a left join qSku_Main x on a.SKu=x.SKu"
RunCQ "Select Distinct x.Co,Stm,PHL4,Sum(x.SC) as SC" & _
" into [$ScOHL4]" & _
" from [#A] x" & _
" left join [$ScOHSku] a on a.Sku=x.Sku" & _
" group by x.Co,Stm,PHL4"
RunCQ "Drop Table [#A]"

'Oup: $ScOHL3  | Fm : $ScOHL4
'Oup: $ScOHL2  | Fm : $ScOHL3
'Oup: $ScOHL1  | Fm : $ScOHL2
'Oup: $ScOHStm | Fm : $ScOHL1
RunCQ "Select Distinct Co,Stm,Left(PHL4,7) as PHL3,Sum(x.SC) as SC into [$ScOHL3] from [$ScOHL4] x Group By Co,Stm,Left(PHL4,7)"
RunCQ "Select Distinct Co,Stm,Left(PHL3,4) as PHL2,Sum(x.SC) as SC into [$ScOHL2] from [$ScOHL3] x Group By Co,Stm,Left(PHL3,4)"
RunCQ "Select Distinct Co,Stm,Left(PHL2,2) as PHL1,Sum(x.SC) as SC into [$ScOHL1] from [$ScOHL2] x Group By Co,Stm,Left(PHL2,2)"
RunCQ "Select Distinct Co,Stm,Sum(x.SC) as SC into [$ScOHStm] from [$ScOHL1] x Group by Co,Stm"

'Oup: $ScOHBus | Fm : $ScOHSku
'Ref: Sku_Main  = Sku Stm BusArea
RunCQ "Select Distinct Co,Stm,BusArea,Sum(x.SC) as SC" & _
" into [$ScOHBus]" & _
" from [$ScOHSku] x left join qSku_Main a on a.Sku=x.Sku" & _
" Group By Co,Stm,BusArea"
End Sub
