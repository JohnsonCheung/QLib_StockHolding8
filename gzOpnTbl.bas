Attribute VB_Name = "gzOpnTbl"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzOpnTbl."
Sub OpnTbl_SkuRepackMulti():     DoCmd.OpenTable "SkuRepackMulti": End Sub
Sub OpnTbl_SkuTaxBy3rdParty():   DoCmd.OpenTable "SkuTaxBy3rdParty": End Sub
Sub OpnTbl_SkuNoLongerTax():     DoCmd.OpenTable "SkuNoLongerTax": End Sub
Sub OpnTbl_YpStk():              DoCmd.OpenTable "YpStk":        End Sub
Sub OpnTbl_BusArea():            DoCmd.OpenTable "PHLBus":       End Sub
Sub OpnTbl_PHTarMthsL1():  XL1:  DoCmd.OpenTable "PHTarMthsL1":  End Sub
Sub OpnTbl_PHTarMthsL2():  XL2:  DoCmd.OpenTable "PHTarMthsL2":  End Sub
Sub OpnTbl_PHTarMthsL3():  XL3:  DoCmd.OpenTable "PHTarMthsL3":  End Sub
Sub OpnTbl_PHTarMthsL4():  XL4:  DoCmd.OpenTable "PHTarMthsL4":  End Sub
Sub OpnTbl_PHTarMthsSku(): XSku: DoCmd.OpenTable "PHTarMthsSku": End Sub
Sub OpnTbl_PHTarMthsBus(): XBus: DoCmd.OpenTable "PHTarMthsBus": End Sub
Sub OpnTbl_PHTarMthsStm():       DoCmd.OpenTable "PHTarMthsStm": End Sub
Private Sub XSku()
DoCmd.SetWarnings False
RunCQ "SELECT Distinct Co,Sku into [#A] from OH"
RunCQ "Insert into PHTarMthsSku (Co,Sku) select x.Co,x.Sku" & _
" from [#A] x" & _
" left join [PHTarMthsSku] a on a.Co=x.Co and a.Sku=x.Sku" & _
" where a.Sku is null"
RunCQ "Drop table [#A]"
End Sub
Private Sub XBus()
DoCmd.SetWarnings False
RunCQ "Select Distinct Co into [#Co] from CoStm"
RunCQ "SELECT Co,Stm,BusArea into [#A] from [#Co],PHLBus"
RunCQ "Insert into PHTarMthsBus (Co,Stm,BusArea) select x.Co,x.Stm,x.BusArea" & _
" from [#A] x" & _
" left join [PHTarMthsBus] a on a.Co=x.Co and a.Stm=x.Stm and a.BusArea=x.BusArea" & _
" where a.Co is null"
DrpCTT "#Co #A"
End Sub

Private Sub XL4()
DoCmd.SetWarnings False
RunCQ "Select Co,x.Sku,ProdHierarchy,Left(ProdHierarchy,10) as PHL4 into [#A] from PHTarMthsSku x inner join Sku a on x.Sku=a.Sku"
RunCQ "Select Distinct Stm into [#Stm] from CoStm group by Stm"
RunCQ "Select Distinct Co,Stm,PHL4 into [#B] from [#A],[#Stm] group by Co,Stm,PHL4"
RunCQ "Insert Into PHTarMthsL4 (Co,Stm,PHL4) select x.Co,x.Stm,x.PHL4 from [#B] x" & _
" left join [PHTarMthsL4] a on a.Co=x.Co and a.Stm=x.Stm and a.PHL4=x.PHL4" & _
" where a.Co is null"
'DrpTT "#A #B #Stm"
End Sub

Private Sub XL3()
DoCmd.SetWarnings False
RunCQ "Select Co,x.Sku,ProdHierarchy,Left(ProdHierarchy,7) as PHL3 into [#A] from PHTarMthsSku x inner join Sku a on x.Sku=a.Sku"
RunCQ "Select Distinct Stm into [#Stm] from CoStm group by Stm"
RunCQ "Select Distinct Co,Stm,PHL3 into [#B] from [#A],[#Stm] group by Co,Stm,PHL3"
RunCQ "Insert Into PHTarMthsL3 (Co,Stm,PHL3) select x.Co,x.Stm,x.PHL3 from [#B] x" & _
" left join [PHTarMthsL3] a on a.Co=x.Co and a.Stm=x.Stm and a.PHL3=x.PHL3" & _
" where a.Co is null"
End Sub

Private Sub XL2()
DoCmd.SetWarnings False
RunCQ "Select Co,x.Sku,ProdHierarchy,Left(ProdHierarchy,4) as PHL2 into [#A] from PHTarMthsSku x inner join Sku a on x.Sku=a.Sku"
RunCQ "Select Distinct Stm into [#Stm] from CoStm group by Stm"
RunCQ "Select Distinct Co,Stm,PHL2 into [#B] from [#A],[#Stm] group by Co,Stm,PHL2"
RunCQ "Insert Into PHTarMthsL2 (Co,Stm,PHL2) select x.Co,x.Stm,x.PHL2 from [#B] x" & _
" left join [PHTarMthsL2] a on a.Co=x.Co and a.Stm=x.Stm and a.PHL2=x.PHL2" & _
" where a.Co is null"
End Sub

Private Sub XL1()
DoCmd.SetWarnings False
RunCQ "Select Co,x.Sku,ProdHierarchy,Left(ProdHierarchy,2) as PHL1 into [#A] from PHTarMthsSku x inner join Sku a on x.Sku=a.Sku"
RunCQ "Select Distinct Stm into [#Stm] from CoStm group by Stm"
RunCQ "Select Distinct Co,Stm,PHL1 into [#B] from [#A],[#Stm] group by Co,Stm,PHL1"
RunCQ "Insert Into PHTarMthsL1 (Co,Stm,PHL1) select x.Co,x.Stm,x.PHL1 from [#B] x" & _
" left join [PHTarMthsL1] a on a.Co=x.Co and a.Stm=x.Stm and a.PHL1=x.PHL1" & _
" where a.Co is null"
End Sub
