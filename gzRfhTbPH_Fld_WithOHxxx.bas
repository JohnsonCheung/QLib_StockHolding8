Attribute VB_Name = "gzRfhTbPH_Fld_WithOHxxx"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzRfhTbPH_Fld_WithOHxxx."
Sub RfhTbPH_Fld_WithOHxxx()
DoCmd.SetWarnings False
RunCQ "Select Distinct Sku into [#OHSkuHst] from OH"
RunCQ "Select Distinct Sku into [#OHSkuCur] from OH" & WhLasOH

RunCQ "Select Distinct ProdHierarchy as PH into [#OHPHCur] from SKU where SKU in (Select Sku From [#OHSkuCur])"
RunCQ "Select Distinct ProdHierarchy as PH into [#OHPHHst] from SKU where SKU in (Select Sku From [#OHSkuHst])"
RunCQ "Select Distinct CByte(4) as Lvl, Left(PH,10) as PHL4 into [#PHL4Cur] from [#OHPHCur]"
RunCQ "Select Distinct CByte(4) as Lvl, Left(PH,10) as PHL4 into [#PHL4Hst] from [#OHPHHst]"
RunCQ "Select Distinct CByte(3) as Lvl, Left(PH,7) as PHL3 into [#PHL3Cur] from [#OHPHCur]"
RunCQ "Select Distinct CByte(3) as Lvl, Left(PH,7) as PHL3 into [#PHL3Hst] from [#OHPHHst]"
RunCQ "Select Distinct CByte(2) as Lvl, Left(PHL3,4) as PHL2 into [#PHL2Cur] from [#PHL3Cur]"
RunCQ "Select Distinct CByte(2) as Lvl, Left(PHL3,4) as PHL2 into [#PHL2Hst] from [#PHL3Hst]"
RunCQ "Select Distinct CByte(1) as Lvl, Left(PHL2,2) as PHL1 into [#PHL1Cur] from [#PHL2Cur]"
RunCQ "Select Distinct CByte(1) as Lvl, Left(PHL2,2) as PHL1 into [#PHL1Hst] from [#PHL2Hst]"

RunCQ "Update ProdHierarchy x Left Join [#OHPHCur] a on x.PH=a.PH set WithOHCur=Not IsNull(a.PH) where Lvl=4"
RunCQ "Update ProdHierarchy x Left Join [#OHPHHst] a on x.PH=a.PH set WithOHHst=Not IsNull(a.PH) where Lvl=4"

RunCQ "Update ProdHierarchy x inner Join [#PHL4Cur] a on x.Lvl=a.Lvl and x.PH=a.PHL4 set WithOHCur=Not IsNull(a.PHL4)"
RunCQ "Update ProdHierarchy x inner Join [#PHL4Hst] a on x.Lvl=a.Lvl and x.PH=a.PHL4 set WithOHHst=Not IsNull(a.PHL4)"

RunCQ "Update ProdHierarchy x inner Join [#PHL3Cur] a on x.Lvl=a.Lvl and x.PH=a.PHL3 set WithOHCur=Not IsNull(a.PHL3)"
RunCQ "Update ProdHierarchy x inner Join [#PHL3Hst] a on x.Lvl=a.Lvl and x.PH=a.PHL3 set WithOHHst=Not IsNull(a.PHL3)"

RunCQ "Update ProdHierarchy x inner Join [#PHL2Cur] a on x.Lvl=a.Lvl and x.PH=a.PHL2 set WithOHCur=Not IsNull(a.PHL2)"
RunCQ "Update ProdHierarchy x inner Join [#PHL2Hst] a on x.Lvl=a.Lvl and x.PH=a.PHL2 set WithOHHst=Not IsNull(a.PHL2)"

RunCQ "Update ProdHierarchy x inner Join [#PHL1Cur] a on x.Lvl=a.Lvl and x.PH=a.PHL1 set WithOHCur=Not IsNull(a.PHL1)"
RunCQ "Update ProdHierarchy x inner Join [#PHL1Hst] a on x.Lvl=a.Lvl and x.PH=a.PHL1 set WithOHHst=Not IsNull(a.PHL1)"

RunCQ "Drop Table [#OHSkuHst]"
RunCQ "Drop Table [#OHSkuCur]"
RunCQ "Drop Table [#OHPHHst]"
RunCQ "Drop Table [#OHPHCur]"
RunCQ "Drop Table [#PHL1Cur]"
RunCQ "Drop Table [#PHL1Hst]"
RunCQ "Drop Table [#PHL2Cur]"
RunCQ "Drop Table [#PHL2Hst]"
RunCQ "Drop Table [#PHL3Cur]"
RunCQ "Drop Table [#PHL3Hst]"
RunCQ "Drop Table [#PHL4Cur]"
RunCQ "Drop Table [#PHL4Hst]"

End Sub
