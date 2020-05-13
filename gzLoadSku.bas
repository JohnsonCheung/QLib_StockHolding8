Attribute VB_Name = "gzLoadSku"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzLoadSku."
Option Base 0
Private Type ErSts
    MsgMat() As String
    MsgSKuDes() As String
    MsgTopaz() As String
    MsgPH() As String
    MsgStkUnit() As String
    MsgStkUnit_ValEr() As String 'Must be PCE COL
    MsgUnit_per_AC() As String
    MsgUnit_per_SC() As String
    MsgLitre_per_Btl() As String
    MsgBusArea() As String
End Type
Const LvcNmf = "Material,Sales text language 1,Topaz Code,Product hierarchy,Base Unit of Measure,Unit per case,SC,COL Per Case,Bottle Capacity,Business Area"

Const ExtNm_BusArea$ = "Business Area"
Const ExtNm_Liter_per_Btl$ = "Bottle Capacity"
Const ExtNm_Mat$ = "Material"
Const ExtNm_PH$ = "Product hierarchy"
Const ExtNm_SkuDes$ = "Sales text language 1"
Const ExtNm_StkUnit$ = "Base Unit of Measure"
Const ExtNm_Topaz$ = "Topaz Code"
Const ExtNm_Btl_per_AC$ = "COL Per Case"
Const ExtNm_Unit_per_AC$ = "COL Per Case"
Const ExtNm_Unit_per_SC$ = "SC"
Const IntNm_BusArea_Sap$ = "BusAreaSap"
Const IntNm_Liter_per_Btl_Sap$ = "Litre/BtlSap"
Const IntNm_Mat$ = "Sku"
Const IntNm_PH$ = "ProdHierarchy"
Const IntNm_SkuDes$ = "SkuDes"
Const IntNm_StkUnit$ = "StkUnit"
Const IntNm_Topaz$ = "CdTopaz"
Const IntNm_Btl_per_AC$ = "Btl/AC"
Const IntNm_Unit_per_AC$ = "Unit/AC"
Const IntNm_Unit_per_SC$ = "Unit/SC"


Sub LoadSku()
DoCmd.SetWarnings False
Dim Tim As Date: Tim = Now
R1ChkColType
R2WarnNullMaterial
R3Tmp8687
R4ChkErVal
R5InsTopaz
R6Upd Tim
R7Ins Tim
R8PromptRslt Tim
RfhTbPHLBus_FmSku_NewBusArea
RfhTbSku_Ovr
DrpCTT "#86Sku #87Sku >Sku86 >Sku87"
End Sub

Private Sub R1ChkColType()
'-- Stp-Lnk-and-Chk-InpFx
Dim IFx$: IFx = SkuPrm_InpFx
Const FldNmCsv$ = "Material,Sales text language 1,Topaz Code,Product hierarchy,Base Unit of Measure,Unit per case,SC,COL Per Case,Bottle Capacity,Business Area"
Const TyChrCsv$ = "T       ,T                    ,T         ,T                ,T                   ,TorN         ,N ,N           ,N              ,T"
ChkFxww IFx, "8601 8701"
ChkWsCol IFx, "8601", FldNmCsv, TyChrCsv
ChkWsCol IFx, "8701", FldNmCsv, TyChrCsv
End Sub

Private Sub R2WarnNullMaterial()
Dim N86&: N86 = VzCQ("Select Count(*) from [>Sku86] where Trim(Nz(Material,''))=''")
Dim N87&: N87 = VzCQ("Select Count(*) from [>Sku87] where Trim(Nz(Material,''))=''")
If N86 > 0 Or N87 > 0 Then
    MsgBox _
    "There are [" & N86 & "] lines in 8601 worksheet with empty [Material]" & vbCrLf & _
    "There are [" & N87 & "] lines in 8701 worksheet with empty [Material]" & vbCrLf & vbCrLf & _
    "These lines are ignored!" & vbCrLf & vbCrLf & _
    "[Ok]=Continue", vbInformation
End If
End Sub
Private Sub R3Tmp8687()
'-------------------
Sts "Running import query ....."
'== Stp-Crt-#86Sku-#87Sku
'Fm : >Sku86 / >Sku87
Const SelIFx$ = "Select " & _
"Trim(Nz(Material               ,'')) AS SKU," & _
"Trim(Nz(`Sales text language 1`,'')) as SkuDes," & _
"Trim(Nz([Topaz Code]           ,'')) as CdTopaz," & _
"Trim(Nz([Product hierarchy]    ,'')) as ProdHierarchy," & _
"Trim(Nz([Base Unit of Measure] ,'')) as StkUnit," & _
"Val(Nz([Unit per case]          ,0)) as [Unit/Ac]," & _
"Val(Nz([SC]                     ,0)) as [Unit/SC]," & _
"Val(Nz([COL Per Case]           ,0)) As [Btl/AC]," & _
"Val(Nz([Bottle Capacity]        ,0)) As [Litre/BtlSap]," & _
"Trim(Nz([Business Area]         ,0)) As BusAreaSap," & _
"CLng(0)                 As Topaz"
RunCQ SelIFx & " into [#86Sku] from [>Sku86] where Trim(Nz(Material,''))<>''"
RunCQ SelIFx & " into [#87Sku] from [>Sku87] where Trim(Nz(Material,''))<>''"
End Sub
Private Sub R4ChkErVal__Tst()
R3Tmp8687
R4ChkErVal
End Sub
Private Sub R4ChkErVal()
If MsgBox("Do data checking?", vbYesNo) = vbNo Then Exit Sub
Dim O$(): O = AddSy(ErLy(86), ErLy(87))
If Si(O) = 0 Then Exit Sub
BrwAy O
If MsgBox("There are errors in the Sales Text Excel file." & vbCrLf & vbCrLf & "[Ok]=Continue loading Sales Text with some data missing...!" & vbCrLf & "[Cancel]=Cancel", vbQuestion + vbOKCancel) = vbOK Then Exit Sub
Raise "Canceled"
End Sub
Private Sub R5InsTopaz()
'== Stp-InsTbl_Topaz
RunCQ "Select Distinct CdTopaz into `#A` from `#86Sku`"
RunCQ "Select Distinct CdTopaz into `#B` from `#87Sku`"
RunCQ "Insert into Topaz Select x.CdTopaz from `#A` x left Join Topaz a on x.CdTopaz=a.CdTopaz where a.CdTopaz is null"
RunCQ "Insert into Topaz Select x.CdTopaz from `#B` x left Join Topaz a on x.CdTopaz=a.CdTopaz where a.CdTopaz is null"
RunCQ "Drop Table `#A`"
RunCQ "Drop Table `#B`"

'== Stp-UpdTmp8687-Fld-Topaz
RunCQ "Update `#86Sku` x inner join Topaz a on x.CdTopaz=a.CdTopaz set x.Topaz=a.Topaz"
RunCQ "Update `#87Sku` x inner join Topaz a on x.CdTopaz=a.CdTopaz set x.Topaz=a.Topaz"
End Sub
Private Sub TstDif()
Const Er1$ = "Unit/SC"
Const A1$ = "Topaz"
Const A2$ = "ProdHierarchy"
Const A3$ = "SkuDes"
Const A4$ = "StkUnit"
Const A5$ = "Unit/AC"
Const A6$ = "Unit/SC"
Const A7$ = "Btl/AC"
Const A8$ = "Litre/BtlSap"
Const A9$ = "BusAreaSap"
Const A10$ = "SkuDes"
Dim A$
A = A6
MsgBox _
VzCQ("Select Count(*) from SKU x inner join `#86Sku` a on x.SKU=a.SKU " & WhNewSku(A)) & " " & _
VzCQ("Select Count(*) from SKU x inner join `#87Sku` a on x.SKU=a.SKU " & WhNewSku(A)) & " "

'Brw_Sql "Select x.SKu,x.[Unit/SC],a.[Unit/SC] from SKU x inner join `#87Sku` a on x.SKU=a.SKU where a.[Unit/SC]<>x.[Unit/SC]"
'Brw_Sql "Select x.SKu,x.SkuDes,a.SkuDes from SKU x inner join `#87Sku` a on x.SKU=a.SKU where a.SkuDes<>x.SkuDes"
End Sub
Private Function WhNewSku$(F$)
WhNewSku$ = " where x.[" & F & "]<>a.[" & F & "]"
End Function

Private Sub R6Upd(Tim As Date)
Const WhNewSku$ = _
"    x.Topaz        <>a.Topaz" & _
" or x.ProdHierarchy<>a.ProdHierarchy" & _
" or x.SkuDes     <>a.SkuDes" & _
" or x.StkUnit    <>a.StkUnit" & _
" or x.[Unit/AC]  <>a.[Unit/AC]" & _
" or x.[Unit/SC]  <>a.[Unit/SC]" & _
" or x.[Btl/AC]   <>a.[Btl/AC]" & _
" or x.[Litre/BtlSap]<>a.[Litre/BtlSap]" & _
" or x.BusAreaSap    <>a.BusAreaSap"

Const SetEq$ = _
"x.Topaz=a.Topaz," & _
"x.ProdHierarchy=a.ProdHierarchy," & _
"x.SkuDes=a.SkuDes," & _
"x.StkUnit=a.StkUnit," & _
"x.[Unit/Ac]=a.[Unit/Ac]," & _
"x.[Unit/Sc]=a.[Unit/Sc]," & _
"x.[Btl/Ac]=a.[Btl/Ac]," & _
"x.[Litre/BtlSap]=a.[Litre/BtlSap]," & _
"x.BusAreaSap=a.BusAreaSap,"

RunCQ "Select * into [#Mge] from [#86Sku]"
RunCQ "Insert into [#Mge] Select x.* from [#87Sku] x left join [#86Sku] a on a.Sku=x.Sku where a.Sku is null"
RunCQ FmtQQ("Update SKU x inner join [#Mge] a on x.SKU=a.SKU set ? x.DteUpdTopaz=#?# where ?", SetEq, Tim, WhNewSku)
DrpCT "#Mge"
End Sub

Private Sub R7Ins(Tim As Date)
Const InsSku$ = "Insert Into Sku (Sku,Topaz,SkuDes,ProdHierarchy,StkUnit,[Unit/SC],[Unit/AC],[Litre/BtlSap],[Btl/AC],BusAreaSap,DteCrt)"
Dim SelInp$: SelInp = "Select X.Sku, X.Topaz," & _
"Trim(Nz(X.SkuDes       ,'')) as SkuDes," & _
"Trim(Nz(X.ProdHierarchy,'')) as ProdHierarchy," & _
"Trim(Nz(X.StkUnit      ,'')) as StkUnit," & _
"Nz(X.[Unit/Sc]     ,0)  As [Unit/SC]," & _
"Nz(X.[Unit/Ac]     ,0)  As [Unit/AC]," & _
"Nz(X.[Litre/BtlSap],0)  As [Litre/BtlSap]," & _
"Nz(X.[Btl/Ac]      ,0)  As [Btl/AC]," & _
"Nz(X.[BusAreaSap]  ,'') As BusAreaSap," & _
"#" & Tim & "# As DteCrt"
RunCQ InsSku & SelInp & " from `#86Sku` x left join SKU a on x.SKU=a.SKU where a.SKU is null"
RunCQ InsSku & SelInp & " from `#87Sku` x left join SKU a on x.SKU=a.SKU where a.SKU is null"
End Sub

Private Sub RfhTbPHLBus_FmSku_NewBusArea()
'Aim: For those Sku->BusAreaSap is not found in PHLBus, add a new record to [PHLBus] and given notePad message.
'     It is called by Sku_Load subr
'     PHLBus: BusArea (4chr) PHBus (Des
RunCQ "Select Distinct BusArea into [#BusArea] from Sku where Nz(BusArea,'')<>''"
RunCQ "Select x.BusArea into [#BusAreaNew] from [#BusArea] x left join PHLBus a on x.BusArea=a.BusArea where a.BusArea is null"
RunCQ "Insert into PHLBus (BusArea, PHBus) Select BusArea,BusArea & ' Des' as PHBus from [#BusAreaNew]"
Dim N%: N = CNReczT("#BusAreaNew")
If N > 0 Then
    MsgBox "There are [" & N & "] new business area are found, please go enter their description", vbInformation
End If
DrpCTT "#BusArea #BusAreaNew"
End Sub
Private Sub R8PromptRslt__Tst()
R8PromptRslt Now
End Sub
Private Sub R8PromptRslt(Tim As Date)
'== Stp-Shw-N-New&Chg
Dim N%: N = NNew(Tim)
Dim C%: C = NChg(Tim)
Dim Tail$: If N <> 0 Or C <> 0 Then Tail = vbCrLf & "Check time Stamp [" & Tim & "]"
MsgBox "Done" & vbCrLf & vbCrLf & "There are [" & N & "] Sku created" & vbCrLf & "There are [" & C & "] Sku changed." & Tail, vbInformation
'== Stp-Rfh-Frm
If IsFrmOpn("Mst_Imp_Topaz") Then Forms("Mst_Imp_Topaz").Requery
End Sub

Private Function ErLy(Co As Byte) As String()
ErLy = ErzErSts(ErSts(Co), Co)
End Function
Private Function IsEmpErSts(A As ErSts) As Boolean
With A
Select Case True
Case _
Si(.MsgBusArea) <> 0, _
Si(.MsgLitre_per_Btl) <> 0, _
Si(.MsgMat) <> 0, _
Si(.MsgPH) <> 0, _
Si(.MsgSKuDes) <> 0, _
Si(.MsgStkUnit) <> 0, _
Si(.MsgStkUnit_ValEr) <> 0, _
Si(.MsgTopaz) <> 0, _
Si(.MsgUnit_per_AC) <> 0, _
Si(.MsgUnit_per_SC) <> 0
Exit Function
End Select
End With
IsEmpErSts = True
End Function
Private Function ErzErSts(A As ErSts, Co As Byte) As String()
If IsEmpErSts(A) Then Exit Function
Dim O$()
PushI O, "There are errors in the Sales Text Excel files:"
PushI O, "==============================================="
PushI O, "Excel Files: [" & SkuPrm_InpFx & "]"
PushI O, "Worksheet  : [" & WsnzCo(Co) & "]"
With A
    PushAy O, BlnkFldMsg(.MsgBusArea, ExtNm_BusArea)
    PushAy O, ZeroNegFldMsg(.MsgLitre_per_Btl, ExtNm_Liter_per_Btl)
    PushAy O, BlnkFldMsg(.MsgMat, ExtNm_Mat)
    PushAy O, BlnkFldMsg(.MsgPH, ExtNm_PH)
    PushAy O, BlnkFldMsg(.MsgSKuDes, ExtNm_SkuDes)
    PushAy O, BlnkFldMsg(.MsgStkUnit, ExtNm_StkUnit)
    PushAy O, StkUnitValErMsg(.MsgStkUnit_ValEr, ExtNm_StkUnit)
    PushAy O, BlnkFldMsg(.MsgTopaz, ExtNm_Topaz)
    PushAy O, ZeroNegFldMsg(.MsgUnit_per_AC, ExtNm_Unit_per_AC)
    PushAy O, ZeroNegFldMsg(.MsgUnit_per_SC, ExtNm_Unit_per_SC)
    PushI O, ""
End With
ErzErSts = O
End Function
Private Function StkUnitValErMsg(Msg$(), ExtNm$) As String()
If Si(Msg) = 0 Then Exit Function
PushS StkUnitValErMsg, "Column[" & ExtNm & "] has [" & Si(Msg) & "] lines has invalid value: Valid value should be COL PCE:"
PushAy StkUnitValErMsg, AmAddPfxTab(Msg)
End Function
Private Function WsnzCo$(Co As Byte): WsnzCo = Co & "01": End Function

Private Function BlnkFldMsg(Msg$(), ExtNm$) As String()
If Si(Msg) = 0 Then Exit Function
PushS BlnkFldMsg, "Column[" & ExtNm & "] has [" & Si(Msg) & "] lines blank value:"
PushAy BlnkFldMsg, AmAddPfxTab(Msg)
End Function
Private Function ZeroNegFldMsg(Msg$(), ExtNm$) As String()
If Si(Msg) = 0 Then Exit Function
PushS ZeroNegFldMsg, "Column[" & ExtNm & "] has [" & Si(Msg) & "] lines with no positive numeric value"
PushAy ZeroNegFldMsg, AmAddPfxTab(Msg)
End Function
Private Function ErSts(Co As Byte) As ErSts
Dim T$: T = "#" & Co
With ErSts
    .MsgTopaz = MsgzBlnkVal(Co, IntNm_Topaz)
    .MsgBusArea = MsgzBlnkVal(Co, IntNm_BusArea_Sap)
    .MsgLitre_per_Btl = MsgYYeroNeg(Co, IntNm_Liter_per_Btl_Sap)
    .MsgMat = MsgzBlnkVal(Co, IntNm_Mat)
    .MsgPH = MsgzBlnkVal(Co, IntNm_PH)
    .MsgSKuDes = MsgzBlnkVal(Co, IntNm_Mat)
    .MsgStkUnit = MsgzBlnkVal(Co, IntNm_StkUnit)
    .MsgTopaz = MsgzBlnkVal(Co, IntNm_Topaz)
    .MsgUnit_per_AC = MsgYYeroNeg(Co, IntNm_Unit_per_AC)
    .MsgUnit_per_SC = MsgYYeroNeg(Co, IntNm_Unit_per_SC)
    .MsgStkUnit_ValEr = ErSkuzSql("Select Sku,SkuDes from [#" & Co & "Sku] where Not StkUnit in ('COL','PCE')")
End With
End Function
Private Sub MsgzBlnkVal__Tst()
'Debug.Print MsgzBlnkVal(87, IntNm_BusArea)
DmpAy MsgzBlnkVal(86, IntNm_Topaz)
End Sub
Private Function MsgzBlnkVal(Co As Byte, IntNm$) As String()
MsgzBlnkVal = ErSkuzSql(sqlzBlnkVal(Co, IntNm))
End Function
Private Function MsgYYeroNeg(Co As Byte, IntNm$) As String()
MsgYYeroNeg = ErSkuzSql(sqlYYeroNeg(Co, IntNm))
End Function
Private Function ErSkuzSql(Sql_of_Sku_and_SkuDes$) As String()
With CurrentDb.OpenRecordset(Sql_of_Sku_and_SkuDes)
    While Not .EOF
        PushI ErSkuzSql, "Sku[" & !SKU & "] Des[" & !SkuDes & "]"
        .MoveNext
    Wend
End With
End Function
Private Function sqlzBlnkVal$(Co As Byte, F$)
Const C$ = "Select SKu,SkuDes from [#?Sku] where Trim(Nz([?],''))=''"
sqlzBlnkVal = FmtQQ(C, Co, F)
End Function
Private Function sqlYYeroNeg$(Co As Byte, F$)
Const C$ = "Select Sku,SkuDes from [#?Sku] where Nz([?],0)<=0"
sqlYYeroNeg = FmtQQ(C, Co, F)
End Function

Private Function CntChgSkuSql$(Tim As Date)
CntChgSkuSql = "Select Count(*) from Sku where DteUpdTopaz=#" & Tim & "#"
End Function
Private Function CntNewSkuSql$(Tim As Date)
CntNewSkuSql = "Select Count(*) from Sku where DteCrt=#" & Tim & "#"
End Function

Function NNew%(Tim As Date): NNew = VzCQ(CntNewSkuSql(Tim)): End Function
Function NChg%(Tim As Date): NChg = VzCQ(CntChgSkuSql(Tim)): End Function
