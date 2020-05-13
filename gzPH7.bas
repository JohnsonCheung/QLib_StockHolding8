Attribute VB_Name = "gzPH7"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzPH7."
Public Const PHStmAtr$ = "Stream"
Public Const PHBusAtr$ = "Stream Srt BusArea PHBus"
Public Const PHL1Atr$ = "Stream Srt1 PHL1 PHNam"
Public Const PHL2Atr$ = "Stream Srt2 PHL2 PHNam PHBrd"
Public Const PHL3Atr$ = "Stream Srt3 PHL3 PHNam PHBrd PHQGp"
Public Const PHL4Atr$ = "Stream Srt4 PHL4 PHNam PHBrd PHQGp PHQly"
Public Const PHSkuAtr$ = "Stream Srt4 PHL4 PHNam PHBrd PHQGp PHQly Sku SkuDes"
Public Const PH7ss$ = "Stm Bus L1 L2 L3 L4 Sku"
Function PH7Ay() As String(): PH7Ay = Split(PH7ss): End Function

'Sub T_Reseq(): SrtFldPH7 "@StkHld?", SH10ValFF: End Sub
Sub DrpCPH7Tbl(QmrkStr$)
DrpCTny PH7Tbny(QmrkStr)
End Sub
Function PH7Tbny(PH7TbnQStr$) As String() 'QStr is a string has ?-substr.
PH7Tbny = AmRplQ(PH7Ay, PH7TbnQStr)
End Function

Private Sub Add3PHRollupCol__Tst()
DoCmd.Close acTable, "#A"
RunCQ "Select Sku,SkuDes into [#A] from Sku"
Add3PHRollupCol "#A"
DoCmd.OpenTable "#A"
End Sub

Sub Add3PHRollupCol(SkuT$) 'Add 3-PHRollupCol with value to Tbl-@SkuT from qry-qSku_Main:qry which based on Tb-SKU and more Tbl.  more..
':3PHRollupCol: are 3 columns used to roll up to 6 lvl above of PH: They are {Stm BusArea PHL1..4}
':qSku_Main: is a query based on Tb-Sku with Sku as Pk and Tb-other and with at 3 fields-[Stm BusArea PHL4]
':SkuT: is a tbn with a fld-Sku.
RunCQ "Alter Table [" & SkuT & "] add column Stm Text(1),BusArea Text(4),PHL4 Text(10)"
RunCQ "Update [" & SkuT & "] x inner join qSku_Main a on a.Sku=x.Sku set x.Stm=a.Stm,x.BusArea=a.BusArea,x.PHL4=a.PHL4"
End Sub

Sub SrtFldPH7(PHTbnQStr$, RstFlds$) ' SrtFld for the-7-PH-Tbl by each own 7-PH-Tbl-Sorting-FF, more..
' 7-PH-Tbn come from 7-RplXXX(@TbnQStr)
'       where 7-RplXXX is 7 fun of RplXXX(PHTbnQStr), where XXX is {Stm Bus L1..L4 Sku} and @PHTblnQStr is a str with question mark, which can deduce to a PH-Tbn.
'       eg.  Given @TbnQStr = "@Fc?" ==> 7-PH-Tbn will be @FcStm @FcBus @FcL1.. @FcSku
' A-PH-Tbn of one 7-Tbl with tbn with one of the substr of [Stm Bus L1..4 Sku].  That PH-tbl must have its own set of FF as described in the 7-Pub-Cnst-PH?Atr$
' 7-PH-Tbl-Sorting-FF means 7-FF each for each of 7-PH-Tbl.  They comes from 7 expr of [7-PHAtr & " " & @RstFlds]
'       where 7-PHAtr are 7-pub-str-cnst of name PH?Atr.
'       Given: @RstFlds = "Qty Amt"
'              7-PH?Atr-pub-str-cnst (which is PHStmAtr, ..)
'       Return: 7-PH-Tbl-Sorting-FF
'               #1 = [Stream Qty Amt]
'               #2 = [Stream Srt BusArea PHBus Qty Amt]
'               ..
'               #7 = [Stream Srt4 PHL4 PHNam PHBrd PHQGp PHQly SkuSkuDes Qty Amt]
SrtCFld RplStm(PHTbnQStr), PHStmAtr & " " & RstFlds
SrtCFld RplBus(PHTbnQStr), PHBusAtr & " " & RstFlds
SrtCFld RplL1(PHTbnQStr), PHL1Atr & " " & RstFlds
SrtCFld RplL2(PHTbnQStr), PHL2Atr & " " & RstFlds
SrtCFld RplL3(PHTbnQStr), PHL3Atr & " " & RstFlds
SrtCFld RplL4(PHTbnQStr), PHL4Atr & " " & RstFlds
SrtCFld RplSku(PHTbnQStr), PHSkuAtr & " " & RstFlds
End Sub

Function TmpPH5() ' Create $PH{5} tables, more..
':$PH{5}: :5-Tables ! #PH-Table#
'Oup: $PHL1   = PHL1 | Srt1 PHNam
'Oup: $PHL2   = PHL2 | Srt2 PHNam PHBrd
'Oup: $PHL3   = PHL3 | Srt3 PHNam PHBrd PHQGp
'Oup: $PHL4   = PHL4 | Srt4 PHNam PHBrd PHQGp PHQly
'Oup: $PHSku  = Sku  | Srt4 PHNam PHBrd PHQGp PHQly SkuDes Stream PHL4 | BusArea Stm
'Note: Fst Col is Pk
'      Snd Col is Atr
'      Thd Col is Rollup
StsQry "TmpPH5"
RunCQ "Select PHL1,Srt1,PHNam into [$PHL1] From qPHL1"
RunCQ "Select PHL2,Srt2,PHNam,PHBrd into [$PHL2] From qPHBrd"
RunCQ "Select PHL3,Srt3,PHNam,PHBrd,PHQGp into [$PHL3] From qPHQGp"
RunCQ "Select PHL4,Srt4,PHNam,PHBrd,PHQGp,PHQly into [$PHL4] From qPHQly"
RunCQ "Select Sku, Srt4,PHNam,PHBrd,PHQGp,PHQly,SkuDes,Stream,PHL4,BusArea,Stm into [$PHSku] From qSku"
Sts ""
End Function
