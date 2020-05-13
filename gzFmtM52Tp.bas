Attribute VB_Name = "gzFmtM52Tp"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzFmtM52Tp."
Const W_PH4Des% = 20
Const W_Sku% = 8
Const W_SkuDes% = 50
Const W_cv_mlBtl% = 6
Const W_cv_BtlAc% = 4
Const W_cv_LitreSc% = 5
Const W_df_Val% = 8
Const W_df_Val2% = 12

'SetLcWdt Lo, "Val", 12
'FmtLcWdt Lo, "Btl AC SC", 8
'SetLcWdt Lo, "Co", 4
'FmtLcWdt Lo, "YY MM DD", 2
'FmtLcWdt Lo, "Stream", 6
'FmtLcWdt Lo, "PHBus", 2
'FmtLcWdt Lo, "BusArea", 6
'
'FmtLcWdt Lo, "TaxLoc TaxItm", 3
'
'SetLcWdt Lo, "NmYpStk", 20
'FmtLcWdt Lo, "YpStk", 2
'SetLcWdt Lo, "Sku", 8
'FmtLcWdt Lo, "Litre/Btl StkUnit", 7
'FmtLcWdt Lo, "Btl/AC Unit/SC Unit/AC Btl/Unit", 5
'FmtLcWdt Lo, "ml/Btl", 7
'FmtLcWdt Lo, "Btl/AC' Litre/SC Btl Btl/AC Unit/SC Unit/AC Btl/Unit", 5

'Sub T_FmtPt(): FmtBchWsPt__Tst: End Sub

Sub FmtMB52Tpaa()
Dim X As Excel.Application: Set X = NwXls
Dim Wb As Workbook: Set Wb = X.Workbooks.Open(MB52Tp)
Dim Ws As Worksheet, Pt As PivotTable
MaxiWb Wb
Stop
Dim Wsn: For Each Wsn In Termy("Bch Sku Naming Brand BusArea Quality [Quality Group]")
    Set Ws = Wb.Sheets(Wsn)
    If Ws.PivotTables.Count >= 1 Then
        Set Pt = FstPt(Ws)
        FmtPivDfFmt Pt, "[Sum of BchAmt] [Sum of ZHT0Amt]", "$#,##0;-$#,##0;#"
        FmtPivDfFmt Pt, "[Average of RatDif] [Sum of AmtDif]", "$#,##0;-$#,##0;#"
    End If
Next
Wb.Save
MaxiWb Wb
Done
End Sub
Sub FmtMB52Tp()
Dim X As Excel.Application: Set X = NwXls
Dim Wb As Workbook: Set Wb = X.Workbooks.Open(MB52Tp)
MaxiWb Wb
FmtDtaWs Wb.Sheets("Data")
'FmtPt Wb
'SetWbOutLinSum Wb
'MiniWbOLvl Wb
'SetWbNoAutoColWdt Wb
Wb.Save
MaxiWb Wb
Done
End Sub
Private Sub FmtDtaWs(Ws As Worksheet)
Ws.Activate
Dim Lo As ListObject: Set Lo = FstLo(Ws)
FmtFmt Lo
'FmtTot Lo
'FmtWdt Lo
'FmtFml Lo
'FmtLvl Lo
End Sub
Private Sub FmtFml(Lo As ListObject)
SetLoFml Lo, MB52DtaWsFml
End Sub
Private Sub FmtFmt(Lo As ListObject)
'== Ws-Data: Format Column ===
SetLcFmt Lo, "Litre/Btl", "0.000;-0.000;#"
SetLcFmt Lo, "Btl/AC", "#"
SetLcFmt Lo, "Unit/AC", "#"
SetLcFmt Lo, "Unit/SC", "0.00"
SetLcFmt Lo, "Btl/Unit", "#"

'--Siz
SetLcFmt Lo, "ml/Btl", "#,###"
SetLcFmt Lo, "Btl/AC'", "#"
SetLcFmt Lo, "Litre/SC", "#,##0.0"
'
SetLcFmt Lo, "Val", "$#,###"
SetLcFmt Lo, "Btl", "#,###"
SetLcFmt Lo, "SC", "#,##0.0"
SetLcFmt Lo, "AC", "#.##0.0"
SetLcFmt Lo, "ScUPr", "$#,###"
SetLcFmt Lo, "AcUPr", "$#,###"
SetLcFmt Lo, "BtlUPr", "$#,###"
SetLcFmt Lo, "BchRat", "$#,###;-$#,###;"""""
SetLcFmt Lo, "ZHT0Rat", "$#,###;-$#,###;"""""
SetLcFmt Lo, "BchAmt", "$#,###;-$#,###;"""""
SetLcFmt Lo, "ZHT0Amt", "$#,###;-$#,###;"""""

SetLcFmt Lo, "RatDif", "$#,###;-$#,###;"""""
SetLcFmt Lo, "AmtDif", "$#,###;-$#,###;"""""
End Sub
Private Sub FmtTot(Lo As ListObject)
SetLccAsSum Lo, "Val Btl AC SC"
SetLccAsAvg Lo, "ScUPr BtlUPr AcUPr"
SetLccAsAvg Lo, "BchRat ZHT0Rat"
SetLccAsSum Lo, "BchAmt ZHT0Amt"
SetLccAsSum Lo, "AmtDif"
SetLcAsCnt Lo, "SkuDes"
End Sub
Private Sub SetWdt(L As ListObject)
 SetLcWdt L, "Val", 12
SetLccWdt L, "Btl AC SC", 8
 SetLcWdt L, "Co", 4
SetLccWdt L, "YY MM DD", 2
 SetLcWdt L, "Stream", 6
 SetLcWdt L, "PHBus", 2
 SetLcWdt L, "BusArea", 6

SetLccWdt L, "TaxLc TaxItm", 3

 SetLcWdt L, "NmYpStk", 20
 SetLcWdt L, "YpStk", 2
 SetLcWdt L, "SkuDes", 50
 SetLcWdt L, "Sku", 8
SetLccWdt L, "Litre/Btl StkUnit", 7
SetLccWdt L, "Btl/AC Unit/SC Unit/AC Btl/Unit", 5
 SetLcWdt L, "ml/Btl", 7
SetLccWdt L, "Btl/AC' Litre/SC Btl Btl/AC Unit/SC Unit/AC Btl/Unit", 5
End Sub
Private Sub FmtLvl(L As ListObject)
SetLcLvl L, "YY MM DD"
SetLcLvl L, "PHBus BusArea NmYpStk YpStk"
SetLcLvl L, "Sku Litre/Btl Btl/AC  Unit/SC Unit/AC StkUnit"
SetLcLvl L, "BchAmt  ZHT0Amt RatDif  AmtDif  BchNo   BchRatTy"
SetLcLvl L, "PH  PHNam   PHBrd   PHQGp   PHQly   PHSStm  PHSBus  PHSrt1  PHSrt2  PHSrt3  PHSrt4  PHL1    PHL2    PHL3    PHL4"
End Sub

Private Sub FmtPt(Wb As Workbook)
FmtBchWsPt Wb.Sheets("Bch1")
Dim I: For Each I In PH7Ay
    FmtPtWs I
Next
End Sub
Private Sub FmtBchWsPt__Tst()
Dim X As New Excel.Application
MaxiXls X
Dim Wb As Workbook: Set Wb = X.Workbooks.Open(MB52Tp)
Dim Ws As Worksheet: Set Ws = Wb.Sheets("Bch1")
FmtBchWsPt Ws
Wb.Save
Done
End Sub
Private Sub FmtBchWsPt(Ws As Worksheet)
Dim Pt As PivotTable: Set Pt = FstPt(Ws)
FmtBchWsPtWdt Pt
FmtBchWsPtFmt Pt
End Sub

Private Sub FmtBchWsPtWdt(Pt As PivotTable)
FmtPivRfWdt Pt, "PHNam PHBrd PHQGp PHQly", W_PH4Des
FmtPivRfWdt Pt, "Sku", W_Sku
FmtPivRfWdt Pt, "SkuDes", W_SkuDes
FmtPivRfWdt Pt, "ml/Btl", W_cv_mlBtl
FmtPivRfWdt Pt, "Btl/AC", W_cv_BtlAc
FmtPivRfWdt Pt, "Litre/SC", W_cv_LitreSc
FmtPivDfWdt Pt, "[Sum of Val] [Sum of AC] [Sum of SC] [Sum of Btl] [Average of BtlUPr] [Average of AcUPr] [Average of ScUPr]" & _
" [Average of BchRat] [Average of RatDif] [Average of ZHT0Rat]", W_df_Val
FmtPivDfWdt Pt, "[Sum of BchAmt] [Sum of ZHT0Amt] [Sum of AmtDif]", W_df_Val2
End Sub
Private Sub FmtBchWsPtFmt(Pt As PivotTable)
'D PtDtaFny(Pt)

'4: Val+3Q
'3: Price
'2: Rate
'2: Amt
'1: RatDif
'1: AmtDif
FmtPivDfFmt Pt, "[Sum Of Val]", "$#,##0"
FmtPivDfFmt Pt, "[Sum of AC]", "#.##0"
FmtPivDfFmt Pt, "[Sum of SC]", "#,##0.0"
FmtPivDfFmt Pt, "[Sum of Btl]", "#,##0"

FmtPivDfFmt Pt, "[Average of AcUPr] [Average of BtlUPr] [Average of ScUPr]", "$#,##0"

FmtPivDfFmt Pt, "[Average of BchRat] [Average of ZHT0Rat]", "$#,##0"
FmtPivDfFmt Pt, "[Sum of BchAmt] [Sum of ZHT0Amt]", "$#,##0"

FmtPivDfFmt Pt, "[Average of RatDif] [Sum of AmtDif]", "$#,##0"
End Sub
Sub FmtPivDfFmt(Pt As PivotTable, DtaFF$, Fmt$)
Dim F: For Each F In Termy(DtaFF)
    SetPivDfFmt Pt, F, Fmt
Next
End Sub
Sub SetPivDfFmt(Pt As PivotTable, F, Fmt$)
Dim Pf As PivotField: Set Pf = Pt.DataFields(F)
Pf.NumberFormat = Fmt
End Sub
Function PtFny(Pt As PivotTable) As String()
Dim F As PivotField: For Each F In Pt.PivotFields
    PushS PtFny, F.Name
Next
End Function
Function PtDtaFny(Pt As PivotTable) As String()
Dim F As PivotField: For Each F In Pt.DataFields
    PushS PtDtaFny, F.Name
Next
End Function
Sub FmtPivDfWdt(Pt As PivotTable, FF$, W)
Dim C: For Each C In Termy(FF)
    SetPivDfWdt Pt, C, W
Next
End Sub
Sub SetPivDfWdt(Pt As PivotTable, F, W)
EntPivDfCol(Pt, F).ColumnWidth = W
End Sub
Function EntPivDfCol(Pt As PivotTable, F) As Range
Dim Pf As PivotField: Set Pf = Pt.DataFields(F)
Set EntPivDfCol = Pf.DataRange.EntireColumn
End Function
Sub FmtPivRfWdt(Pt As PivotTable, RowFF$, W)
Dim C: For Each C In Termy(RowFF)
    SetPivRfWdt Pt, C, W
Next
End Sub
Sub SetPivRfWdt(Pt As PivotTable, Rf, W)
EntPivRfCol(Pt, Rf).ColumnWidth = W
End Sub

Function EntPivRfCol(Pt As PivotTable, PivRfn) As Range
Dim F As PivotField: Set F = Pt.PivotFields(PivRfn)
If F.Orientation <> xlRowField Then PmEr "EntPivRfCol", "PivRfhn", PivRfn, "xlRowField", , "PivRfh.Orientation is not a row"
Set EntPivRfCol = F.DataRange.EntireColumn
End Function
Private Sub CrtPt()
'
'Const AtAdr$ = "Bch1!A5"
'Dim Pt As PivotTable
'Dim Pc As PivotCache
'RmvPt BchWb
'    set Pt =    WbzWs(BchWs).PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
'        "Data", Version:=6).CreatePivotTable TableDestination:=AtAdr, TableName:="PivotTable3", _
'        DefaultVersion:=6
'    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Stream")
'        .Orientation = xlPageField
'        .Position = 1
'    End With
'    With ActiveSheet.PivotTables("PivotTable3").PivotFields("PHBus")
'        .Orientation = xlRowField
'        .Position = 1
'    End With
'    With ActiveSheet.PivotTables("PivotTable3").PivotFields("BusArea")
'        .Orientation = xlRowField
'        .Position = 2
'    End With
'    With ActiveSheet.PivotTables("PivotTable3").PivotFields("NmYpStk")
'        .Orientation = xlColumnField
'        .Position = 1
'    End With
'    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
'        "PivotTable3").PivotFields("Val"), "Sum of Val", xlSum
'    With ActiveSheet.PivotTables("PivotTable3")
'        .InGridDropZones = True
'        .RowAxisLayout xlTabularRow
'    End With
'    Range("A8").Select
'    ActiveSheet.PivotTables("PivotTable3").PivotFields("PHBus").Subtotals = Array( _
'        False, False, False, False, False, False, False, False, False, False, False, False)
'End Sub
'
End Sub
Function CvPts(A) As PivotTable
Set CvPts = A
End Function

Sub RmvPt(Ws As Worksheet)
Dim Pt As PivotTable: For Each Pt In Ws.PivotTables
    Pt.TableRange2.ClearContents
Next
End Sub
Private Sub FmtPtWs(PHItm)
Dim Wsn$
End Sub

'==
Function MB52TpWb() As Workbook
Set MB52TpWb = TpWb
End Function
Private Function TpWb() As Workbook
Set TpWb = NwXls.Workbooks.Open(MB52Tp)
End Function
