Attribute VB_Name = "gzLoadZHT0"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzLoadZHT0."
Option Base 0
Const HKCusNo& = 5007964
Const MOCusNo& = 5007960

Sub LoadZHT0()
ChkFfnExist ZHT0IFx
If Not Start("Start load ZHT0 tax rate") Then Exit Sub
W1WFx
W1Imp
DrpCT ">ZHT0"
End Sub

Private Sub W1Imp() 'Import ZHT0WFx into Sku..
' ZHT0WFx  -> [>ZHT0] -> #IZHT0 -> (#MO #HK) -> Sku
' #MO where Customer = MOCusNo
' #HK where Customer = HKCusNo

'#IZHT0 from >ZHT0
'#1 Crt #MO #HK             ::[Sku TaxRatePerUom SapUom] where
'#2 Crt #IZHT0 from #MO #HK
'#3 Upd Sku from #IZHT0
'#4   Prompt msg
'#5 Drp #MO #HK #IZHT0
Const T$ = "#IZHT0"

CLnkFxw ZHT0WFx, "Sheet1", ">ZHT0"
W1TmpIZHT0
DoCmd.SetWarnings False

'-- Crt #HK #MO -> Crt #IZHT0
RunCQ "select Material As SKU, CCur(Amount) As TaxRatePerUom, Uom as SapUom" & _
" into `#MO`" & _
" from [#IZHT0] where Customer=" & MOCusNo
RunCQ "select Material As SKU, CCur(Amount) As TaxRatePerUom, Uom as SapUom" & _
" into `#HK`" & _
" from [#IZHT0] where Customer=" & HKCusNo

'-- Upd/Ins tb-Sku by [#IZHT0]
RunCQ "Update Sku set TaxRateHK=Null,TaxRateMO=Null,TaxUOMHK=Null,TaxUOMMO=Null,DteUpdTaxRate=Null"
Dim Tim$: Tim = Format(Now, "YYYY-MM-DD HH:MM:SS")
RunCQ "Update Sku x inner join `#HK` a on x.Sku=a.Sku set x.TaxRateHK=a.TaxRatePerUom,x.TaxUOMHK=a.SapUom, DteUpdTaxRate=#" & Tim & "#"
RunCQ "Update Sku x inner join `#MO` a on x.Sku=a.Sku set x.TaxRateMO=a.TaxRatePerUom,x.TaxUOMMO=a.SapUom, DteUpdTaxRate=#" & Tim & "#"
'-- Show message ========================================================================
Dim NHK%, NMO%
NMO = CNReczT("#MO")
NHK = CNReczT("#HK")
MsgBox "[" & NHK & "] HK Sku" & vbCrLf & "[" & NMO & "] Macau Sku have been imported." & vbCrLf & vbCrLf & "All old ZHT0 rate are totally replaced", vbInformation
DrpCTT "#HK #MO #IZHT0"
End Sub

Private Sub W1TmpIZHT0() ' Crt [>ZHT0] from [#IZHT0].  Dlt rec for columns have invalid value: col[Customer Uom CnTy]..
'#1 Col-Customer must be 5007960 5007964
'#2 Col-CnTy     must be ZHT0
'#3 Col-Uom      must in [COL PCE CA]
DoCmd.SetWarnings False
RunCQ "Select * into [#IZHT0] from [>ZHT0]"

Const WhCnTy$ = "Nz(CnTy,'')<>'ZHT0'"
Const WhUom$ = "Not Uom in ('COL','PCE','CA')"
Const WhCus$ = "Not Customer in (5007960,5007964)"

Dim CntQ$: CntQ = SqlSelCnt_Fm("#IZHT0") & " where "
Dim NCnTy%: NCnTy = VzCQ(CntQ & WhCnTy)
Dim NUom%:  NUom = VzCQ(CntQ & WhUom)
Dim NCus%:  NCus = VzCQ(CntQ & WhCus)

Select Case True
Case NCnTy <> 0, NUom <> 0, NCus <> 0
    MsgBox "There are records with invalid data in fields.  These records are ignored:" & vbCrLf & _
    "CnTy=[" & NCnTy & "]" & vbCrLf & _
    "Uom=[" & NCnTy & "]" & vbCrLf & _
    "Customer=[" & NCus & "]"
    Const DltQ$ = "Delete * from [#IZHT0] where "
    If NCnTy > 0 Then RunCQ DltQ & WhCnTy
    If NUom > 0 Then RunCQ DltQ & WhUom
    If NCus > 0 Then RunCQ DltQ & WhCus
End Select
End Sub

Private Sub W1WFx() ' Crt ZHT0WFx from ZTH0IFx
DltFfnIf ZHT0WFx
Dim Wb As Workbook: Set Wb = NwWbzFx(ZHT0IFx)
Dim Ws As Worksheet: Set Ws = FstWs(Wb)
'--Delete first 4 rows
WsRR(Ws, 6, 7).Delete
WsRR(Ws, 1, 4).Delete
WsC(Ws, "D").Delete
WsC(Ws, "A").Delete
WsC(Ws, "L").Delete
WsRC(Ws, 1, "F").Value = Trim(WsRC(Ws, 1, "F").Value)
WsC(Ws, "D").NumberFormat = "@"                 'format Material as text
'To get ride of the blank columns & lines when using LasCells in CrtLo
Dim Lo As ListObject: Set Lo = CrtLoByWsDta(Ws)
Dim R%(): R = W1Rny(Lo)
Dim J%: For J = UB(R) To 0 Step -1
    Lo.ListRows(R(J)).Range.Delete
Next
Wb.SaveAs ZHT0WFx, XlFileFormat.xlOpenXMLWorkbook
QuitWb Wb
End Sub

Private Function W1Rny(Lo As ListObject) As Integer() ' ret @Rny::Row-no-Ay of @Lo, if col-Customer2 has EMpty value.
Dim Sq(): Sq = Lo.ListColumns("Customer2").DataBodyRange.Value
Dim R%: For R = 1 To UBound(Sq, 1)
    Dim V: V = Sq(R, 1)
    Select Case True
    Case IsEmpty(V), V = "Customer"
        PushI W1Rny, R
    End Select
Next
End Function
