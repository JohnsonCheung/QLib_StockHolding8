Attribute VB_Name = "gzFmtSHTp"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzFmtSHTp."
Private Const SHWsPfx$ = "StkHld "
Private Const FcWsPfx$ = "Fc "
Private Const SdWsPfx$ = "StkDays "
Private Const Wdt_Fc_SC% = 8
Private Const Wdt_Fc_RemSC% = 8
Private Const Wdt_Fc_StkDays% = 5
Private Const Wdt_Fc_M15% = 6

'Ns*******************************************************************************
'Sub Stp(): Stop: End Sub
'Sub T_FmtShTp(): FmtShTp__Tst: End Sub
'Sub T_FmtSh():   FmtSh__Tst:   End Sub
'Sub T_FmtFc():   FmtFc__Tst:   End Sub
'Sub T_FmtSd():   FmtSd__Tst:   End Sub

'Ns*********************************************************************************
Function ShTpWb() As Workbook
Static Wb As Workbook
If Not IsGoodWb(Wb) Then Set Wb = WbzFx(ShTp)
Set ShTpWb = Wb
End Function
Sub FmtShTp()
Dim Wb As Workbook: Set Wb = ShTpWb
Stop
FmtFcDteTit Wb, YM(19, 5)  'At Wb-Lvl  'This is needed to be called each time generating the report
FmtSdDteTit Wb, YM(19, 5)  'At Wb-Lvl 'This is needed to be called each time generating the report
'--- no need to call when gen rpt
Wb.Application.WindowState = xlMinimized
Wb.Application.Visible = True
Stop
FmtKeyCnt Wb                'At Wb-Lvl
FmtSh Wb    'StkHld
FmtFc Wb    'Forecast
FmtSd Wb    'StkDays
SetWbOutLinSum Wb
MiniWbOLvl Wb
SetWbNoAutoColWdt Wb
Wb.Save
Wb.Application.WindowState = xlMaximized
End Sub
'Ns**============================================================================== A_FmtFc
Private Sub FmtFc__Tst(): FmtFc SampShWs: End Sub
Private Sub FmtFc(Wb As Workbook)
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    If IsFc(Ws.Name) Then FmtFcWs Ws
Next
End Sub
Private Sub FmtFcWs(Ws As Worksheet)
Ws.Activate
Dim Lo As ListObject: Set Lo = FstLo(Ws)
SetLcWdt Lo, FcSSM15, Wdt_Fc_M15
SetLcWdt Lo, "SC", Wdt_Fc_SC
SetLcWdt Lo, "StkDays", Wdt_Fc_StkDays
SetLcWdt Lo, "RemSC", Wdt_Fc_RemSC
SetLcAsSum Lo, "SC RemSC"
SetLcFmt Lo, FcSSM15, "#,###"
SetLcFmt Lo, "SC StkDays RemSC", "#,###"
End Sub

Private Sub FmtSd__Tst()
FmtSd SampShWb
End Sub
Private Sub FmtSd(Wb As Workbook)
'#Fmt-StkDays-Ws-es#
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    If IsSd(Ws.Name) Then FmtSdWs Ws
Next
End Sub
Private Sub FmtSdWs(Ws As Worksheet)
Ws.Activate
Dim L As ListObject: Set L = FstLo(Ws)
SetLcWdt L, SdSSStkDays, 6
SetLcWdt L, SdSSRemSc, 6
SetLcWdt L, "SC", 7
SetLcAsSum L, SdSSRemSc
SetLcFmt L, SdSSRemSc, "#,###"
SetLcFmt L, SdSSStkDays, "#,###"
End Sub
'Ns**============================================================================== A_FmtSH
Sub AA_Rpt_SHld_Fmt_Sh(): End Sub
Private Sub FmtSh__Tst(): FmtSh SampShWb: End Sub
Private Sub FmtSh(Wb As Workbook)
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    If IsSh(Ws.Name) Then FmtShWs Ws
Next
End Sub
Private Sub FmtShWs(Ws As Worksheet)
Ws.Activate
Dim Lo As ListObject: Set Lo = FstLo(Ws)
SetLcWdt Lo, "F1 F2", 0.5
SetLcWdt Lo, ShSSHkd, 8
SetLcWdt Lo, ShSSSc, 7
SetLcWdt Lo, ShSSKpi, 5
SetLcAsSum Lo, ShSSSc
SetLcAsSum Lo, ShSSHkd
SetLcFmt Lo, ShSSHkd, "#,###K"
SetLcFmt Lo, ShSSSc, "#,###"
End Sub
'---=========================================================================== FmtKeyCnt
Sub AA_Rpt_SHld_Fmt_AllWs_KeyCnt(): End Sub
Private Sub FmtKeyCnt__Tst()
FmtKeyCnt ShTpWb
End Sub
Private Sub FmtKeyCnt(Wb As Workbook)
'Each SHRptWs using the Wsn to find the Key Column, set this KeyCol.Calc=Cnt
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    If IsShRpt(Ws.Name) Then
        SetLcAsCnt FstLo(Ws), KeyColNm(Ws.Name)
    End If
Next
End Sub
Private Function KeyColNm$(RptWsn$)
':RptWsn: :Nm ! #Rpt-Ws-Nm# must in format of [S3 PH7], where S3 = SktHld | Fc | StkDays & PH7 is one of PH7ss
Dim O$
Select Case AftSpc(RptWsn)
Case "Sku": O = "Sku"
Case "L4": O = "PHQly"
Case "L3": O = "PHQGp"
Case "L2": O = "PHBrd"
Case "L1": O = "PHNam"
Case "Bus": O = "PHBus"
Case "Stm": O = "Stream"
Case Else: Stop
End Select
KeyColNm = O
End Function
'---=========================================================================== A_IsXXXWs
Sub AA_Rpt_SHld_IsXXXWs(): End Sub
Function IsSh(Wsn$) As Boolean: IsSh = HasPfx(Wsn, SHWsPfx): End Function
Function IsFc(Wsn$) As Boolean: IsFc = HasPfx(Wsn, FcWsPfx): End Function
Function IsSd(Wsn$) As Boolean: IsSd = HasPfx(Wsn, SdWsPfx): End Function
Function IsShRpt(Wsn$) As Boolean
Select Case True
Case HasPfx(Wsn, FcWsPfx), HasPfx(Wsn, SdWsPfx), HasPfx(Wsn, SHWsPfx): IsShRpt = True
End Select
End Function

'---=========================================================================== A_SS_Sh
Sub AA_Rpt_SHld_Ws_Sh_SS(): End Sub
Private Function ShSS$(): ShSS = "*Key ScCsg HkdCsg .. ScTot HkdTot StkDays StkMths RemSC TarStkMths": End Function
Private Function ShSSFive$(): ShSSFive = "Csg Df Dp Git Tot": End Function
Private Function ShSSHkd$(): ShSSHkd = ExpandPfxSS("Hkd", ShSSFive): End Function
Private Function ShSSSc$(): ShSSSc = ExpandPfxSS("Sc", ShSSFive): End Function
Private Function ShSSKpi$(): ShSSKpi = "StkDays StkMths RemSC TarStkMths": End Function
'---=========================================================================== A_SS_Fc
Sub AA_Rpt_SHld_Ws_Fc_SS(): End Sub
Private Function FcSS$():       FcSS = "*Key SC StkDays RemSC *M15": End Function
Private Function FcSSM15$(): FcSSM15 = ExpandPfxNN("M", 1, 15, "00"): End Function
'---=========================================================================== A_SS_Sd
Sub AA_Rpt_SHld_Ws_Sd_SS(): End Sub
Private Function SdSS$():               SdSS = "*Key SC *StkDays/RemSC15": End Function
Private Function SdSSRemSc$():     SdSSRemSc = ExpandPfxNN("RemSC", 1, 15, "00"): End Function
Private Function SdSSStkDays$(): SdSSStkDays = ExpandPfxNN("StkDays", 1, 15, "00"): End Function
